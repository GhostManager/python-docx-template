# -*- coding: utf-8 -*-
"""
Created : 2021-07-30

@author: Eric Lapouyade
"""
from xml.sax.saxutils import escape as xml_escape

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.oxml.shape import CT_Inline
from docx.shared import Emu


def _get_single_xpath(element, xpath, description):
    matches = element.xpath(xpath)
    if len(matches) != 1:
        raise RuntimeError(
            "python-docx generated inline image XML is incompatible with "
            "docxtpl's fast inline image template: expected exactly one "
            "%s at %s, found %d." % (description, xpath, len(matches))
        )
    return matches[0]


def _build_inline_image_xml_template():
    """Generate the XML format string by calling python-docx once.

    This ensures the template always matches the installed python-docx version's
    XML structure, even after upgrades. We create one inline image element with
    valid values, then replace the exact XML attributes with Python format
    placeholders before serializing it.
    """
    inline = CT_Inline.new_pic_inline(
        1,
        "rId",
        "filename",
        Emu(1),
        Emu(1),
    )

    extent = _get_single_xpath(inline, "./wp:extent", "drawing extent")
    doc_pr = _get_single_xpath(inline, "./wp:docPr", "drawing properties")
    c_nv_pr = _get_single_xpath(inline, ".//pic:cNvPr", "picture properties")
    blip = _get_single_xpath(inline, ".//a:blip", "image relationship")
    shape_extent = _get_single_xpath(inline, ".//a:ext", "picture extent")

    extent.set("cx", "{cx}")
    extent.set("cy", "{cy}")
    doc_pr.set("id", "{shape_id}")
    doc_pr.set("name", "Picture {shape_id}")
    c_nv_pr.set("name", "{filename}")
    blip.set(qn("r:embed"), "{rId}")
    shape_extent.set("cx", "{cx}")
    shape_extent.set("cy", "{cy}")

    return inline.xml


# Pre-built XML template for inline images, derived from the installed
# python-docx version. Using str.format() on this template avoids calling
# CT_Inline.new_pic_inline() per image (which does 2x parse_xml() +
# element manipulation + .xml serialization each time).
_INLINE_IMAGE_XML = None


def _get_inline_image_xml_template():
    global _INLINE_IMAGE_XML
    if _INLINE_IMAGE_XML is None:
        _INLINE_IMAGE_XML = _build_inline_image_xml_template()
    return _INLINE_IMAGE_XML


def _format_inline_image_xml(shape_id, rId, filename, cx, cy):
    try:
        template = _get_inline_image_xml_template()
    except RuntimeError:
        return CT_Inline.new_pic_inline(
            shape_id,
            rId,
            filename or "",
            Emu(int(cx)),
            Emu(int(cy)),
        ).xml

    return template.format(
        cx=int(cx),
        cy=int(cy),
        shape_id=shape_id,
        filename=xml_escape(filename or "", {'"': "&quot;"}),
        rId=rId,
    )


class InlineImage(object):
    """Class to generate an inline image

    This is much faster than using Subdoc class.
    """

    tpl = None
    image_descriptor = None
    width = None
    height = None
    anchor = None

    def __init__(self, tpl, image_descriptor, width=None, height=None, anchor=None):
        self.tpl, self.image_descriptor = tpl, image_descriptor
        self.width, self.height = width, height
        self.anchor = anchor

    def _add_hyperlink(self, run, url, part):
        # Create a relationship for the hyperlink
        r_id = part.relate_to(
            url,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            is_external=True,
        )

        # Find the <wp:docPr> and <pic:cNvPr> element
        docPr = run.xpath(".//wp:docPr")[0]
        cNvPr = run.xpath(".//pic:cNvPr")[0]

        # Create the <a:hlinkClick> element
        hlinkClick1 = OxmlElement("a:hlinkClick")
        hlinkClick1.set(qn("r:id"), r_id)
        hlinkClick2 = OxmlElement("a:hlinkClick")
        hlinkClick2.set(qn("r:id"), r_id)

        # Insert the <a:hlinkClick> element right after the <wp:docPr> element
        docPr.append(hlinkClick1)
        cNvPr.append(hlinkClick2)

        return run

    def _insert_image(self):
        part = self.tpl.current_rendering_part
        image_descriptor = self.image_descriptor

        # Cache the expensive parts (image part lookup, rId, dimensions) per
        # (part, descriptor, width, height).  The XML string itself is NOT
        # cached because each insertion needs a unique shape_id - header/footer
        # and footnote parts are not renumbered by fix_docpr_ids().
        cache = self.tpl._image_cache
        # For hashable, value-stable descriptors (strings, paths), cache by
        # value. File-like objects are mutable even when hashable (BytesIO,
        # open file handles), so never cache their image metadata.
        try:
            if hasattr(image_descriptor, "read"):
                raise TypeError
            cache_key = (id(part), image_descriptor, self.width, self.height)
            hash(cache_key) is not None  # trigger TypeError if unhashable
        except TypeError:
            cache_key = None

        if cache_key is not None and cache_key in cache:
            rId, cx, cy, filename = cache[cache_key]
        else:
            # Get or add the image part with O(1) descriptor-based dedup,
            # avoiding the O(n) linear scan in python-docx's default path.
            image_part, image = self.tpl._get_or_add_image_part(image_descriptor)
            rId = part.relate_to(image_part, RT.IMAGE)
            cx, cy = image.scaled_dimensions(self.width, self.height)
            # image.filename is None for file-like descriptors (BytesIO);
            # normalize to empty string to match python-docx's behavior.
            filename = image.filename or ""
            if cache_key is not None:
                cache[cache_key] = (rId, int(cx), int(cy), filename)

        # Always assign a fresh shape_id per insertion so that drawing IDs
        # are unique in every part (including headers/footers/footnotes
        # which are not renumbered by fix_docpr_ids()).
        self.tpl.docx_ids_index += 1
        shape_id = self.tpl.docx_ids_index

        # Generate XML from the fast template when compatible, with a native
        # python-docx fallback if its generated XML shape ever changes.
        pic = _format_inline_image_xml(shape_id, rId, filename, cx, cy)

        if self.anchor:
            run = parse_xml(pic)
            if run.xpath(".//a:blip"):
                hyperlink = self._add_hyperlink(
                    run, self.anchor, part
                )
                pic = hyperlink.xml

        return (
            "</w:t></w:r><w:r><w:drawing>%s</w:drawing></w:r><w:r>"
            '<w:t xml:space="preserve">' % pic
        )

    def __unicode__(self):
        return self._insert_image()

    def __str__(self):
        return self._insert_image()

    def __html__(self):
        return self._insert_image()
