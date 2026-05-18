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


def _build_inline_image_xml_template():
    """Generate the XML format string by calling python-docx with sentinel values.

    This ensures the template always matches the installed python-docx version's
    XML structure, even after upgrades. We call CT_Inline.new_pic_inline() once
    with recognizable sentinel values, serialize to XML, then replace the
    sentinels with Python format placeholders.
    """
    import uuid

    # Use GUIDs for string sentinels - guaranteed no collision with XML content
    _RID_SENTINEL = str(uuid.uuid4())
    _FILENAME_SENTINEL = str(uuid.uuid4())

    # For numeric sentinels, use unique integers derived from UUIDs.
    # shape_id is xsd:unsignedInt (max 4,294,967,295 / 32-bit).
    # cx/cy are EMU values typed as xsd:long (64-bit).
    # All use 9-digit range [100000000, 999999999] to stay within 32-bit
    # and avoid any accidental collisions with each other.
    _SHAPE_ID = uuid.uuid4().int % (9 * 10**8) + 10**8
    _CX_INT = uuid.uuid4().int % (9 * 10**8) + 10**8
    _CY_INT = uuid.uuid4().int % (9 * 10**8) + 10**8

    inline = CT_Inline.new_pic_inline(
        _SHAPE_ID,
        _RID_SENTINEL,
        _FILENAME_SENTINEL,
        Emu(_CX_INT),
        Emu(_CY_INT),
    )
    xml = inline.xml

    # Replace sentinel values with format string placeholders
    xml = xml.replace(str(_SHAPE_ID), "{shape_id}")
    xml = xml.replace(_RID_SENTINEL, "{rId}")
    xml = xml.replace(_FILENAME_SENTINEL, "{filename}")
    xml = xml.replace(str(_CX_INT), "{cx}")
    xml = xml.replace(str(_CY_INT), "{cy}")

    return xml


# Pre-built XML template for inline images, derived from the installed
# python-docx version. Using str.format() on this template avoids calling
# CT_Inline.new_pic_inline() per image (which does 2x parse_xml() +
# element manipulation + .xml serialization each time).
_INLINE_IMAGE_XML = _build_inline_image_xml_template()


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

        # Cache generated XML per (part, descriptor, width, height) to avoid
        # repeated file I/O, SHA1 computation, and header parsing.
        cache = self.tpl._image_cache
        cache_key = (id(part), image_descriptor, self.width, self.height)

        if cache_key in cache:
            pic = cache[cache_key]
        else:
            # Get or add the image part (handles deduplication via SHA1 internally)
            package = part._package
            image_part = package.get_or_add_image_part(image_descriptor)
            rId = part.relate_to(image_part, RT.IMAGE)
            image = image_part.image
            cx, cy = image.scaled_dimensions(self.width, self.height)

            # Assign shape_id from a simple counter. python-docx's
            # new_pic_inline() would call its next_id property which does an
            # XPath("//@id") over the entire XML tree on every call - but we
            # bypass that entirely by generating the XML ourselves.
            # fix_docpr_ids() renumbers all IDs after rendering anyway.
            self.tpl.docx_ids_index += 1
            shape_id = self.tpl.docx_ids_index

            # Generate XML directly as a string using a pre-built template
            # rather than calling CT_Inline.new_pic_inline() per image.
            pic = _INLINE_IMAGE_XML.format(
                cx=int(cx),
                cy=int(cy),
                shape_id=shape_id,
                filename=xml_escape(image.filename),
                rId=rId,
            )
            cache[cache_key] = pic

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
