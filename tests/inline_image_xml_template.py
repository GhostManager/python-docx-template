from docx.oxml import parse_xml

import docxtpl.inline_image as inline_image


_INLINE_IMAGE_XML = inline_image._get_inline_image_xml_template()

assert _INLINE_IMAGE_XML.count("{shape_id}") == 2
assert _INLINE_IMAGE_XML.count("{cx}") == 2
assert _INLINE_IMAGE_XML.count("{cy}") == 2
assert _INLINE_IMAGE_XML.count("{rId}") == 1
assert _INLINE_IMAGE_XML.count("{filename}") == 1

parse_xml(
    _INLINE_IMAGE_XML.format(
        shape_id=1,
        cx=2,
        cy=3,
        rId="rId1",
        filename="image.png",
    )
)


def raise_incompatible_template():
    raise RuntimeError("incompatible template")


original_template = inline_image._INLINE_IMAGE_XML
original_builder = inline_image._build_inline_image_xml_template
try:
    inline_image._INLINE_IMAGE_XML = None
    inline_image._build_inline_image_xml_template = raise_incompatible_template
    fallback_xml = inline_image._format_inline_image_xml(
        shape_id=1,
        rId="rId1",
        filename='quoted " image.png',
        cx=2,
        cy=3,
    )
    parse_xml(fallback_xml)
    assert 'name="quoted &quot; image.png"' in fallback_xml
finally:
    inline_image._INLINE_IMAGE_XML = original_template
    inline_image._build_inline_image_xml_template = original_builder
