import io
import re

from docxtpl import DocxTemplate, InlineImage


def image_bytes(path):
    with open(path, "rb") as image_file:
        return image_file.read()


def embedded_rid(xml):
    return re.search(r'r:embed="([^"]+)"', xml).group(1)


tpl = DocxTemplate("templates/inline_image_tpl.docx")
tpl.render_init()
tpl.current_rendering_part = tpl.docx._part

stream = io.BytesIO(image_bytes("templates/django.png"))
first_xml = str(InlineImage(tpl, stream))

stream.seek(0)
stream.truncate()
stream.write(image_bytes("templates/python.png"))
stream.seek(0)
second_xml = str(InlineImage(tpl, stream))

assert embedded_rid(first_xml) != embedded_rid(second_xml)
assert tpl._image_cache == {}
