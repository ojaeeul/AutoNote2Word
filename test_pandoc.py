import pypandoc
from docx import Document
from docx.shared import Cm
import os

markdown_text = "Here is an equation: $$E = mc^2$$"
output_filename = "test_hq.docx"
margins = {'top': 2.0, 'bottom': 2.0, 'left': 2.5, 'right': 2.5}

try:
    pypandoc.get_pandoc_version()
except OSError:
    print("Downloading pandoc...")
    pypandoc.download_pandoc()

temp_md = "temp.md"
with open(temp_md, "w", encoding="utf-8") as f:
    f.write(markdown_text)

pypandoc.convert_file(temp_md, 'docx', outputfile=output_filename)

doc = Document(output_filename)
for section in doc.sections:
    section.top_margin = Cm(margins['top'])
    section.bottom_margin = Cm(margins['bottom'])
    section.left_margin = Cm(margins['left'])
    section.right_margin = Cm(margins['right'])
doc.save(output_filename)
os.remove(temp_md)
print("Success!")
