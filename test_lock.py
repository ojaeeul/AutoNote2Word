from docx import Document
from docx.oxml import parse_xml

doc = Document()
doc.add_paragraph("This is a test document.")
doc.settings.element.append(parse_xml(r'<w:documentProtection w:edit="readOnly" w:enforcement="1" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
doc.save("test_readonly.docx")
print("Saved test_readonly.docx")
