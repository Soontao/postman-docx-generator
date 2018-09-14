from generator import generate_document, DocumentGenerator
from postman_parser import parsePostmanJson

metadata = parsePostmanJson("./tests/odata.json")
document = DocumentGenerator(metadata)
document.save("{}.docx".format(metadata.info.name))
