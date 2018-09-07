from generator import generate_document
from parser import parsePostmanJson

generate_document(parsePostmanJson("./tests/odata.json"))