from generator import generate_document
from postman_parser import parsePostmanJson

generate_document(parsePostmanJson("./tests/odata.json"))