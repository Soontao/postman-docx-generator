from docx import Document
from docx.shared import Inches
from postman_schema import PostmanObject, Items
from string import Template


def generate_document(postmanObject: PostmanObject):
  document = Document()
  document.add_heading(postmanObject.info.name, 0)
  document.add_paragraph(postmanObject.info.description)
  assert len(postmanObject.item) > 0, "collection must have at least one request"
  for api_item in postmanObject.item:
    assert len(api_item.response) > 0, "each api must have at least one sample"
    generate_api_chapter(api_item, document)
  document.save("{}.docx".format(postmanObject.info.name))


def generate_api_chapter(api_item: Items, document: Document):
  document.add_heading(api_item.name, 1)
  document.add_paragraph(api_item.request.description)

  # request basic info
  document.add_heading("Request", 2)

  add_table_to_document(document, ["Method", "URL"], [[api_item.request.method, api_item.request.url.raw]])

  return document


def add_table_to_document(
    document: Document,
    headers,
    rows,
):
  assert len(rows) > 0, "table must have at list one records"
  assert len(headers) == len(rows[0]), "column number must equal to headers number"
  column_number = len(headers)
  rows_number = len(rows)
  table = document.add_table(rows=1, cols=column_number)
  header_row = table.rows[0].cells
  for index, header in enumerate(headers):
    header_row[index].text = header
  for index in range(rows_number):
    row = table.add_row().cells
    for col_index, col_content in enumerate(rows[index]):
      row[col_index].text = col_content
  return document