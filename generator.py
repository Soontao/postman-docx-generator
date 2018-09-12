from docx import Document
from docx.shared import Inches
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Pt
from postman_schema import PostmanObject, Items
from string import Template
import datetime


def apply_style(document: Document):
  styles = document.styles
  heading_1 = styles['Heading 1']
  heading_1.font.size = Pt(25)


def generate_document(postmanObject: PostmanObject):
  document = Document()
  apply_style(document)
  document.add_heading(postmanObject.info.name, 0)
  document.add_paragraph(postmanObject.info.description)
  assert len(postmanObject.item) > 0, "collection must have at least one request"
  document.add_heading("Version", 1)
  add_table_to_document(
      document,
      [["Version", "Date", "Comments"], [1, datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), "init document"]])

  for api_item in postmanObject.item:
    assert len(api_item.response) > 0, "each api must have at least one sample"
    generate_api_chapter(api_item, document)
  document.save("{}.docx".format(postmanObject.info.name))


def generate_api_chapter(api_item: Items, document: Document):

  document.add_heading(api_item.name, 1)
  document.add_paragraph(api_item.request.description)

  # request basic info
  document.add_heading("Request", 2)
  add_table_to_document(document, [["Method", "URL"], [api_item.request.method, api_item.request.url.raw]])

  return document


def add_table_to_document(
    document: Document,
    rows,
):
  assert len(rows) > 0, "table must have at list one records"
  column_number = len(rows[0])
  rows_number = len(rows)
  table = document.add_table(rows=0, cols=column_number)
  table.autofit = True
  table.alignment = WD_TABLE_ALIGNMENT.CENTER
  for index in range(rows_number):
    row = table.add_row().cells
    for col_index, col_content in enumerate(rows[index]):
      cell = row[col_index]
      cell.text = str(col_content)
      cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
  return document