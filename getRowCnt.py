import xlrd
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
import sys

wb = xlrd.open_workbook(sys.argv[1])
ws = wb.sheet_by_index(0)
rowNum = ws.nrows
print(rowNum - 1)
