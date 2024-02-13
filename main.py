import datetime
import time
from PyPDF2 import PdfReader, PdfWriter
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from PIL import ImageFont
import openpyxl

def create_cert(lastname, firstname, midname, id_text):

    old_pdf = PdfReader('inputs/cert_without_sign.pdf')
    sz = old_pdf.pages[0].mediabox

    packet = io.BytesIO()

    font_name = "fonts/static/OpenSans-Regular.ttf"
    font = ImageFont.truetype(font_name, 18)
    name = f'{lastname} {firstname} {midname}'
    name_width, name_height = font.getmask(name).size

    can = canvas.Canvas(packet, pagesize=letter)

    text_x = (sz.width - name_width) / 2
    text_y = (sz.height - name_height) / 2
    can.setFont('OpenSansItalic', 12)
    can.drawString(200, int(text_y) - 210, id_text)

    can.setFont('OpenSans', 18)
    can.drawString(int(text_x), int(text_y) - 25 , name)

    packet.seek(0)
    can.save()

    new_pdf = PdfReader(packet)
    existing_pdf = PdfWriter(open("inputs/cert_without_sign.pdf", "rb"))
    output = PdfWriter()

    page = old_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)
    output_stream = open(f"outputs/{lastname}_{firstname[0]}._{midname[0]}.pdf", "wb")
    output.write(output_stream)
    output_stream.close()
    print(f'{lastname}_{firstname[0]}._{midname[0]}.pdf was created!')

def extract_elements(row):
    elements = []
    for cell in row:
        elements.append(str(cell.value))
    return elements

def extract_students():
    wb = openpyxl.load_workbook('inputs/students.xlsx')
    sheet = wb['Лист1']
    for row in sheet.iter_rows():
        elements = extract_elements(row)
        create_cert(elements[1], elements[2], elements[3], elements[0])

if __name__ == "__main__":
    # Register font to support Cyrillic characters
    pdfmetrics.registerFont(TTFont('OpenSans', 'fonts/static/OpenSans-Regular.ttf'))
    pdfmetrics.registerFont(TTFont('OpenSansItalic', 'fonts/static/OpenSans_SemiCondensed-Italic.ttf'))
    extract_students()