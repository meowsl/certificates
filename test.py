from PyPDF2 import PdfReader, PdfWriter
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from PIL import ImageFont
import openpyxl

def create_cert(firstname, lastname, midname, id_text):

    packet = io.BytesIO()

    can = canvas.Canvas(packet, pagesize=letter)
    name = f'{lastname} {firstname} {midname}'
    can.setFont('OpenSans', 18)
    can.drawString(320, 260, name)
    can.save()

    packet.seek(0)

    new_pdf = PdfReader(packet)
    old_pdf = PdfReader('inputs/cert_without_sign.pdf')
    existing_pdf = PdfWriter(open("inputs/cert_without_sign.pdf", "rb"))
    output = PdfWriter()
    page = old_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)
    output_stream = open(f"outputs/{lastname}_{firstname[0]}._{midname[0]}.pdf", "wb")
    output.write(output_stream)
    output_stream.close()

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
    pdfmetrics.registerFont(TTFont('OpenSans', 'fonts/static/OpenSans-Regular.ttf'))
    extract_students()