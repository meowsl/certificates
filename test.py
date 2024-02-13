from PyPDF2 import PdfReader, PdfWriter
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import ImageFont
import openpyxl
import time

# def create_cert(lastname, firstname, midname, id_text):
#     packet = io.BytesIO()
#     name = f'{lastname} {firstname} {midname}'
#     name_width, name_height = font.getmask(name).size

#     can = canvas.Canvas(packet, pagesize=letter)
#     text_x = (size.width - name_width) / 2
#     text_y = (size.height - name_height) / 2

#     can.setFont('OpenSans', 18)
#     can.drawString(int(text_x), int(text_y), name)
#     can.setFont('OpenSansItalic', 12)
#     can.drawString(200, int(text_y) - 50, id_text)

#     packet.seek(0)
#     can.save()

#     new_pdf = PdfReader(packet)
#     existing_pdf = PdfWriter(open("inputs/cert.pdf", "rb"))
#     output = PdfWriter()

#     page = old_page.pages[0]
#     page.merge_page(new_pdf.pages[0])
#     output.add_page(page)

#     output_filename = f"outputs/{lastname}_{firstname[0]}._{midname[0]}.pdf"
#     with open(output_filename, "wb") as output_stream:
#         output.write(output_stream)

#     print(f'{output_filename} was created!')

def create_cert(lastname, firstname, midname, id_text, font_size=18):
    packet = io.BytesIO()
    name = f'{lastname} {firstname} {midname}'
    name_width, name_height = font.getmask(name).size

    can = canvas.Canvas(packet, pagesize=letter)
    text_x = (size.width - name_width) / 2
    text_y = (size.height - name_height) / 2

    can.setFont('OpenSans', font_size)
    can.drawString(int(text_x), int(text_y), name)
    can.setFont('OpenSansItalic', 12)
    can.drawString(200, int(text_y) - 50, id_text)

    can.save()

    # Merge pages efficiently
    existing_pdf = PdfReader(packet)
    output = PdfWriter()
    old_page.pages[0].merge_page(existing_pdf.pages[0])
    output.add_page(old_page)

    output_filename = f"outputs/{lastname}_{firstname[0]}._{midname[0]}.pdf"
    with open(output_filename, "wb") as output_stream:
        output.write(output_stream)

    print(f'{output_filename} was created!')

def extract_elements(row):
    """Extracts elements from a single row in an Excel sheet.

    Args:
        row (openpyxl.row.Row): A row object from an Excel sheet.

    Returns:
        list: A list of cell values as strings.
    """

    elements = [str(cell.value).strip() for cell in row]  # Strip whitespace
    return elements


def extract_students():
    """Extracts student data from an Excel sheet and creates personalized certificates.

    Returns:
        None
    """

    wb = openpyxl.load_workbook('inputs/students.xlsx')
    sheet = wb['Лист1']

    for row in sheet.iter_rows():
        elements = extract_elements(row)
        create_cert(elements[1], elements[2], elements[3], elements[0])


if __name__ == "__main__":
    # Register fonts once, reducing overhead
    pdfmetrics.registerFont(TTFont('OpenSans', 'fonts/OpenSans-Regular.ttf'))
    pdfmetrics.registerFont(TTFont('OpenSansItalic', 'fonts/OpenSans_SemiCondensed-Italic.ttf'))

    old_page = PdfReader('inputs/cert.pdf')  # Cache template PDF
    size = old_page.pages[0].mediabox
    font = ImageFont.truetype("fonts/OpenSans-Regular.ttf", 18)  # Create font once
    start = time.time()
    extract_students()
    end = time.time()
    print(f'Выполнилось за {start - end}')
