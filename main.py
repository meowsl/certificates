import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from PIL import ImageFont
from PyPDF2 import PdfReader, PdfWriter
import time
import openpyxl
from concurrent.futures import ThreadPoolExecutor
from progress.bar import IncrementalBar, Bar

def create_cert(lastname, firstname, midname, id_text):
    packet = io.BytesIO()
    name = f'{lastname} {firstname} {midname}'
    name_width, name_height = font.getmask(name).size
    can = canvas.Canvas(packet, pagesize=letter)

    text_x = (size.width - name_width) / 2
    text_y = (size.height - name_height) / 2

    can.setFont("OpenSansItalic", 12)
    can.drawString(200, int(text_y) - 210, id_text)

    can.setFont("OpenSans", 18)
    can.drawString(int(text_x), int(text_y) - 25, name)

    packet.seek(0)
    can.save()

    new_pdf = PdfReader(packet)
    output = PdfWriter()
    page = old_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)

    with open(f"outputs/{lastname}_{firstname[0]}_{midname[0]}.pdf", "wb") as output_stream:
        output.write(output_stream)

    bar.next()

def process_row(row):
    elems = [str(cell.value) for cell in row]
    create_cert(elems[1], elems[2], elems[3], elems[0])

def extract_students():
    global old_pdf, font, size

    pdfmetrics.registerFont(TTFont('OpenSans', 'fonts/OpenSans-Regular.ttf'))
    pdfmetrics.registerFont(TTFont('OpenSansItalic', 'fonts/OpenSans_SemiCondensed-Italic.ttf'))
    old_pdf = PdfReader('inputs/cert.pdf')
    old_pdf_page = old_pdf.pages[0]
    size = old_pdf_page.mediabox
    font = ImageFont.truetype("fonts/OpenSans-Regular.ttf", 18)

    with ThreadPoolExecutor() as executor:
        for row in sheet.iter_rows():
            executor.submit(process_row, row)

if __name__ == "__main__":
    start = time.time()

    wb = openpyxl.load_workbook('inputs/students.xlsx')
    sheet = wb['Лист1']
    bar = Bar('Progress: ', max=sheet.max_row, fill='*', suffix='%(percent)d%%')
    bar.update()
    extract_students()

    print(f'Выполнилось за {time.time() - start} секунд')