from PIL import ImageFont
from PIL import Image
from PIL import ImageDraw
import openpyxl
from PyPDF2 import PdfReader, PdfWriter
import fitz
from io import BytesIO

def create_cert(firstname, lastname, midname, id_text):

    doc = fitz.open('inputs/cert_without_sign.pdf')
    page = doc.load_page(0)

    pdf_width, pdf_height = page.rect[2:]

    img = Image.new('L', size=(int(pdf_width), int(pdf_height)))
    draw = ImageDraw.Draw(img)

    # Calculate appropriate image DPI based on PDF DPI and dimensions
    pdf_dpi = 300  # Assuming horizontal and vertical DPI are equal
    img_dpi = int(round(pdf_dpi * img.width / pdf_width))

    name = f'{lastname} {firstname} {midname}'
    # Choose appropriate font sizes and positions based on your layout
    name_font = ImageFont.truetype("fonts/static/OpenSans-Regular.ttf", size=18)
    name_color = (0, 0, 0)
    name_position = (((img.width - name_font.getlength(name)) / 2), 320)

    draw.text(name_position, name, font=name_font, fill=name_color)

    id_font = ImageFont.truetype("fonts/static/OpenSans_SemiCondensed-Italic.ttf", size=12)
    id_position = (190, img.height - 95)
    id_color = (0, 0, 0)

    draw.text(id_position, id_text, font=id_font, fill=id_color)

    img.save(f'outputs/{lastname}_{firstname[0]}_{midname[0]}.png', format='PNG', dpi=img_dpi, quality=100)

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
    extract_students()

# from PIL import ImageFont
# from PIL import Image
# from PIL import ImageDraw
# import openpyxl
# from PyPDF2 import PdfReader, PdfWriter
# import fitz
# from io import BytesIO

# # определяете шрифт
# def create_cert(firstname, lastname, midname, id_text):
#     #
#     reader = PdfReader('inputs/cert_without_sign.pdf')
#     doc = fitz.open('inputs/cert_without_sign.pdf')
#     page = doc.load_page(0)

#     pix = page.get_pixmap()
#     byt = pix.tobytes(output='png')

#     img = Image.open(BytesIO(byt))
#     img.save('test/test.png', "PNG", quality=100)
#     draw = ImageDraw.Draw(img)

#     name_font = ImageFont.truetype("fonts/static/OpenSans-Regular.ttf", size=18)
#     name_color = (0,0, 0)
#     name = f'{lastname} {firstname} {midname}'
#     name_position = (((img.width - name_font.getlength(name)) / 2), 320)
#     draw.text(name_position, name, font=name_font, fill=name_color)

#     id_font = ImageFont.truetype("fonts/static/OpenSans_SemiCondensed-Italic.ttf", size=12)
#     id_position = (190, img.height - 95)
#     id_color = (0, 0, 0)
#     draw.text(id_position, id_text, font=id_font, fill=id_color)

#     img.save(f'outputs/{lastname}_{firstname[0]}_{midname[0]}.png', "PNG", quality=100)

# def extract_elements(row):
#     elements = []
#     for cell in row:
#       elements.append(str(cell.value))
#     return elements

# def extract_students():
#     wb = openpyxl.load_workbook('inputs/students.xlsx')
#     sheet = wb['Лист1']
#     for row in sheet.iter_rows():
#       elements = extract_elements(row)
#       create_cert(elements[1], elements[2], elements[3], elements[0])

# if __name__ == "__main__":
#     extract_students()
#     # create_cert()