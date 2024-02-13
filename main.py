from PyPDF2 import PdfReader, PdfWriter
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from PIL import ImageFont
import openpyxl
import time
import threading
from queue import Queue
from alive_progress import alive_bar
from progress.bar import IncrementalBar, Bar

def create_cert(lastname, firstname, midname, id_text):

    packet = io.BytesIO()

    name = f'{lastname} {firstname} {midname}'
    name_width, name_height = font.getmask(name).size

    can = canvas.Canvas(packet, pagesize=letter)

    text_x = (size.width - name_width) / 2
    text_y = (size.height - name_height) / 2
    can.setFont('OpenSansItalic', 12)
    can.drawString(200, int(text_y) - 210, id_text)

    can.setFont('OpenSans', 18)
    can.drawString(int(text_x), int(text_y) - 25 , name)

    packet.seek(0)
    can.save()

    new_pdf = PdfReader(packet)
    output = PdfWriter()

    page = old_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)

    with open(f"outputs/{lastname}_{firstname[0]}._{midname[0]}.pdf", "wb") as output_stream:
      output.write(output_stream)

    bar.next()

def extract_elements(row):
    elements = []
    for cell in row:
        elements.append(str(cell.value))
    return elements

def extract_students():
    global bar, start
    wb = openpyxl.load_workbook('inputs/students.xlsx')
    sheet = wb['Лист1']

    # Очередь для хранения данных о студентах
    queue = Queue()
    start = time.time()

    # Инициализация progress bar
    bar = Bar('Progress: ', max=sheet.max_row, fill='*', suffix='%(percent)d%%')

    # Функция для потоков
    def worker():
        while True:
            # Получить данные о студенте из очереди
            elements = queue.get()
            if elements is None:
                break
            create_cert(elements[1], elements[2], elements[3], elements[0])

    # Запуск потоков
    num_threads = 10  # Количество потоков
    threads = []
    for _ in range(num_threads):
        thread = threading.Thread(target=worker)
        thread.start()
        threads.append(thread)
        bar.update()

    # Добавление данных о студентах в очередь
    for row in sheet.iter_rows():
        elements = extract_elements(row)
        queue.put(elements)

    # Остановка потоков
    for _ in range(num_threads):
        queue.put(None)

    for thread in threads:
        thread.join()

    bar.finish()

if __name__ == "__main__":
    # Register font to support Cyrillic characters
    pdfmetrics.registerFont(TTFont('OpenSans', 'fonts/OpenSans-Regular.ttf'))
    pdfmetrics.registerFont(TTFont('OpenSansItalic', 'fonts/OpenSans_SemiCondensed-Italic.ttf'))

    old_pdf = PdfReader('inputs/cert.pdf')
    size = old_pdf.pages[0].mediabox
    font = ImageFont.truetype("fonts/OpenSans-Regular.ttf", 18)

    print('Start')
    extract_students()
    end = time.time()
    print(f'Выполнилось за {end - start}')
