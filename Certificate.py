import io
from openpyxl.reader.excel import load_workbook
from reportlab.lib import pagesizes
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
import EmbedOnPDF
import re


def nameFormatter(fullName):
    modName = ""
    modList = re.split('[. -]', str(fullName))
    initialAdded = False
    try:
        for i in modList:
            if i != "" and i.isalpha():
                if len(i) == 1:
                    initialAdded = True
                    modName = modName + i.upper()
                elif len(i) == 2:
                    initialAdded = True
                    modName = modName + i.upper()
                else:
                    if initialAdded:
                        modName = modName + " " + i.capitalize()
                    else:
                        modName = modName + i.capitalize() + " "
        if modName[-1] == " ":
            modName = modName[:len(modName) - 1]
    except IndexError:
        modName = ""
    return modName


def create_certificates(dataFile, templatePdf, certificatesFolder, pdfOnly):
    wb = load_workbook(filename=dataFile, data_only=True)
    sheetData = wb.get_sheet_names()
    for sheet_name in sheetData:
        data = wb.get_sheet_by_name(sheet_name).values
        for row_num, row in enumerate(data):
            if row[0] is None:
                break
            if row_num == 0:
                continue
            name = nameFormatter(row[0])
            packet = io.BytesIO()
            pdfmetrics.registerFont(TTFont('GothamMedium', 'assets/GothamMedium.ttf'))
            pdfmetrics.registerFont(TTFont('GothamBook', 'assets/GothamBook.ttf'))
            can = canvas.Canvas(packet, pagesize=pagesizes.landscape(pagesizes.B0))
            can.setFont('GothamMedium', 25)
            can.setFillColor("#ef5362")
            can.drawCentredString(430, 273, name)
            can.setFont('GothamBook', 18)
            print(name)
            can.save()
            # move to the beginning of the StringIO buffer
            packet.seek(0)
            EmbedOnPDF.embedOnPDF(certificatesFolder, sheet_name, name, packet, templatePdf, pdfOnly)


# Replace these with your file paths and starting number
data_file = "sheet.xlsx"
template_pdf = "certificate.pdf"
certificates_folder = "Certificates"
saveState = 1 # 0: Only Pdf, 1: Only Image, 2: Both Pdf and Image
create_certificates(data_file, template_pdf, certificates_folder, saveState)
print("Certificates created successfully!")
