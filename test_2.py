import io

from PyPDF2 import PdfWriter, PdfReader
from reportlab.lib import pagesizes
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

packet = io.BytesIO()
pdfmetrics.registerFont(TTFont('GothamMedium', 'assets/GothamMedium.ttf'))
pdfmetrics.registerFont(TTFont('GothamBook', 'assets/GothamBook.ttf'))
can = canvas.Canvas(packet, pagesize=pagesizes.landscape(pagesizes.B0))
can.setFont('GothamMedium', 25)
can.setFillColor("#ef5362")
can.drawCentredString(430, 273, "Aisha Muhammed musthafa")
can.setFont('GothamBook', 18)
can.drawCentredString(535, 225, "152")
can.save()

# move to the beginning of the StringIO buffer
packet.seek(0)

# create a new PDF with Reportlab
new_pdf = PdfReader(packet)
# read your existing PDF
existing_pdf = PdfReader(open("certificate.pdf", "rb"))
output = PdfWriter()
# add the "watermark" (which is the new pdf) on the existing page
page = existing_pdf.pages[0]
page.merge_page(new_pdf.pages[0])
output.add_page(page)

# finally, write "output" to a real file
output_stream = open("destination.pdf", "wb")
output.write(output_stream)
output_stream.close()

# # Load a document
# pdf = pdfium.PdfDocument("destination.pdf")
#
# # Loop over pages and render
# for i in range(len(pdf)):
#     page = pdf[i]
#     image = page.render(scale=4).to_pil()
#     image.save(f"output_{i:03d}.jpg")