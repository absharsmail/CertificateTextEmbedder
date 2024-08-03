import os
import pypdfium2 as pdfium
from PyPDF2 import PdfWriter, PdfReader


def embedOnPDF(certificatesFolder, sheet_name, name, packet, templatePdf, saveState):
    # create a new PDF with Reportlab
    new_pdf = PdfReader(packet)
    output = PdfWriter()
    existing_pdf = PdfReader(open(templatePdf, "rb"))
    # add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)
    directory = f'{certificatesFolder} {sheet_name}' if saveState != 2 else f'{certificatesFolder} {sheet_name}/{name}'
    os.makedirs(directory, exist_ok=True)  # exist_ok prevents errors if folder already exists
    new_pdf_name = os.path.join(directory, f"{name}.pdf")  # f"certificate_{row_num}.pdf"
    # finally, write "output" to a real file
    output_stream = open(new_pdf_name, "wb")
    output.write(output_stream)
    output_stream.close()
    if saveState != 0:
        # Load a document
        pdf = pdfium.PdfDocument(f"{new_pdf_name}")
        # Loop over pages and render
        for i in range(len(pdf)):
            page = pdf[i]
            image = page.render(scale=4).to_pil()
            image.save(f"{new_pdf_name}.jpg".replace(".pdf", ""))
        if saveState == 1:
            os.remove(new_pdf_name)
