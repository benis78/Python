from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter

pdf_file = "/Users/jesper/Dropbox/Coding/PDF Watermark/0000.pdf"
watermark = "/Users/jesper/Dropbox/Coding/PDF Watermark/WatermarkA4.pdf"
merged = "/Users/jesper/Dropbox/Coding/PDF Watermark/merged.pdf"

with open(pdf_file, "rb") as input_file, open(watermark, "rb") as watermark_file:
    input_pdf = PdfFileReader(input_file)
    watermark_pdf = PdfFileReader(watermark_file)
    watermark_page = watermark_pdf.getPage(0)
    xw, yw, ww, hw = watermark_page.getPage(0).mediaBox 
    output = PdfFileWriter()

    for i in range(input_pdf.getNumPages()):
        xi, yi, wi, hi = input_pdf.getPage(i).mediaBox 
        if 
        pdf_page = input_pdf.getPage(i)
        pdf_size = input_pdf.getPage(i).mediaBox
        print(pdf_size)
        pdf_page.mergePage(watermark_page)
        output.addPage(pdf_page)

    with open(merged, "wb") as merged_file:
        output.write(merged_file)