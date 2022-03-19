from reportlab.pdfgen import canvas

pdf = canvas.Canvas(filemane)
pdf.setTitle(Nicolas)
pdf.save()