# import os
# from datetime import datetime, timedelta
# from reportlab.platypus import *
# from reportlab.lib import colors
# from reportlab.lib.units import cm
from reportlab.lib.pagesizes import letter
# from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER
# from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.utils import ImageReader

from reportlab.pdfgen import canvas

class Factura(object):
    def __init__(self, name, background = True):
        self.width, self.height = letter
        self.height *= 0.5

        self.canvas = canvas.Canvas(name, pagesize = (self.width, self.height))

        self.canvas.setLineWidth(.3)
        self.canvas.setFont('Helvetica', 10)

        if background:
            logo = ImageReader('factura_tausand.jpg')
            self.canvas.drawImage(logo, 0, 0, height = self.height,
                                    width = self.width, mask = 'auto')

        self.setFacturaNumber("0001")
        self.setDate("01", "01", "2018")
        self.setSir("Juan Barbosa")
        self.setDocumento("1015457552")
        self.setDireccion("Calle 126 # 52 a 92 apto 101")
        self.setCiudad("Bogotá")
        self.setTelefono("6135083")
        self.setCorreo("js.barbosa10@uniandes.edu.co")

        cant = [str(i) for i in range(9)]
        desc = ["Descripción de prueba" for i in range(9)]
        valor = ["100,000" for i in range(9)]
        self.setItems(zip(cant, desc, valor))

    def setFacturaNumber(self, number):
        self.canvas.setFont('Helvetica', 15)
        self.canvas.drawCentredString(506, 336, number)

    def setDate(self, dd, mm, yy):
        h = 315
        self.canvas.setFont('Helvetica', 8)
        self.canvas.drawCentredString(487, h, dd)
        self.canvas.drawCentredString(516, h, mm)
        self.canvas.drawCentredString(552, h, yy)

    def setSir(self, text):
        self.canvas.drawString(96, 296, text)

    def setDocumento(self, text):
        self.canvas.drawString(463, 296, text)

    def setDireccion(self, text):
        self.canvas.drawString(97, 279, text)

    def setCiudad(self, text):
        self.canvas.drawString(277, 279, text)

    def setTelefono(self, text):
        self.canvas.drawString(376, 279, text)

    def setCorreo(self, text):
        self.canvas.drawString(461, 279, text)

    def setItems(self, items):
        y0 = 243
        y1 = 133
        d = (y1 - y0)/8

        for (i, row) in enumerate(items):
            self.canvas.drawString(74, y0 + i*d, row[0])
            self.canvas.drawString(94, y0 + i*d, row[1])
            self.canvas.drawString(499, y0 + i*d, row[2])

    def save(self):
        self.canvas.save()

factura = Factura("form.pdf", True)

#
# canvas.drawString(30,735,'OF ACME INDUSTRIES')
# canvas.drawString(500,750,"12/12/2010")
# canvas.line(480,747,580,747)
#
# canvas.drawString(275,725,'AMOUNT OWED:')
# canvas.drawString(500,725,"$1,000.00")
# canvas.line(378,723,580,723)
#
# canvas.drawString(30,703,'RECEIVED BY:')
# canvas.line(120,700,580,700)
# canvas.drawString(120,703,"JOHN DOE")

factura.save()
