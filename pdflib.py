import os
import constants
from textwrap import TextWrapper
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph
from reportlab.lib.pagesizes import letter
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.utils import ImageReader
from reportlab.lib.styles import ParagraphStyle

class FacturaPDFs(object):
    def __init__(self, factura):
        dir = factura.getPDFDir()

        digital = dir + constants.PDF_DIGITAL
        print_ = dir + constants.PDF_PRINT

        self.digital = FacturaPDF(digital, background = True)
        self.print = FacturaPDF(print_, background = False)
        self.setFromFactura(factura)
        self.save()

    def setFromFactura(self, factura):
        self.digital.setFromFactura(factura)
        self.print.setFromFactura(factura)

    def save(self):
        self.digital.save()
        self.print.save()

class FacturaPDF(object):
    def __init__(self, name, background = True):
        self.background = background
        self.width, self.height = letter
        self.height *= 0.5

        self.canvas = canvas.Canvas(name, pagesize = (self.width, self.height))

        self.font_name = 'Courier'
        self.canvas.setLineWidth(.3)
        self.canvas.setFont(self.font_name, 8)

        if self.background:
            logo = ImageReader(constants.BACKGROUND_FACTURA)
            self.canvas.drawImage(logo, 0, 0, height = self.height,
                                    width = self.width, mask = 'auto')

    def setFacturaNumber(self, number):
        self.canvas.setFont(self.font_name, 15)
        self.canvas.drawCentredString(506, 336 + 1*mm, number)
        self.canvas.setFont(self.font_name, 8)

    def setDate(self, dd, mm_, yy):
        h = 315
        if not self.background:
            h += 2*mm
        self.canvas.drawCentredString(487, h, dd)
        self.canvas.drawCentredString(516, h, mm_)
        self.canvas.drawCentredString(552, h, yy)

    def setSir(self, text):
        h = 296
        w = 96
        if not self.background:
            h += 3*mm
            w -= 3*mm
        self.canvas.drawString(w, h, text)

    def setDocumento(self, text):
        h = 296
        w = 463
        if not self.background:
            h += 3*mm
        self.canvas.drawString(w, h, text)

    def setDireccion(self, text):
        h = 279
        w = 97
        if not self.background:
            h += 3*mm
            w -= 3*mm
        self.canvas.drawString(w, h, text)

    def setCiudad(self, text):
        h = 279
        w = 277
        if not self.background:
            h += 3*mm
        self.canvas.drawString(w, h, text)

    def setTelefono(self, text):
        h = 279
        w = 376
        if not self.background:
            h += 3*mm
        self.canvas.drawString(w, h, text)

    def setCorreo(self, text):
        h = 279
        w = 461
        if not self.background:
            h += 3*mm
        self.canvas.setFont(self.font_name, 6)
        self.canvas.drawString(w, h, text)
        self.canvas.setFont(self.font_name, 8)

    def setItems(self, items):
        y0 = 245
        y1 = 135
        d = (y1 - y0)/8
        d1 = 0
        d2 = 0
        d3 = 0
        x1 = 0
        x2 = 0
        x3 = 0

        if not self.background:
            d1 = 2*mm
            d2 = 2*mm
            d3 = 1*mm
            x1 = -3*mm
            x2 = -3*mm

        for (i, row) in enumerate(items):
            self.canvas.drawRightString(88 + x1, y0 + i*d + d1, row[0])
            self.canvas.drawString(94 + x2, y0 + i*d + d2, row[1])
            self.canvas.drawRightString(574 + x3, y0 + i*d + d3, row[2])

    def setLast(self, sub, iva, flete, retefuente, total):
        y0 = 124
        y1 = 90
        d = (y1 - y0)/3
        if not self.background:
            y0 += 1*mm

        self.canvas.drawRightString(574, y0, sub)
        self.canvas.drawRightString(574, y0 + d, iva)
        self.canvas.drawRightString(574, y0 + 2*d, flete)
        self.canvas.drawRightString(574, y0 + 3*d, retefuente)

        self.canvas.setFont(self.font_name, 10)
        if self.background:
            self.canvas.drawRightString(574, 76, total)
        else:
            self.canvas.drawRightString(574, 76 + 1*mm, total)
        self.canvas.setFont(self.font_name, 8)

    def setObservaciones(self, text):
        style = ParagraphStyle(name = 'Justify', alignment = TA_JUSTIFY, fontSize = 8,
                           fontName = self.font_name)

        w = TextWrapper(width = 75, break_long_words = False, replace_whitespace = False)
        text = '\n'.join(w.wrap(text))
        text = text.replace('\n', '<br/>')
        n = text.count('<br/>')

        p = Paragraph(text, style)

        voffset = 0

        x1, y1 = 59, 108
        x2, y2 = 445, 76

        if not self.background:
            y1 += 2*mm
            y2 += 2*mm

        p.wrapOn(self.canvas, x2 - x1, y1 - y2)
        p.drawOn(self.canvas, x1, y1 - n*12)

    def setFromFactura(self, factura):
        now = datetime.now()
        if self.background:
            self.setFacturaNumber(factura.getNumeroS())
        self.setDate("%02d"%now.day, "%02d"%now.month, "%d"%now.year)
        self.setSir(factura.getNombre())
        self.setDocumento(factura.getDocumento())
        self.setDireccion(factura.getDireccion())
        self.setCiudad(factura.getCiudad())
        self.setTelefono(factura.getTelefono())
        self.setCorreo(factura.getCorreo())
        self.setItems(factura.makeTable())

        self.setLast(factura.getSubTotalS(), factura.getIVAS(), factura.getFleteS(),
                        factura.getReteFuenteS(), factura.getTotalS())

        obs = factura.getObservaciones()
        self.setObservaciones(obs)

    def save(self):
        self.canvas.save()

class CotizacionPDF(object):
    def __init__(self, cotizacion):
        self.width, self.height = letter
        self.height *= 0.5

        name = cotizacion.getPDFDir()

        self.canvas = canvas.Canvas(name, pagesize = (self.width, self.height))

        self.font_name = 'Courier'
        self.canvas.setLineWidth(.3)
        self.canvas.setFont(self.font_name, 10)

        logo = ImageReader(constants.BACKGROUND_COTIZACION)
        self.canvas.drawImage(logo, 0, 0, height = self.height,
                                width = self.width, mask = 'auto')

        self.setFromCotizacion(cotizacion)
        self.save()

    def setCotizacionNumber(self, number):
        self.canvas.setFont(self.font_name, 15)
        self.canvas.drawCentredString(506, 336, number)
        self.canvas.setFont(self.font_name, 8)

    def setDate(self, dd, mm, yy):
        h = 315
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
        self.canvas.setFont(self.font_name, 6)
        self.canvas.drawString(461, 279, text)
        self.canvas.setFont(self.font_name, 8)

    def setItems(self, items):
        y0 = 245
        y1 = 135
        d = (y1 - y0)/8

        for (i, row) in enumerate(items):
            self.canvas.drawRightString(88, y0 + i*d, row[0])
            self.canvas.drawString(94, y0 + i*d, row[1])
            self.canvas.drawRightString(574, y0 + i*d, row[2])

    def setLast(self, sub, iva, flete, retefuente, total):
        y0 = 124
        y1 = 90
        d = (y1 - y0)/3

        self.canvas.drawRightString(574, y0, sub)
        self.canvas.drawRightString(574, y0 + d, iva)
        self.canvas.drawRightString(574, y0 + 2*d, flete)
        self.canvas.drawRightString(574, y0 + 3*d, retefuente)

        self.canvas.setFont(self.font_name, 10)
        self.canvas.drawRightString(574, 76, total)

    def setObservaciones(self, text):
        style = ParagraphStyle(name = 'Justify', alignment = TA_JUSTIFY, fontSize = 8,
                           fontName=self.font_name)
        p = Paragraph(text, style)

        voffset = 0

        x1, y1 = 59, 120
        x2, y2 = 439, 76

        p.wrapOn(self.canvas, x2 - x1, y1 - y2)
        p.drawOn(self.canvas, x1, y2)

    def setFromCotizacion(self, cotizacion):
        now = datetime.now()
        self.setCotizacionNumber(cotizacion.getNumeroS())
        self.setDate("%02d"%now.day, "%02d"%now.month, "%d"%now.year)
        self.setSir(cotizacion.getNombre())
        self.setDocumento(cotizacion.getDocumento())
        self.setDireccion(cotizacion.getDireccion())
        self.setCiudad(cotizacion.getCiudad())
        self.setTelefono(cotizacion.getTelefono())
        self.setCorreo(cotizacion.getCorreo())
        self.setItems(cotizacion.makeTable())

        self.setLast(cotizacion.getSubTotalS(), cotizacion.getIVAS(), cotizacion.getFleteS(),
                        cotizacion.getReteFuenteS(), cotizacion.getTotalS())

        obs = cotizacion.getObservaciones().replace('\n','<br/>')
        self.setObservaciones(obs)

    def save(self):
        self.canvas.save()
