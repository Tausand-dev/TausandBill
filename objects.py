import os
import pickle
import numpy as np
import pandas as pd
from textwrap import wrap
from datetime import datetime

import constants
from unidecode import unidecode
from pdflib import FacturaPDFs, CotizacionPDF

def makeDir(dir):
    if os.path.isdir(dir): pass
    else: os.makedirs(dir)

makeDir(constants.PDF_DIR)
makeDir(constants.OLD_DIR)

def readDataFrames():
    c = pd.read_excel(constants.CLIENTES_FILE, dtype = str)#.fillna("").astype(str)
    r = pd.read_excel(constants.REGISTRO_FILE, dtype = str)#.fillna("").astype(str)
    return c, r

CLIENTES_DATAFRAME, REGISTRO_DATAFRAME = readDataFrames()

class Documento(object):
    def __init__(self, numero = None, usuario = None, servicios = [], observaciones = "", iva = 0.19, flete = 0, retefuente = 0):
        self.numero = numero
        self.usuario = usuario
        self.observaciones = observaciones
        self.iva_coeff = iva
        self.flete = flete
        self.retefuente = retefuente
        self.setServicios(servicios)

    def getUsuario(self):
        return self.usuario

    def getCodigos(self):
        return [servicio.getCodigo() for servicio in self.servicios]

    def getNumero(self):
        return self.numero

    def getSubTotal(self):
        return sum([servicio.getValorTotal() for servicio in self.servicios])

    def getIVACoeff(self):
        return self.iva_coeff

    def getIVA(self):
        return int(self.getSubTotal() * self.iva_coeff)

    def getFlete(self):
        return self.flete

    def getReteFuente(self):
        return self.retefuente

    def getTotal(self):
        return self.getSubTotal() + self.getIVA() + self.getFlete() - self.getReteFuente()

    def getNumeroS(self):
        return str(self.getNumero()).zfill(4)

    def getSubTotalS(self):
        return self.formatComma(self.getSubTotal())

    def getIVAS(self):
        return self.formatComma(self.getIVA())

    def getFleteS(self):
        return self.formatComma(self.getFlete())

    def getReteFuenteS(self):
        return self.formatComma(self.getReteFuente())

    def getTotalS(self):
        return self.formatComma(self.getTotal())

    def formatComma(self, val):
        return "{:,}".format(val)

    def getServicio(self, cod):
        i = self.getCodigos().index(cod)
        return self.servicios[i]

    def getServicios(self):
        return self.servicios

    def getObservaciones(self):
        return self.observaciones

    def getNombre(self):
        return self.usuario.getNombre()

    def getDocumento(self):
        return self.usuario.getDocumento()

    def getDireccion(self):
        return self.usuario.getDireccion()

    def getCiudad(self):
        return self.usuario.getCiudad()

    def getTelefono(self):
        return self.usuario.getTelefono()

    def getCorreo(self):
        return self.usuario.getCorreo()

    def setNumero(self, numero):
        self.numero = numero

    def setUsuario(self, usuario):
        self.usuario = usuario

    def setServicios(self, servicios):
        codigos = [servicio.getCodigo() for servicio in servicios]
        if len(codigos) != len(set(codigos)):
            raise(Exception("Existe un código repetido."))
        else:
            self.servicios = servicios

    def setObservaciones(self, text):
        self.observaciones = text

    def setIvaCoeff(self, val):
        self.iva_coeff = val

    def setFlete(self, val):
        self.flete = val

    def setReteFuente(self, val):
        self.retefuente = val

    def removeServicio(self, index):
        del self.servicios[index]

    def addServicio(self, servicio):
        servicios = self.servicios + [servicio]
        self.setServicios(servicios)

    def addServicios(self, servicios):
        servicios = self.servicios + servicios
        self.setServicios(servicios)

    def makeTable(self):
        table = []
        for servicio in self.servicios:
            table += servicio.makeTable()
        return table

    def toRegistro(self, fields):
        global REGISTRO_DATAFRAME

        last = REGISTRO_DATAFRAME.shape[0]

        REGISTRO_DATAFRAME.loc[last] = fields

        REGISTRO_DATAFRAME = REGISTRO_DATAFRAME.drop_duplicates("Factura", "last")
        REGISTRO_DATAFRAME = REGISTRO_DATAFRAME.sort_values("Factura", ascending = False)
        REGISTRO_DATAFRAME = REGISTRO_DATAFRAME.reset_index(drop = True)

        writer = pd.ExcelWriter(constants.REGISTRO_FILE, engine='xlsxwriter',
                    datetime_format= "dd/mm/yy hh:mm")

        REGISTRO_DATAFRAME.to_excel(writer, index = False)

    def __repr__(self):
        l = (self.getNumeroS(), self.getNombre(), self.getDocumento(), self.getTotalS())
        return "Numero: %s\t\t Nombre: %s, Documento: %s, Total: %s"%l

class Factura(Documento):
    def __init__(self, numero = None, usuario = None, servicios = [], observaciones = "", iva = 0.19, flete = 0, retefuente = 0):
        super(Factura, self).__init__(numero, usuario, servicios, observaciones, iva, flete, retefuente)
        self.setPDFDir()

    def getPDFDir(self):
        return self.pdf_dir

    def setPDFDir(self, loc = None):
        if loc == None:
            self.pdf_dir =  os.path.join(constants.PDF_DIR, "Factura_" + self.getNumeroS())
        else:
            self.pdf_dir = loc

    def save(self):
        usuario = self.getUsuario()
        obs = self.getObservaciones()
        if obs == "":
            obs = "-"
        else:
            obs = obs.replace("\n", " ")
        fields = [self.getNumeroS(), datetime.now().replace(second = 0, microsecond = 0),
                    usuario.getNombre(), usuario.getDocumento(), usuario.getDireccion(),
                    usuario.getCiudad(), usuario.getTelefono(), usuario.getCorreo(),
                    self.getTotalS(), obs]

        self.usuario.save()
        self.toRegistro(fields)
        self.makePDF()

    def makePDF(self):
        FacturaPDFs(self)

    def fromDocumento(documento):
        d = documento
        return Factura(d.getNumero(), d.getUsuario(), d.getServicios(),
            d.getObservaciones(), d.getIVACoeff(), d.getFlete(), d.getReteFuente())

class Cotizacion(Documento):
    def __init__(self, numero = None, usuario = None, servicios = [], observaciones = "", iva = 0.19, flete = 0, retefuente = 0):
        super(Cotizacion, self).__init__(numero, usuario, servicios, observaciones, iva, flete, retefuente)
        self.setPDFDir()

    def getPDFDir(self):
        return self.pdf_dir

    def setPDFDir(self, loc = None):
        if loc == None:
            self.pdf_dir =  os.path.join(constants.PDF_DIR, "Cotizacion_" + self.getNumeroS() + ".pdf")
        else:
            self.pdf_dir = loc

    def save(self):
        usuario = self.getUsuario()
        obs = self.getObservaciones()
        if obs == "":
            obs = "-"
        else:
            obs = obs.replace("\n", " ")
        fields = ["C-" + self.getNumeroS(), datetime.now().replace(second = 0, microsecond = 0),
                    usuario.getNombre(), usuario.getDocumento(), usuario.getDireccion(),
                    usuario.getCiudad(), usuario.getTelefono(), usuario.getCorreo(),
                    self.getTotalS(), obs]

        self.usuario.save()
        self.toRegistro(fields)
        self.makePDF()
        with open(os.path.join(constants.OLD_DIR, self.getNumeroS() + ".pkl"), "wb") as file:
            pickle.dump(self, file)

    def makePDF(self):
        CotizacionPDF(self)

    def toDocumento(self):
        return Documento(self.getNumero(), self.getUsuario(), self.getServicios(),
            self.getObservaciones(), self.getIVA(), self.getFlete(), self.getReteFuente())

    def fromDocumento(documento):
        d = documento
        return Cotizacion(d.getNumero(), d.getUsuario(), d.getServicios(),
            d.getObservaciones(), d.getIVACoeff(), d.getFlete(), d.getReteFuente())

class Usuario(object):
    def __init__(self, nombre = None, documento = None, direccion = None, ciudad = None, telefono = None, correo = None):
        self.nombre = nombre
        self.documento = documento
        self.direccion = direccion
        self.ciudad = ciudad
        self.telefono = telefono
        self.correo = correo

    def getNombre(self):
        return self.nombre

    def getCorreo(self):
        return self.correo

    def getDocumento(self):
        return self.documento

    def getDireccion(self):
        return self.direccion

    def getCiudad(self):
        return self.ciudad

    def getTelefono(self):
        return self.telefono

    def setNombre(self, nombre):
        self.nombre = nombre

    def setCorreo(self, correo):
        self.correo = correo

    def setDocumento(self, documento):
        self.documento = documento

    def setDireccion(self, direccion):
        self.direccion = direccion

    def setCiudad(self, ciudad):
        self.ciudad = ciudad

    def setTelefono(self, telefono):
        self.telefono = telefono

    def save(self):
        global CLIENTES_DATAFRAME

        last = CLIENTES_DATAFRAME.shape[0]

        fields = []
        for key in CLIENTES_DATAFRAME.keys():
            key = unidecode(key)
            try:
                fields.append(eval("self.get%s()"%key))
            except: pass

        CLIENTES_DATAFRAME.loc[last] = fields
        CLIENTES_DATAFRAME = CLIENTES_DATAFRAME.drop_duplicates("Nombre", "last")
        CLIENTES_DATAFRAME = CLIENTES_DATAFRAME.sort_values("Nombre")
        CLIENTES_DATAFRAME.to_excel(constants.CLIENTES_FILE, index = False, na_rep = '')

class Servicio(object):
    def __init__(self, codigo = None, cantidad = None):
        self.codigo = codigo
        self.cantidad = cantidad

        self.valor_unitario = None
        self.valor_total = None
        self.descripcion = None
        self.setValorUnitario()
        self.setValorTotal()
        self.setDescripcion()

    def getCodigo(self):
        return self.codigo

    def getCantidad(self):
        return self.cantidad

    def getValorUnitario(self):
        return self.valor_unitario

    def getValorTotal(self):
        return self.valor_total

    def getDescripcion(self):
        return self.descripcion

    def setCodigo(self, codigo):
        self.codigo = codigo
        self.setDescripcion()
        self.setValorUnitario()
        self.setValorTotal()

    def setCantidad(self, cantidad):
        self.cantidad = cantidad
        self.setValorTotal()

    def setValorUnitario(self, valor = None):
        if valor == None:
            line = constants.PRECIOS_DATAFRAME[constants.PRECIOS_DATAFRAME["Código"] == self.codigo]
            if len(line) == 0: raise(Exception("Código inválido."))
            self.valor_unitario = int(line["Valor"].values[0])
        else:
            self.valor_unitario = valor
        self.setValorTotal()

    def setValorTotal(self, valor = None):
        if valor == None:
            self.valor_total = int(self.getValorUnitario() * self.getCantidad())
        else:
            self.valor_total = valor

    def setDescripcion(self, valor = None):
        if valor == None:
            line = constants.PRECIOS_DATAFRAME[constants.PRECIOS_DATAFRAME["Código"] == self.codigo]
            if len(line) == 0: raise(Exception("Código inválido."))
            self.descripcion = line["Descripción"].values[0]
        else:
            self.descripcion = valor

    def makeTable(self):
        desc = wrap(self.getDescripcion(), width = 80)
        table = []
        table.append(["%d"%self.getCantidad(), desc[0], "{:,}".format(self.getValorTotal())])
        desc.pop(0)
        for text in desc:
            table.append(["", text, ""])
        return table
