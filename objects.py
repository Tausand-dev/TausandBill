import os
import pickle
import numpy as np
import pandas as pd
from datetime import datetime

import constants
from pdflib import FacturaPDFs
from unidecode import unidecode

def makeDir(dir):
    if os.path.isdir(dir): pass
    else: os.makedirs(dir)

makeDir(constants.PDF_DIR)

# if os.path.isdir(constants.OLD_DIR): pass
# else: os.makedirs(constants.OLD_DIR)

def readDataFrames():
    c = pd.read_excel(constants.CLIENTES_FILE).fillna("").astype(str)
    r = pd.read_excel(constants.REGISTRO_FILE).fillna("").astype(str)
    return c, r

CLIENTES_DATAFRAME, REGISTRO_DATAFRAME = readDataFrames()

class Factura(object):
    def __init__(self, numero = None, usuario = None, servicios = [], observaciones = ""):
        self.numero = numero
        self.usuario = usuario
        self.observaciones = observaciones
        self.pdf_dir = ""
        self.setServicios(servicios)

    def getUsuario(self):
        return self.usuario

    def getCodigos(self):
        return [servicio.getCodigo() for servicio in self.servicios]

    def getNumero(self):
        return self.numero

    def getSubTotal(self):
        return sum([servicio.getValorTotal() for servicio in self.servicios])

    def getIVA(self):
        return int(self.getSubTotal() * constants.IVA_COEFF)

    def getFlete(self):
        return 0

    def getReteFuente(self):
        return 0

    def getTotal(self):
        return self.getSubTotal() + self.getIVA() + self.getFlete() + self.getReteFuente()

    def getNumeroS(self):
        return str(self.getNumero())

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

    def setFileName(self, name):
        self.pdf_file_name = name

    def setObservaciones(self, text):
        self.observaciones = text

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
            row = servicio.makeTable()
            table.append(row)
        return table

    def save(self):
        self.usuario.save()
        self.toRegistro()
        self.makePDF()

    def makePDF(self):
        FacturaPDFs(self)

    def toRegistro(self):
        global REGISTRO_DATAFRAME

        last = REGISTRO_DATAFRAME.shape[0]

        usuario = self.getUsuario()
        fields = [self.getNumeroS(), datetime.now().replace(second = 0, microsecond = 0),
                    usuario.getNombre(), usuario.getDocumento(), usuario.getDireccion(),
                    usuario.getCiudad(), usuario.getTelefono(), usuario.getCorreo(),
                    self.getTotalS(), self.getObservaciones()]

        REGISTRO_DATAFRAME.loc[last] = fields

        REGISTRO_DATAFRAME = REGISTRO_DATAFRAME.drop_duplicates("Factura", "last")
        REGISTRO_DATAFRAME = REGISTRO_DATAFRAME.sort_values("Factura", ascending = False)
        REGISTRO_DATAFRAME = REGISTRO_DATAFRAME.reset_index(drop = True)

        writer = pd.ExcelWriter(constants.REGISTRO_FILE, engine='xlsxwriter',
                    datetime_format= "dd/mm/yy hh:mm")
        REGISTRO_DATAFRAME.to_excel(writer, index = False)

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
        return ["%d"%self.getCantidad(), self.getDescripcion(), "{:,}".format(self.getValorTotal())]
