import os
import pandas as pd

IVA_COEFF = 0.19

OLD_DIR = "Cotizaciones"
PDF_DIR = "PDF"
REGISTERS_DIR = "Registers"

PDF_PRINT = "_print.pdf"
PDF_DIGITAL = "_digital.pdf"

CLIENTES_FILE = "Clientes.xlsx"
REGISTRO_FILE = "Registro.xlsx"
PRECIOS_FILE = "Precios.xlsx"

CLIENTES_FILE = os.path.join(REGISTERS_DIR, CLIENTES_FILE)
REGISTRO_FILE = os.path.join(REGISTERS_DIR, REGISTRO_FILE)

PRECIOS_FILE = os.path.join(REGISTERS_DIR, PRECIOS_FILE)

PRECIOS_KEYS = ['Código', 'Descripción', 'Valor']
REGISTRO_KEYS = ['Factura', 'Fecha', 'Nombre', 'Documento', 'Dirección', 'Ciudad', 'Teléfono', 'Correo', 'Valor', 'Observaciones']
CLIENTES_KEYS = ['Nombre', 'Documento', 'Dirección', 'Ciudad', 'Teléfono', 'Correo']

PRECIOS_DATAFRAME = pd.read_excel(PRECIOS_FILE).fillna("").astype(str)

BACKGROUND_FACTURA = "Registers/factura_tausand.jpg"
BACKGROUND_COTIZACION = "Registers/cotizacion_tausand.jpg"
