import sys
from PyQt5 import QtCore, QtGui, QtWidgets

app = QtWidgets.QApplication(sys.argv)

icon = QtGui.QIcon('Registers/icon.ico')
app.setWindowIcon(icon)

splash_pix = QtGui.QPixmap('Registers/logo.png').scaledToWidth(500)
splash = QtWidgets.QSplashScreen(splash_pix, QtCore.Qt.WindowStaysOnTopHint)
splash.show()
app.processEvents()

def testFiles():
    """
        Registro y clientes
    """
    if (list(REGISTRO_DATAFRAME.keys()) != constants.REGISTRO_KEYS):
        fields = ", ".join(constants.REGISTRO_KEYS)
        txt = "Registro no está bien configurado, las columnas deben ser: %s."%fields
        raise(Exception(txt))

    if (list(CLIENTES_DATAFRAME.keys()) != constants.CLIENTES_KEYS):
        fields = ", ".join(constants.CLIENTES_KEYS)
        txt = "Clientes no está bien configurado, las columnas deben ser: %s."%fields
        raise(Exception(txt))
    """
        Equipo
    """
    if (list(constants.PRECIOS_DATAFRAME.keys()) != constants.PRECIOS_KEYS):
        fields = ", ".join(constants.PRECIOS_KEYS)
        txt = "Precios.xlsx no está bien configurado, las columnas deben ser: %s."%fields
        raise(Exception(txt))

try:
    import os
    import constants
    from time import sleep
    from windows import FacturaWindow
    from objects import REGISTRO_DATAFRAME, CLIENTES_DATAFRAME
    testFiles()

except Exception as e:
    msg = QtWidgets.QMessageBox()
    msg.setIcon(QtWidgets.QMessageBox.Warning)
    msg.setText("Error de configuración inicial:\n%s"%str(e))
    msg.setWindowTitle("Error")
    msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
    msg.exec_()
    sys.exit()

main = FacturaWindow()

QtWidgets.QApplication.setStyle(QtWidgets.QStyleFactory.create('Fusion'))

main.setWindowIcon(icon)
splash.close()
main.show()

sys.exit(app.exec_())
