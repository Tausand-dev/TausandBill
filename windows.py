import os
import sys
import numpy as np
from time import sleep
from datetime import datetime
from PyQt5 import QtCore, QtWidgets, QtGui

import objects
import constants

import psutil
from subprocess import Popen
from threading import Thread

class PandasModel(QtCore.QAbstractTableModel):
    def __init__(self, data, parent = None, checkbox = True):
        QtCore.QAbstractTableModel.__init__(self, parent)
        self.dataframe = data
        self._data = data.values
        self.checkbox = checkbox

        if self.checkbox:
            temp = []
            for i in range(self._data.shape[0]):
                c = QtWidgets.QCheckBox()
                c.setChecked(True)
                temp.append(c)

            new = np.zeros((self._data.shape[0], self._data.shape[1] + 1), dtype = object)
            new[:, 0] = temp
            new[:, 1:] = self._data

            self._data = new

            self.headerdata = ["Guardar"] + list(data.keys())
        else:
            self.headerdata = list(data.keys())

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role = QtCore.Qt.DisplayRole):
        if not index.isValid():
            return None
        if (index.column() == 0 and self.checkbox):
            value = self._data[index.row(), index.column()].text()
        else:
            value = self._data[index.row(), index.column()]
        if role == QtCore.Qt.EditRole:
            return value
        elif role == QtCore.Qt.DisplayRole:
            return value
        elif role == QtCore.Qt.CheckStateRole:
            if index.column() == 0 and self.checkbox:
                if self._data[index.row(), index.column()].isChecked():
                    return QtCore.Qt.Checked
                else:
                    return QtCore.Qt.Unchecked

    def flags(self, index):
        if not index.isValid():
            return None
        if index.column() == 0 and self.checkbox:
            return QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsUserCheckable
        else:
            return QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable

    def headerData(self, col, orientation, role):
        if orientation == QtCore.Qt.Horizontal and role == QtCore.Qt.DisplayRole:
            return QtCore.QVariant(self.headerdata[col])

        return QtCore.QVariant()

    def setData(self, index, value, role):
        if not index.isValid():
            return False
        if role == QtCore.Qt.CheckStateRole and index.column() == 0:
            if value == QtCore.Qt.Checked:
                self._data[index.row(), index.column()].setChecked(True)
            else:
                self._data[index.row(), index.column()].setChecked(False)

        self.dataChanged.emit(index, index)
        return True

    def whereIsChecked(self):
        return np.array([self._data[i, 0].isChecked() for i in range(self.rowCount())], dtype = bool)

class Table(QtWidgets.QTableWidget):
    HEADER = ['Código', 'Descripción', 'Cantidad', 'Valor Unitario', 'Valor Total']
    def __init__(self, parent, rows = 9, cols = 5):
        super(Table, self).__init__(rows, cols)

        self.parent = parent
        self.n_rows = rows
        self.n_cols = cols
        self.setHorizontalHeaderLabels(self.HEADER)
        self.resizeRowsToContents()

        header = self.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QtWidgets.QHeaderView.ResizeToContents)
        self.clean()

        self.cellChanged.connect(self.handler)

    def removeServicio(self):
        table_codigos = self.getCodigos()
        regis_codigos = self.parent.getCodigos()
        for (i, codigo) in enumerate(regis_codigos):
            if not codigo in table_codigos: self.parent.removeServicio(i)

    def handler(self, row, col):
        self.blockSignals(True)
        item = self.item(row, col)
        try:
            cod = self.item(row, 0).text()
            if col == 0:
                self.removeServicio()
                if cod == "":
                    self.item(row, 1).setText("")
                    self.item(row, 2).setText("")
                    self.item(row, 3).setText("")
                    self.item(row, 4).setText("")

                else:
                    try:
                        n = int(self.item(row, 2).text())
                        total = 0
                    except:
                        try:
                            total = int(self.item(row, 4).text().replace(",", ""))
                            n = 1
                        except:
                            n = 1
                            total = 0
                            self.item(row, 2).setText("1")

                    try:
                        servicio = objects.Servicio(codigo = cod, cantidad = n)
                        self.parent.addServicio(servicio)
                    except:
                        self.item(row, 0).setText("")
                        self.item(row, 1).setText("")
                        self.item(row, 2).setText("")
                        self.item(row, 3).setText("")
                        self.item(row, 4).setText("")
                        raise(Exception("Código inválido."))

                    desc = servicio.getDescripcion()
                    valor = servicio.getValorUnitario()
                    total = servicio.getValorTotal()
                    self.item(row, 1).setText(desc)
                    self.item(row, 3).setText("{:,}".format(valor))
                    self.item(row, 4).setText("{:,}".format(total))

            if col == 2:
                try: n = int(self.item(row, 2).text())
                except: raise(Exception("Cantidad inválida.")); self.item(row, 2).setText("")

                self.item(row, 2).setText("%d"%n)

                if cod != "":
                    servicio = self.parent.getServicio(cod)
                    servicio.setCantidad(n)
                    total = servicio.getValorTotal()

                    self.item(row, 4).setText("{:,}".format(total))

            self.parent.setTotal()
        except Exception as e:
            self.parent.errorWindow(e)
        self.blockSignals(False)

    def getCodigos(self):
        return [self.item(i, 0).text() for i in range(self.n_rows)]

    def clean(self):
        self.blockSignals(True)
        for r in range(self.n_rows):
            for c in range(self.n_cols):
                item = QtWidgets.QTableWidgetItem("")
                self.setItem(r, c, item)
        self.readOnly()
        self.blockSignals(False)

    def readOnly(self):
        flags = QtCore.Qt.ItemIsEditable
        for r in range(self.n_rows):
            for c in [1, 3, 4]:
                item = QtWidgets.QTableWidgetItem("")
                item.setFlags(flags)
                self.setItem(r, c, item)

class AutoLineEdit(QtWidgets.QLineEdit):
    AUTOCOMPLETE = ["Nombre", "Correo", "Documento", "Teléfono", "Cotización"]
    def __init__(self, target, parent, autochange = True):
        super(AutoLineEdit, self).__init__()
        self.target = target
        self.parent = parent
        self.model = QtCore.QStringListModel()
        completer = QtWidgets.QCompleter()
        completer.setCaseSensitivity(False)
        completer.setModelSorting(0)
        completer.setModel(self.model)
        self.setCompleter(completer)
        self.update()

    def event(self, event):
        if event.type() == QtCore.QEvent.KeyPress and event.key() == QtCore.Qt.Key_Tab:
            try:
                self.parent.changeAutocompletar()
            except:
                pass
            return False
        return QtWidgets.QWidget.event(self, event)

    def update(self):
        if type(self.parent) is FacturaWindow:
            dataframe = objects.CLIENTES_DATAFRAME
            order = 1
        else:
            dataframe = objects.REGISTRO_DATAFRAME
            order = -1
        data = list(set(dataframe[self.target].values.astype('str')))
        data = sorted(data)[::order]
        self.model.setStringList(data)

class CodigosDialog(QtWidgets.QDialog):
    def __init__(self):
        super(CodigosDialog, self).__init__()
        self.setWindowTitle("Ver códigos")
        self.layout = QtWidgets.QVBoxLayout()
        self.setLayout(self.layout)

        self.table = QtWidgets.QTableView()
        self.layout.addWidget(self.table)

    def setModel(self, df):
        self.table.setModel(PandasModel(df, checkbox = False))

        self.table.setSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)

        self.table.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.table.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)

        self.table.resizeColumnsToContents()
        self.table.setFixedSize(self.table.horizontalHeader().length() + self.table.verticalHeader().width(),
                                self.table.verticalHeader().length() + self.table.horizontalHeader().height())

        self.resize(self.table.sizeHint())

class FacturaWindow(QtWidgets.QMainWindow):
    FIELDS = ["Nombre", "Documento", "Dirección", "Ciudad", "Teléfono", "Correo"]
    WIDGETS = ["nombre", "documento", "direccion", "ciudad", "telefono", "correo"]
    AUTOCOMPLETE_FIELDS = ["Nombre", "Correo", "Documento", "Teléfono"]
    AUTOCOMPLETE_WIDGETS = ["nombre", "correo", "documento", "telefono"]
    def __init__(self, parent = None):
        super(QtWidgets.QMainWindow, self).__init__(parent)
        self.setWindowTitle("Factura")

        wid = QtWidgets.QWidget(self)
        self.setCentralWidget(wid)

        self.is_closed = True
        self.ver_dialog = CodigosDialog()
        self.verticalLayout = QtWidgets.QVBoxLayout(wid)

        self.verticalLayout.setContentsMargins(11, 11, 11, 11)
        self.verticalLayout.setSpacing(6)

        self.factura_frame = QtWidgets.QFrame()
        self.form_frame = QtWidgets.QFrame()
        self.button_frame = QtWidgets.QFrame()
        self.total_frame = QtWidgets.QFrame()

        self.factura_frame_layout = QtWidgets.QHBoxLayout(self.factura_frame)
        self.numero_factura = QtWidgets.QPushButton()
        self.factura_frame_layout.addWidget(self.numero_factura)

        self.form_frame_layout = QtWidgets.QGridLayout(self.form_frame)
        self.form_frame_layout.setContentsMargins(0, 0, 0, 0)
        self.form_frame_layout.setSpacing(6)

        self.autocompletar_widget = QtWidgets.QCheckBox("Autocompletar")
        self.autocompletar_widget.setChecked(True)

        nombre_label = QtWidgets.QLabel("Nombre:")
        self.nombre_widget = AutoLineEdit("Nombre", self)
        correo_label = QtWidgets.QLabel("Correo:")
        self.correo_widget = AutoLineEdit("Correo", self)

        documento_label = QtWidgets.QLabel("Nit/CC:")
        self.documento_widget = AutoLineEdit("Documento", self)
        direccion_label = QtWidgets.QLabel("Dirección:")
        self.direccion_widget = AutoLineEdit("Dirección", self)
        ciudad_label = QtWidgets.QLabel("Ciudad:")
        self.ciudad_widget = AutoLineEdit("Ciudad", self)
        telefono_label = QtWidgets.QLabel("Teléfono:")
        self.telefono_widget = AutoLineEdit("Teléfono", self)

        observaciones_label = QtWidgets.QLabel("Observaciones")
        self.observaciones_widget = QtWidgets.QTextEdit("")
        self.observaciones_widget.setFixedHeight(66)

        self.form_frame_layout.addWidget(nombre_label, 0, 0)
        self.form_frame_layout.addWidget(self.nombre_widget, 0, 1)
        self.form_frame_layout.addWidget(documento_label, 0, 2)
        self.form_frame_layout.addWidget(self.documento_widget, 0, 3)

        self.form_frame_layout.addWidget(direccion_label, 2, 0)
        self.form_frame_layout.addWidget(self.direccion_widget, 2, 1)
        self.form_frame_layout.addWidget(ciudad_label, 2, 2)
        self.form_frame_layout.addWidget(self.ciudad_widget, 2, 3)

        self.form_frame_layout.addWidget(telefono_label, 3, 0)
        self.form_frame_layout.addWidget(self.telefono_widget, 3, 1)
        self.form_frame_layout.addWidget(correo_label, 3, 2)
        self.form_frame_layout.addWidget(self.correo_widget, 3, 3)

        self.table = Table(self)

        self.button_frame_layout = QtWidgets.QHBoxLayout(self.button_frame)
        self.view_button = QtWidgets.QPushButton("Ver códigos")
        self.guardar_button = QtWidgets.QPushButton("Guardar")
        self.limpiar_button = QtWidgets.QPushButton("Limpiar")

        self.button_frame_layout.addWidget(self.guardar_button)
        self.button_frame_layout.addWidget(self.limpiar_button)
        self.button_frame_layout.addWidget(self.view_button)

        self.total_frame_layout = QtWidgets.QFormLayout(self.total_frame)
        subtotal_label = QtWidgets.QLabel("Sub-total:")
        self.subtotal_widget = QtWidgets.QLabel()
        iva_label = QtWidgets.QLabel("IVA:")
        self.iva_widget = QtWidgets.QLabel()
        flete_label = QtWidgets.QLabel("Flete:")
        self.flete_widget = QtWidgets.QLabel()
        retefuente_label = QtWidgets.QLabel("ReteFuente:")
        self.retefuente_widget = QtWidgets.QLabel()
        total_label = QtWidgets.QLabel("Total:")
        self.total_widget = QtWidgets.QLabel()

        subtotal_label.setAlignment(QtCore.Qt.AlignRight)
        self.subtotal_widget.setAlignment(QtCore.Qt.AlignRight)
        iva_label.setAlignment(QtCore.Qt.AlignRight)
        self.iva_widget.setAlignment(QtCore.Qt.AlignRight)
        flete_label.setAlignment(QtCore.Qt.AlignRight)
        self.flete_widget.setAlignment(QtCore.Qt.AlignRight)
        retefuente_label.setAlignment(QtCore.Qt.AlignRight)
        self.retefuente_widget.setAlignment(QtCore.Qt.AlignRight)
        total_label.setAlignment(QtCore.Qt.AlignRight)
        self.total_widget.setAlignment(QtCore.Qt.AlignRight)

        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        self.factura_frame.setSizePolicy(sizePolicy)
        self.total_frame.setSizePolicy(sizePolicy)
        self.button_frame.setSizePolicy(sizePolicy)

        self.factura_frame_layout.setAlignment(QtCore.Qt.AlignRight)
        self.total_frame_layout.setAlignment(QtCore.Qt.AlignRight)
        self.button_frame_layout.setAlignment(QtCore.Qt.AlignRight)
        self.verticalLayout.setAlignment(QtCore.Qt.AlignRight)

        self.total_frame_layout.addRow(subtotal_label, self.subtotal_widget)
        self.total_frame_layout.addRow(iva_label, self.iva_widget)
        self.total_frame_layout.addRow(flete_label, self.flete_widget)
        self.total_frame_layout.addRow(retefuente_label, self.retefuente_widget)
        self.total_frame_layout.addRow(total_label, self.total_widget)
        self.total_frame_layout.addRow(observaciones_label, self.observaciones_widget)

        self.verticalLayout.addWidget(self.factura_frame)
        self.verticalLayout.addWidget(self.autocompletar_widget)
        self.verticalLayout.addWidget(self.form_frame)
        self.verticalLayout.addWidget(self.table)
        self.verticalLayout.addWidget(self.total_frame)
        self.verticalLayout.addWidget(self.button_frame)

        self.setAutoCompletar()

        self.factura = objects.Factura()

        self.setLastFactura()

        self.resize(600, 550)

        self.limpiar_button.clicked.connect(self.limpiar)
        self.guardar_button.clicked.connect(self.guardar)
        self.view_button.clicked.connect(self.verCodigos)

    def setAutoCompletar(self):
        for item in self.AUTOCOMPLETE_WIDGETS:
            widget = eval("self.%s_widget"%item)
            widget.textChanged.connect(self.autoCompletar)
            widget.returnPressed.connect(self.changeAutocompletar)

    def changeAutocompletar(self):
        self.autocompletar_widget.setChecked(False)

    def updateAutoCompletar(self):
        for item in self.AUTOCOMPLETE_WIDGETS:
            exec("self.%s_widget.update()"%item)

    def autoCompletar(self, text):
        if text != "":
            df = objects.CLIENTES_DATAFRAME[self.AUTOCOMPLETE_FIELDS]
            booleans = df.isin([text]).values.sum(axis = 1)
            pos = np.where(booleans)[0]
            cliente = objects.CLIENTES_DATAFRAME.iloc[pos]

            if len(pos):
                if self.autocompletar_widget.isChecked():
                    for (field, widgetT) in zip(self.FIELDS, self.WIDGETS):
                        if field in objects.CLIENTES_DATAFRAME.keys():
                            val = str(cliente[field].values[0])
                            if val == "nan": val = ""
                            widget = eval("self.%s_widget"%widgetT)
                            widget.blockSignals(True)
                            widget.setText(val)
                            widget.blockSignals(False)

    def limpiar(self):
        self.table.clean()
        for field in self.WIDGETS:
            widget = eval("self.%s_widget"%field)
            widget.blockSignals(True)
            widget.setText("")
            widget.blockSignals(False)
        self.setLastFactura()
        self.factura.setServicios([])
        self.subtotal_widget.setText("")
        self.iva_widget.setText("")
        self.flete_widget.setText("")
        self.retefuente_widget.setText("")
        self.total_widget.setText("")
        self.observaciones_widget.setText("")
        self.autocompletar_widget.setChecked(True)

    def verCodigos(self):
        self.ver_dialog.setModel(constants.PRECIOS_DATAFRAME)
        self.ver_dialog.show()

    def closePDF(self, p1, old):
        new = [proc.pid for proc in psutil.process_iter()]
        try:
            new.remove(p1.pid)
        except:
            return None

        new = [proc for proc in new if proc not in old]
        try:
            for proc in new:
                p = psutil.Process(proc)
                if p.parent().name() == "cmd.exe":
                    if (p.parent().parent().name() == "MicroBill.exe") or (p.parent().parent().name() == "python.exe"):
                        p.terminate()
            p1.kill()
        except:
            pass

    def confirmGuardar(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setText("¿Está seguro que desea guardar esta cotización?.\nVerifique los datos.")
        msg.setWindowTitle("Confirmar")
        msg.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.Cancel)
        ans = msg.exec_()
        if ans == QtWidgets.QMessageBox.Yes:
            return True
        return False

    def guardar(self):
        try:
            dic = {}
            for key in self.WIDGETS:
                value = eval("self.%s_widget.text()"%key)
                if value == "":
                    raise(Exception("Existen campos sin llenar en la información del usuario."))
                dic[key] = value

            if len(self.getServicios()) == 0:
                raise(Exception("No existen servicios cotizados."))

            usuario = objects.Usuario(**dic)
            self.factura.setUsuario(usuario)
            self.factura.setObservaciones(self.observaciones_widget.toPlainText())
            self.factura.makePDF()

            path = os.path.dirname(sys.executable)
            path = os.path.join(path, self.factura.pdf_dir)
            if not os.path.exists(path):
                path = os.path.join(os.getcwd(), self.factura.pdf_dir)

            path1 = os.path.join(path, constants.PDF_PRINT)
            path2 = os.path.join(path, constants.PDF_DIGITAL)

            old = [proc.pid for proc in psutil.process_iter()]

            p1 = Popen(path2, shell = True)

            if self.confirmGuardar():
                self.closePDF(p1, old)
                for i in range(10):
                    try:
                        self.factura.save()
                        break
                    except PermissionError:
                        sleep(0.1)

                self.updateAutoCompletar()
                self.limpiar()
                self.setLastFactura()
            else:
                self.closePDF(p1, old)
                self.factura.setUsuario(None)
                self.factura.setObservaciones("")
                os.remove(path1)
                for i in range(10):
                    try:
                        os.remove(path2)
                        break
                    except PermissionError:
                        sleep(0.1)

        except Exception as e:
            self.errorWindow(e)

        self.factura.setUsuario(None)
        self.factura.setObservaciones("")

    def setLastFactura(self):
        try:
            cot = int(objects.REGISTRO_DATAFRAME["Factura"].values[0]) + 1
        except IndexError:
            cot = 1

        self.numero_factura.setText(str(cot))
        self.factura.setNumero(cot)

    def centerOnScreen(self):
        resolution = QtWidgets.QDesktopWidget().screenGeometry()
        self.move((resolution.width() / 2) - (self.frameSize().width() / 2),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))

    def addServicio(self, servicio):
        self.factura.addServicio(servicio)

    def getCodigos(self):
        return self.factura.getCodigos()

    def getServicio(self, cod):
        return self.factura.getServicio(cod)

    def getServicios(self):
        return self.factura.getServicios()

    def removeServicio(self, index):
        self.factura.removeServicio(index)

    def setTotal(self):
        self.subtotal_widget.setText(self.factura.getSubTotalS())
        self.iva_widget.setText(self.factura.getIVAS())
        self.flete_widget.setText(self.factura.getFleteS())
        self.retefuente_widget.setText(self.factura.getReteFuenteS())
        self.total_widget.setText(self.factura.getTotalS())

    def errorWindow(self, exception):
        error_text = str(exception)
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Warning)
        msg.setText(error_text)
        msg.setWindowTitle("Error")
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.exec_()

    def closeEvent(self, event):
        self.is_closed = True
        event.accept()
