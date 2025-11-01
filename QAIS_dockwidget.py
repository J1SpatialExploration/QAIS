from qgis.PyQt import QtWidgets, uic
import os

FORM_CLASS, _ = uic.loadUiType(os.path.join(
    os.path.dirname(__file__), 'QAIS_dockwidget_base.ui'))

class QAISDockWidget(QtWidgets.QDockWidget, FORM_CLASS):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)