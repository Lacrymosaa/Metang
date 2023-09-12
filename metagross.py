import sys
import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QFileDialog, QLabel, QMessageBox
from PyQt5.QtGui import QPixmap, QIcon
import pandas as pd

class Metagross(QMainWindow):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle('Metagross')
        self.setWindowIcon(QIcon("metagross.ico"))
        self.setFixedSize(300, 100)

        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()

        self.load_button = QPushButton('Carregar Arquivo Json', self)
        self.load_button.clicked.connect(self.loadJson)

        self.convert_button = QPushButton('Converter para XLSX', self)
        self.convert_button.clicked.connect(self.convert)
        self.convert_button.setEnabled(False)

        self.layout.addWidget(self.load_button)
        self.layout.addWidget(self.convert_button)

        self.image_label = QLabel(self)
        self.layout.addWidget(self.image_label)

        self.central_widget.setLayout(self.layout)

    def loadJson(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        fileName, _ = QFileDialog.getOpenFileName(self, 'Carregar Arquivo Json', '', 'json (*.json)', options=options)

        if fileName:
            self.json_file = fileName
            self.convert_button.setEnabled(True)

    def showImage(self, fileName):
        pixmap = QPixmap(fileName)
        self.image_label.setPixmap(pixmap)
        self.image_label.show()

    def convert(self):
        try:
            base_name = os.path.splitext(os.path.basename(self.json_file))[0]
            xlsx_file = f'{base_name}.xlsx'
            
            data = pd.read_json(self.json_file)

            data.to_excel(xlsx_file, index = None)
            
            print(f'Arquivo CSV "{xlsx_file}" gerado com sucesso!')
        except Exception as e:
            print(f'Ocorreu um erro ao converter o JSON para CSV: {str(e)}')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Metagross()
    window.show()
    sys.exit(app.exec_())