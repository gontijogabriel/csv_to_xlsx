from PyQt5.QtWidgets import QMainWindow, QApplication
from interface import Ui_MainWindow
import pandas as pd
import sys
import os


class GUI_cont(QMainWindow, Ui_MainWindow):

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        super().setupUi(self)
        self.comboBox_tipo.addItem('.csv       ->     .xlsx')
        self.comboBox_tipo.addItem('.xlsx      ->     .csv')
        self.label_console.setText('')
        self.pushButton_converter.clicked.connect(self.converter)
        self.pushButton_contato.clicked.connect(self.sobre)
        self.pushButton_desktop.clicked.connect(self.desktop)

    def sobre(self):
        os.startfile('sobre.txt')

    def desktop(self):
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        self.lineEdit_local_final.setText(str(desktop))

    def linkCreator(self, link):
        new = ''
        for x in range(len(link)):
            if link[x] == '\\':
                new = new + link[x] + '\\'
            else:
                new = new + link[x]

        return new

    def converter(self):
        self.label_console.setText('Convertendo...')
        try:
            if self.comboBox_tipo.currentText() == '.csv       ->     .xlsx' and self.lineEdit_local.text() != '' and self.lineEdit_local.text() != '':
                self.label_console.setText('CSV TO XLSX')
                caminho = self.linkCreator(self.lineEdit_local.text())
                local_final = self.linkCreator(self.lineEdit_local_final.text())
                nome_arquivo = self.lineEdit.text()

                df = pd.read_csv(f'{caminho}')

                df.to_excel(f'{local_final}\\{nome_arquivo}.xlsx', index=False)
                self.label_console.setText('Sucesso')

            elif self.comboBox_tipo.currentText() == '.xlsx      ->     .csv' and self.lineEdit_local.text() != '' and self.lineEdit_local.text() != '':
                self.label_console.setText('XLSX TO CSV')
                caminho = self.linkCreator(self.lineEdit_local.text())
                local_final = self.linkCreator(self.lineEdit_local_final.text())
                nome_arquivo = self.lineEdit.text()

                df = pd.read_excel(f'{caminho}')

                df.to_csv(f'{local_final}\\{nome_arquivo}.csv', index=False)
                self.label_console.setText('Sucesso')

            else:
                self.label_console.setText('erro, algum campo em branco')

        except ValueError:
            self.label_console.setText(f'ERRO: {str(ValueError)}')


if __name__ == "__main__":
    qt = QApplication(sys.argv)
    view = GUI_cont()
    view.show()
    qt.exec()