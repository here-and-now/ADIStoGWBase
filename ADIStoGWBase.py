from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
from PyQt5 import uic
import pandas as pd
from datetime import datetime
import xlrd
import sys
import os



class ADIS(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()

        if getattr(sys, 'frozen', False):
            # frozen
            dirname = os.path.dirname(sys.executable)
        else:
            # unfrozen
            dirname = os.path.dirname(os.path.realpath(__file__))

        self.ui = uic.loadUi('Einstellungen/ADIS_Import.ui', self)

        self.inputFileLine = self.inputFile
        self.outputFileLine = self.outputFile
        self.log = self.logTextEdit

        self.missingParams = []
        date = datetime.today().strftime('%d-%m-%Y')

        self.outputFileLine.setText(dirname + '\\' + 'Output_für_GWBase\\' + "ADIS_GWBase_"+ str(date)+ '.xlsx')

    def openImportFileNameDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self, "QFileDialog.getOpenFileName()", "",
                                                  "All Files (*);;Python Files (*.py)", options=options)

        self.inputFileLine.setText(fileName)

    def openExportFileNameDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self, "QFileDialog.getOpenFileName()", "",
                                                  "All Files (*);;Python Files (*.py)", options=options)

        self.outputFileLine.setText(fileName)

    def ADIStoExcel(self):
        if self.inputFileLine:
            try:

                df = pd.read_table(self.inputFileLine.text(), skiprows=1, encoding="ISO-8859-1", header=None)

                df_param = pd.read_excel("Einstellungen\Parameter_Zuweisung.xlsx", index_col=0)
                df_mest = pd.read_excel("Einstellungen\Messstellen_Zuweisung.xlsx", index_col=0)

                dict_param = df_param.to_dict()
                dict_mest = df_mest.to_dict()

                df = df[df[0] == "PPA"]

                df[6] = df[6].fillna('')
                df[7] = df[6].map(str) + df[7]
                df = df.replace({'n.b.': 'n.n.'})

                df_pivot = df[[1, 4, 7]]

                for param in df[4]:
                    if param not in dict_param['GWBase']:
                        self.missingParams.append(param)


                df_datetime = df[[1, 2]]
                df_datetime = df_datetime.rename(columns={1: "Messstelle", 2: "Datum/Zeit"})
                df_datetime = df_datetime.drop_duplicates()


                df_pivot = df.pivot(index=1, columns=4, values=7).reset_index()
                df_pivot = df_pivot.rename(columns={1: "Messstelle"})
                df_pivot = df_pivot.rename(columns=dict_param['GWBase'])

                df_pivot = df_pivot.join(df_datetime.set_index('Messstelle'), on='Messstelle', how='left')
                df_pivot = df_pivot.set_index('Datum/Zeit').reset_index()
                df_pivot['Datum/Zeit'] = pd.to_datetime(df_pivot['Datum/Zeit'])

                def excel_date(date1):
                    temp = datetime(1899, 12, 30)
                    delta = date1 - temp
                    return float(delta.days) + (float(delta.seconds) / 86400)

                df_pivot['Datum/Zeit'] = df_pivot['Datum/Zeit'].map(excel_date)
                df_pivot = df_pivot.replace({"Messstelle": dict_mest['GWBase']})
                df_pivot.set_index('Messstelle', inplace=True)

                try:
                    df_pivot.to_excel(self.outputFileLine.text())
                except:
                    self.log.append("Excel Zugriffsfehler")

                if self.missingParams:
                    self.log.append("Warnung: für folgende ADIS Parameter existiert keine Zuordnung nach GWBase, bitte unter Parameter Zuweisung hinzufügen")
                    for param in self.missingParams:
                        self.log.append(param)
                    self.log.append("Importdatei erstellt, eventuell Probleme beim GWBase Import möglich")
                    # self.log.append("--------------------------------------------------")

                if not self.missingParams:
                    self.log.append("Importdatei erfolgreich erstellt")

                self.openExcel.setEnabled(True)

            except:
                self.log.append("Keine gültige ADIS Datei")


    def paramZuweisung(self):
        os.startfile('Einstellungen\Parameter_Zuweisung.xlsx')

    def mestZuweisung(self):
        os.startfile('Einstellungen\Messstellen_Zuweisung.xlsx')


    def openExcelFile(self):
        os.startfile(self.outputFileLine.text())

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ADIS()
    ex.importFileButton.clicked.connect(ex.openImportFileNameDialog)
    ex.exportFileButton.clicked.connect(ex.openExportFileNameDialog)

    ex.doImportButton.clicked.connect(ex.ADIStoExcel)
    ex.paramButton.clicked.connect(ex.paramZuweisung)
    ex.mestButton.clicked.connect(ex.mestZuweisung)

    ex.openExcel.clicked.connect(ex.openExcelFile)

    ex.show()
    sys.exit(app.exec_())