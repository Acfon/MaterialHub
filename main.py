import sys
import pandas as pd

from PyQt6 import uic
from PyQt6.QtWidgets import (QApplication, QMainWindow, QFileDialog, QTableWidgetItem, QHeaderView, QMenu, QDialog, \
                             QListWidget, QScrollArea, QMessageBox, QVBoxLayout)
from PyQt6.QtCore import QSignalMapper, Qt, QDateTime
from PyQt6.QtGui import QIcon


class Dobav(QDialog):
    def __init__(self, fname, data):
        super().__init__()
        uic.loadUi('form.ui', self)
        self.setWindowIcon(QIcon('logo.png'))
        self.setWindowTitle('Добавление нового материала')
        self.date.setCalendarPopup(True)
        self.date.setDateTime(QDateTime.currentDateTime())
        self.name = fname
        self.data = data
        self.doba.clicked.connect(self.dobZap)

    def dobZap(self):
        df2 = pd.DataFrame({"Дата": [self.date.text()],
                            "Вид материала": [self.material.text()],
                            "Размер катушки, вес кг.": [self.size.text()],
                            "Сечение": [self.doub.text()],
                            "Цвет": [self.color.text()],
                            "Условия хранения": [self.uslovia.text()],
                            "Статус": [self.status.currentText()],
                            "Остаток": [self.ostatok.text()]})
        new_data = pd.concat([self.data, df2])
        new_data.to_excel(self.name, index=False)
        self.close()


class AccountingSystem(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('design.ui', self)
        self.zagruz.clicked.connect(self.run)
        self.export_2.clicked.connect(self.export_to_xlsx)
        self.nofilt.clicked.connect(self.unfilter)
        self.horizontalHeader = self.tableWidget.horizontalHeader()
        self.horizontalHeader.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.horizontalHeader.sectionClicked.connect(self.header_clicked)
        self.dob.setEnabled(False)
        self.dob.clicked.connect(self.dobav)
        self.obnov.clicked.connect(self.Obnov)

    def dobav(self):
        self.w = Dobav(self.fname, self.data)
        self.w.show()

    def Obnov(self):
        self.data = pd.read_excel(self.fname)
        df = len(self.data["Дата"].index)
        headers = self.data.columns.values.tolist()
        self.tableWidget.setColumnCount(len(headers))
        self.tableWidget.setHorizontalHeaderLabels(headers)
        for i, row in self.data.iterrows():
            self.tableWidget.setRowCount(df)
            for j in range(self.tableWidget.columnCount()):
                self.tableWidget.setItem(i, j, QTableWidgetItem(str(row[j])))



    def run(self):
        self.fname = QFileDialog.getOpenFileName(
            self, 'Выбрать файл с данными для учета', '',
            'Файл (*.xlsx);;Файл (*.xls);;Файл (*.xlsm);;'
            'Файл (*.odf);;Файл (*.ods);;Файл (*.odt);;'
            'Файл (*.xlsb);;Все файлы (*)')[0]
        if not self.fname:
            self.statusbar.showMessage('Не выбран файл с данными')
            return
        self.statusbar.showMessage(f'Открыт файл {self.fname}')
        # добавить проверку на расширение файла
        self.dob.setEnabled(True)
        self.data = pd.read_excel(self.fname)
        headers = self.data.columns.values.tolist()
        self.tableWidget.setColumnCount(len(headers))
        self.tableWidget.setHorizontalHeaderLabels(headers)
        for i, row in self.data.iterrows():
            self.tableWidget.setRowCount(self.tableWidget.rowCount() + 1)
            for j in range(self.tableWidget.columnCount()):
                self.tableWidget.setItem(i, j, QTableWidgetItem(str(row[j])))

    def export_to_xlsx(self):
        if not self.fname:
            self.statusbar.showMessage('Не выбран файл с данными')
            return
        columnHeaders = []
        for j in range(self.tableWidget.model().columnCount()):
            columnHeaders.append(self.tableWidget.horizontalHeaderItem(j).text())
        df = pd.DataFrame(columns=columnHeaders)
        for row in range(self.tableWidget.rowCount()):
            for col in range(self.tableWidget.columnCount()):
                df.at[row, columnHeaders[col]] = self.tableWidget.item(row, col).text()
        df.to_excel(self.fname, index=False)
        self.statusbar.showMessage(f'Данные успешно сохранены в файл {self.fname}')

    def header_clicked(self, logicalIndex):
        self.logicalIndex = logicalIndex
        self.menuValues = QMenu(self)
        self.signalMapper = QSignalMapper(self)
        valuesUnique = []
        for row in range(self.tableWidget.rowCount()):
            if not self.tableWidget.isRowHidden(row):
                valuesUnique.append(self.tableWidget.item(row, self.logicalIndex).text())
        self.Menudialog = QDialog(self)
        self.Menudialog.setWindowTitle('Выберите элемент для фильтра')
        self.Menudialog.setWindowIcon(QIcon('logo.png'))
        self.Menudialog.setGeometry(self.pos().x() + 100, self.pos().y() + 100, 200, 300)

        self.list_widget = QListWidget(self)
        self.list_widget.addItems(sorted(list(set(valuesUnique))))
        scroll_area = QScrollArea(self.Menudialog)
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(self.list_widget)
        self.list_widget.currentItemChanged.connect(self.item_select)
        self.Menudialog.exec()

    def item_select(self):
        item = self.list_widget.currentItem()
        column = self.logicalIndex
        for i in range(self.tableWidget.rowCount()):
            if self.tableWidget.item(i, column).text() != item.text():
                self.tableWidget.setRowHidden(i, True)
        self.Menudialog.close()

    def on_signalMapper_mapped(self, i):
        for i in range(self.tableWidget.rowCount()):
            if self.tableWidget.item(i, self.logicalIndex).text() != self.signalMapper.mapping(i).text():
                self.tableWidget.setRowHidden(i, True)

    def filter(self):
        for i in range(self.tableWidget.rowCount()):
            if (self.tableWidget.item(i, self.filter_box.currentIndex()).text() != self.search_box.text()
                    and self.search_box.text()):
                self.tableWidget.setRowHidden(i, True)

    def search(self):
        for i in range(self.tableWidget.rowCount()):
            self.tableWidget.setRowHidden(i, False)
        for i in range(self.tableWidget.rowCount()):
            for j in range(self.tableWidget.columnCount()):
                if self.search_box.text() not in self.tableWidget.item(i, j).text():
                    self.tableWidget.setRowHidden(i, True)

    def unfilter(self):
        for i in range(self.tableWidget.rowCount()):
            self.tableWidget.showRow(i)

    def insert_data(self, df):
        self.header = df.columns.tolist()
        new_data = df.values.tolist()
        self.tableWidget.setRowCount(len(new_data))
        self.tableWidget.setColumnCount(len(new_data[0]))
        self.tableWidget.setHorizontalHeaderLabels(self.header)
        for r in range(len(new_data)):
            for c in range(len(new_data[0])):
                self.tableWidget.setItem(r, c, QTableWidgetItem(str(new_data[r][c])))
                self.tableWidget.item(r, c).setFlags(Qt.ItemFlag.ItemIsEnabled)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = AccountingSystem()
    ex.show()
    sys.exit(app.exec())
