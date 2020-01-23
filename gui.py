import sys
from PyQt5.QtWidgets import QApplication, QWidget, QTableWidget, QTableWidgetItem, QVBoxLayout,\
    QPushButton, QLineEdit, QFileDialog
from PyQt5.QtCore import pyqtSlot
import xlwt


class App(QWidget):

    def __init__(self):
        super().__init__()
        self.title = 'PyQt5 table'
        self.left = 0
        self.top = 0
        self.width = 500
        self.height = 500
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.text_row = QLineEdit(self)
        self.text_row.setReadOnly(True)
        self.text_row.setPlaceholderText('Insert number of Rows: ')
        self.textbox1 = QLineEdit(self)
        self.textbox2 = QLineEdit(self)
        self.text_col = QLineEdit(self)
        self.text_col.setPlaceholderText('Insert number of Columns: ')
        self.text_col.setReadOnly(True)
        self.button = QPushButton('Create', self)
        self.button.clicked.connect(self.on_click)
        self.saveFile =  QPushButton('Save', self)
        self.saveFile.clicked.connect(self.on_save)
        self.tableWidget = QTableWidget()

        self.layout = QVBoxLayout()
        self.layout.addWidget(self.tableWidget)
        self.layout.addWidget(self.text_row)
        self.layout.addWidget(self.textbox1)
        self.layout.addWidget(self.text_col)
        self.layout.addWidget(self.textbox2)
        self.layout.addWidget(self.button)
        self.layout.addWidget(self.saveFile)
        self.tableWidget.setRowCount(1)
        self.tableWidget.setColumnCount(1)
        self.setLayout(self.layout)

        self.show()

    def file_save(self):
        options = QFileDialog.Options()
        filename, _ = QFileDialog.getSaveFileName(self, '&Save File', '', ".odt(*.odt)", options=options)
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet(self.title, cell_overwrite_ok=True)
        self.s_f(sheet)
        workbook.save(filename)

    def s_f(self, sheet):
        for currentColumn in range(self.tableWidget.columnCount()):
            for currentRow in range(self.tableWidget.rowCount()):
                cell = self.tableWidget.item(currentRow, currentColumn)
                plain_text = cell.text()
                sheet.write(currentRow, currentColumn, plain_text)

    def createtable(self):
        n = int(self.textbox1.text())
        m = int(self.textbox2.text())
        self.tableWidget.setRowCount(n)
        self.tableWidget.setColumnCount(m)
        for i in range(0, n):
            for j in range(0, m):
                self.tableWidget.setItem(i, j, QTableWidgetItem('Cell {},{}'.format(i, j)))
        self.tableWidget.move(0, 0)

    @pyqtSlot()
    def on_click(self):
        self.createtable()

    @pyqtSlot()
    def on_save(self):
        self.file_save()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
