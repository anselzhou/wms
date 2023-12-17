# coding:utf-8
import os
import threading
import xlrd
import xlwt
from playsound import playsound
from PyQt5 import QtCore
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QWidget, QFileDialog, QTableWidgetItem, QMessageBox
from Ui_PickingList import Ui_Form

class PickingList(QWidget, Ui_Form):
    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.setupUi(self)
        self.file_path2 = ''
        self.PushButton.clicked.connect(self.choice_file)
        self.PushButton_2.clicked.connect(self.open_file)
        self.PrimaryPushButton.clicked.connect(self.to_xls)
        # self.LineEdit.returnPressed.connect(self.scaner)

    def choice_file(self):
        now_path = os.getcwd()
        choice_file_path2 = QFileDialog.getOpenFileName(self, '选择文件', now_path, 'Excel files(*.xlsx , *.xls)')
        # 如果路径存在，设置文件路径和输入框内容
        if choice_file_path2[0]:
            self.file_path2 = choice_file_path2[0]
            self.LineEdit.setText(choice_file_path2[0])

    def open_file(self):
        # 判断文件路径是否有值(就是有没有选择了文件)
        if self.file_path2:
            # 获取文件后缀
            file_format = self.file_path2.split('.')[-1]
            # 判断文件格式是否是'xls'或者'xlsx'
            if file_format == 'xls' or file_format == 'xlsx' or file_format == 'XLS' or file_format == 'XLSX':
                # 使用xlrd提取excel数据
                workbook = xlrd.open_workbook(self.file_path2)
                # 根据sheet索引或者名称获取sheet内容
                # sheet2_name = workbook.sheet_names()[0]  # 获取表格的名称
                sheet1 = workbook.sheet_by_index(0)  # sheet索引从0开始
                # cols = sheet1.col_values(0)  # 获取第1列内容
                rows = sheet1.nrows
                columns = sheet1.ncols
                self.TableWidget.setRowCount(rows)
                self.TableWidget.setColumnCount(columns + 1)
                self.TableWidget.verticalHeader().setVisible(False)
                data = sheet1.row_values(0)  # 获取第1行内容
                data.append('已扫描数量')
                self.TableWidget.setHorizontalHeaderLabels(data)  # 将第1行内容设置成表头
                self.TableWidget.resizeColumnsToContents()
                for r in range(1, rows):
                    for c in range(columns):
                        # 创建QTableWidgetItem对象，并设置值
                        item = QTableWidgetItem(str(sheet1.cell_value(r, c)))
                        # 设置QTableWidgetItem对象启用、可以选中和可以编辑
                        item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
                        # 设置QTableWidgetItem对象里面的值水平，垂直居中
                        item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                        # 将新建的QTableWidgetItem添加到tableWidget中
                        self.TableWidget.setItem(r - 1, c, item)
                self.LineEdit.setText("")
                self.TableWidget.resizeColumnToContents(0)
                self.TableWidget.resizeColumnToContents(1)
                c2 = self.TableWidget.columnCount()
                c2 = c2 - 1
                for r2 in range(self.TableWidget.rowCount()):
                    item2 = QTableWidgetItem(str(0))
                    self.TableWidget.setItem(r2, c2, item2)
            else:
                QMessageBox.warning(self, '警告', '请选择xlsx或者xls文件！')
        else:
            QMessageBox.warning(self, '警告', '请选择文件！')


    def to_xls(self):
        now_path = os.getcwd()
        choice_file_path = QFileDialog.getSaveFileName(self, '选择文件', now_path, 'Excel files(*.xls)')
        if choice_file_path[0]:
            file_path = choice_file_path[0]
        # 创建新的workbook（其实就是创建新的excel）
        workbook = xlwt.Workbook(encoding='utf-8')
        rows = self.TableWidget.rowCount()
        columns = self.TableWidget.columnCount()
        # 创建新的sheet表
        worksheet = workbook.add_sheet("My new Sheet")

        for row in range(rows):
            for col in range(columns):
                item = self.TableWidget.item(row, col)
                if item is not None:
                    # 往表格写入内容,这里可以通过嵌套两层For循环历便表格内容后写入到excel文件中
                    worksheet.write(row, col, item.text())
        # 保存
        workbook.save(file_path)