# coding:utf-8
import os
import threading

import pymysql
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
        db = pymysql.connect(host='localhost', user='root', password='123456', charset='utf8', db='wms')
        cursor = db.cursor()
        sql_select = '''
        select * from material_address where ean=%s
        '''
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
                rows = sheet1.nrows
                table_data = []
                for i in range(1, rows):
                    data = sheet1.row_values(i)
                    pick = int(data[1])
                    cursor.execute(sql_select, data[0])
                    s_data = cursor.fetchall()
                    for j in s_data:
                        j = list(j)
                        if pick < j[4]:
                            j[4] = pick
                        elif pick > j[4] or pick == j[4]:
                            pick = pick - j[4]
                        if j[4] != 0:
                            table_data.append(j)
                row = len(table_data)
                column = 5
                self.TableWidget.setRowCount(row)
                self.TableWidget.setColumnCount(column)
                for r in range(row):
                    for c in range(column):
                        # 创建QTableWidgetItem对象，并设置值
                        item = QTableWidgetItem(str(table_data[r][c]))
                        # 设置QTableWidgetItem对象启用、可以选中和可以编辑
                        item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
                        # 设置QTableWidgetItem对象里面的值水平，垂直居中
                        item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                        # 将新建的QTableWidgetItem添加到tableWidget中
                        self.TableWidget.setItem(r, c, item)
                self.TableWidget.resizeColumnsToContents()  # 设置所有列自适应宽度
                self.TableWidget.verticalHeader().setVisible(False)  # 设置列头不可见
            else:
                QMessageBox.warning(self, '警告', '请选择xlsx或者xls文件！')
        else:
            QMessageBox.warning(self, '警告', '请选择文件！')

    def to_xls(self):
        # 保存到Excel文件
        try:
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
        except Exception as e:
            print(f"未知异常发生了：{e}")

        # 将库位中的货品移到cache_1库位
        try:
            # 链接数据库
            db = pymysql.connect(host='localhost', user='root', password='123456', charset='utf8', db='wms')
            cursor = db.cursor()
            change_address = '''
                    update material_address set qty=qty-%s where address=%s and ean=%s
                    '''
            change_cache = '''
                    insert into material_address values('cache_1',%s,%s,%s,%s) ON DUPLICATE KEY UPDATE qty = qty+%s
                    '''
            rows = self.TableWidget.rowCount()
            all_address_data = []
            all_cache_data = []
            address_for = [0, 1]
            cache_for = [1, 2, 3]
            for i in range(rows):
                address_data = []
                cache_data = []
                for j in address_for:
                    address_info = self.TableWidget.item(i, j).text()
                    address_data.append(address_info)
                address_qty = self.TableWidget.item(i, 4).text()
                address_qty = int(address_qty)
                address_data.insert(0, address_qty)
                # address_data = tuple(address_data)
                all_address_data.append(address_data)
                for j in cache_for:
                    cache_info = self.TableWidget.item(i, j).text()
                    cache_data.append(cache_info)
                cache_qty = self.TableWidget.item(i, 4).text()
                cache_qty = int(cache_qty)
                cache_data.append(cache_qty)
                cache_data.append(cache_qty)
                # cache_data = tuple(cache_data)
                all_cache_data.append(cache_data)
            print(all_address_data)
            print(all_cache_data)
            cursor.executemany(change_address, all_address_data)
            cursor.executemany(change_cache, all_cache_data)
        except pymysql.Error as e:
            print(f"Error: {e}")
        finally:
            cursor.execute('delete from material_address where qty=0')
            db.commit()
            cursor.close()
            db.close()
