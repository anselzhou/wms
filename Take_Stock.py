# coding:utf-8
import os
import threading

import pymysql
import xlwt
from PyQt5.QtWidgets import QWidget, QTableWidgetItem, QFileDialog
from playsound import playsound

from Ui_TakeStock import Ui_Form


class TakeStock(QWidget, Ui_Form):
    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.setupUi(self)
        self.LineEdit_2.returnPressed.connect(self.thread)
        self.PushButton.clicked.connect(self.thread)
        self.PrimaryPushButton.clicked.connect(self.excel)

    def thread(self):
        t = threading.Thread(target=self.add)
        t.start()

    def add(self):
        root_path = os.getcwd()  # 获取play_sound.py的根目录
        music_sure = os.path.join(root_path, 'audio', 'sure.wav')  # 获取音频路径
        music_error = os.path.join(root_path, 'audio', 'error.wav')  # 获取音频路径
        if self.LineEdit.text() == "":
            self.BodyLabel.setText("请输入库位！")
        else:
            address = self.LineEdit.text()
            ean = self.LineEdit_2.text()
            db = pymysql.connect(host='localhost', user='root', password='123456', charset='utf8', db='wms')
            cursor = db.cursor()
            select_sql = '''
            select * from material_info where ean=%s
            '''
            try:
                cursor.execute(select_sql, ean)
                data = cursor.fetchall()
                data = list(data[0])
                if data:
                    playsound(music_sure)
                    data.insert(0, address)
                    col = len(data)
                    row = self.TableWidget.rowCount()
                    self.TableWidget.setRowCount(row+1)
                    self.TableWidget.setColumnCount(col)
                    for a in range(col):
                        item = QTableWidgetItem(str(data[a]))
                        self.TableWidget.setItem(row, a, item)
            except Exception as e:
                print(e)
            finally:
                cursor.close()
                db.close()

    def excel(self):
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