# coding:utf-8
import sys
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QWidget, QTableWidgetItem
import pymysql

from Ui_SelectStock import Ui_SelectStock


class SelectStock(QWidget, Ui_SelectStock):
    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.setupUi(self)
        self.BodyLabel.setText("")
        self.LineEdit.returnPressed.connect(self.select)
        self.PushButton.clicked.connect(self.select)


    def select(self):
        db = pymysql.connect(host='localhost', user='root', password='123456', charset='utf8', db='wms')
        cursor = db.cursor()
        ean = self.LineEdit.text()
        select_sql = '''
        select * from material_address where ean like %s
        '''
        try:
            cursor.execute(select_sql, ean)
            data = cursor.fetchall()
            herder = cursor.description
            self.TableWidget.verticalHeader().setVisible(False)
            self.TableWidget.setRowCount(len(data))
            self.TableWidget.setColumnCount(5)
            for i in range(5):
                self.TableWidget.setHorizontalHeaderItem(i, QTableWidgetItem(str(herder[i][0])))
            for a in range(len(data)):
                for b in range(5):
                    self.TableWidget.setItem(a, b, QTableWidgetItem(str(data[a][b])))
            self.TableWidget.resizeColumnsToContents()
        except Exception as e:
            print(e)
        finally:
            cursor.close()
            db.close()
