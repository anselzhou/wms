# coding:utf-8
import os
import threading

from playsound import playsound
from PyQt5.QtWidgets import QWidget
import pymysql

from Ui_Putaway import Ui_putaway


class Putaway(QWidget, Ui_putaway):
    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.setupUi(self)
        self.BodyLabel.setText("请先指定库位！")
        self.PushButton.clicked.connect(self.put)
        self.LineEdit.returnPressed.connect(self.put_theard)

    def put_theard(self):
        t = threading.Thread(target=self.put)
        t.start()

    def put(self):
        ean = self.LineEdit.text()
        address = self.BodyLabel.text()
        root_path = os.getcwd()  # 获取play_sound.py的根目录
        music_sure = os.path.join(root_path, 'audio', 'sure.wav')  # 获取音频路径
        music_error = os.path.join(root_path, 'audio', 'error.wav')  # 获取音频路径
        music_success = os.path.join(root_path, 'audio', 'Weird_Coin_07_Alert_HUD_Quirky_Short_Cm.wav')  # 获取音频路径
        db = pymysql.connect(host='localhost', user='root', password='123456', charset='utf8', db='wms')
        cursor = db.cursor()
        select_material = '''
        select * from material_address where ean=%s and address='cache_1'
        '''
        select_address = '''
        select * from address_info where address=%s
        '''
        insert_sql = '''
        insert into material_address value(%s,%s,%s,%s,1) ON DUPLICATE KEY UPDATE qty = qty + 1
        '''
        insert_cache = '''
        update material_address set qty = qty -1 where address = 'cache_1' and ean = %s
        '''
        try:
            cursor.execute(select_address, ean)
            if cursor.fetchall():
                self.BodyLabel.setText(ean)
                self.LineEdit.setText("")
                playsound(music_success)
            elif address == "请先指定库位！":
                self.LineEdit.setText("")
                playsound(music_error)
            else:
                cursor.execute(select_material, ean)
                data = cursor.fetchall()[0]
                if data:
                    cursor.execute(insert_cache, ean)
                    cursor.execute(insert_sql, (address, ean, data[2], data[3]))
                    self.LineEdit.setText("")
                    playsound(music_sure)
                    self.ListWidget.insertItem(0, f'{ean}                 -------已上架至{address}')
                else:
                    playsound(music_error)
                    self.LineEdit.setText("")
        except Exception as e:
            print(e)
        finally:
            cursor.execute('delete from material_address where qty=0')
            db.commit()
            db.close()
            cursor.close()


