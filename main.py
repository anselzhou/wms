import threading
import Ui_mymain
import os
import pymysql
import sys
import test
import xlrd
from playsound import playsound
from PyQt5 import QtCore, QtWidgets, QtGui
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QFileDialog, QTableWidgetItem, QMessageBox
from im_sure import Ui_Form
from Select_Stock import SelectStock
from Put_away import Putaway
from All_Scaner import AllScaner
from Take_Stock import TakeStock
from Picking_List import PickingList

class LoginWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = test.Ui_login()
        self.ui.setupUi(self)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.shadow = QtWidgets.QGraphicsDropShadowEffect(self)
        self.shadow.setOffset(0, 0)
        self.shadow.setBlurRadius(18)
        self.shadow.setColor(QtCore.Qt.black)
        self.ui.frame.setGraphicsEffect(self.shadow)
        db = pymysql.connect(host='localhost', user='root', password='123456', charset='utf8', db='wms')
        cursor = db.cursor()
        # 执行查询语句
        cursor.execute('select * from user_record')
        record_f = cursor.fetchall()
        for i in record_f:
            if i[2] == 1:
                self.ui.EditableComboBox.setText(i[0])
                self.ui.PasswordLineEdit.setText(i[1])
                self.ui.CheckBox.setChecked(True)
            else:
                self.ui.CheckBox.setChecked(False)
        cursor.close()
        db.close()
        self.ui.PushButton.clicked.connect(lambda: self.ui.stackedWidget.setCurrentIndex(0))
        self.ui.PushButton_2.clicked.connect(lambda: self.ui.stackedWidget.setCurrentIndex(1))
        self.ui.PrimaryPushButton.clicked.connect(self.login)
        self.ui.EditableComboBox.returnPressed.connect(self.login)
        self.ui.PasswordLineEdit.returnPressed.connect(self.login)
        self.ui.LineEdit_3.returnPressed.connect(self.register)
        self.ui.LineEdit_2.returnPressed.connect(self.register)
        self.ui.PrimaryPushButton_2.clicked.connect(self.register)

    def login(self):
        label = self.ui.label
        db = pymysql.connect(host='localhost', user='root', password='123456', charset='utf8', db='wms')
        cursor = db.cursor()
        try:
            account = self.ui.EditableComboBox.text()
            password = self.ui.PasswordLineEdit.text()
            cursor.execute('select * from users where user_name=%s and user_password=%s', (account, password))
            if account == "" or password == "":
                label.setText("账号密码不能为空！")
            elif cursor.fetchall():
                if self.ui.CheckBox.isChecked():
                    cursor.execute(
                        'update user_record set user_name=%s, user_password=%s, remember=1 where record_id=1',
                        (account, password))
                else:
                    cursor.execute('update user_record set remember=0 where record_id=1')

                self.ui = MyMainWindow()
                self.close()
            else:
                label.setText("账号或密码错误！")
        except Exception as e:
            print(e)
        finally:
            db.commit()
            cursor.close()
            db.close()

    def register(self):
        account = self.ui.LineEdit.text()
        password = self.ui.LineEdit_2.text()
        password_sure = self.ui.LineEdit_3.text()
        if account == "" or password == "" or password_sure == "":
            self.ui.label.setText("密码账号不能为空！")
        elif password_sure != password:
            self.ui.label.setText("两次输入的密码不一致！")
        else:
            db = pymysql.connect(host='localhost', user='root', password='123456', charset='utf8', db='wms')
            cursor = db.cursor()
            try:
                cursor.execute('select * from users where user_name=%s', (account))
                if cursor.fetchall():
                    self.ui.label.setText("账户已存在！")
                else:
                    cursor.execute('insert into users values(%s,%s)', (account, password))
                    db.commit()
                    self.ui.label.setText("注册成功！")
            except Exception as e:
                self.ui.label.setText(f'{e}')
            finally:
                cursor.close()
                db.close()

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton and self.isMaximized() == False:
            self.m_flag = True
            self.m_Position = event.globalPos() - self.pos()  # 获取鼠标相对窗口的位置
            event.accept()
            self.setCursor(QtGui.QCursor(QtCore.Qt.OpenHandCursor))  # 更改鼠标图标

    def mouseMoveEvent(self, mouse_event):
        if QtCore.Qt.LeftButton and self.m_flag:
            self.move(mouse_event.globalPos() - self.m_Position)  # 更改窗口位置
            mouse_event.accept()

    def mouseReleaseEvent(self, mouse_event):
        self.m_flag = False
        self.setCursor(QtGui.QCursor(QtCore.Qt.ArrowCursor))


class MyMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_mymain.Ui_MainWindow()
        self.ui.setupUi(self)
        # self.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.setWindowFlags(Qt.FramelessWindowHint)  # 设置无边框
        self.shadow = QtWidgets.QGraphicsDropShadowEffect(self)
        self.shadow.setOffset(0, 0)
        self.shadow.setBlurRadius(18)
        self.shadow.setColor(QtCore.Qt.black)
        self.ui.frame.setGraphicsEffect(self.shadow)
        self.show()
        self.ui.ListWidget.currentRowChanged.connect(self.menu_chang)
        self.ui.lineEdit.returnPressed.connect(self.thread_add)
        self.ui.lineEdit_2.returnPressed.connect(self.thread_add2)
        self.ui.pushButton_15.clicked.connect(self.delete)
        self.ui.pushButton_19.clicked.connect(self.delete_2)
        self.ui.pushButton_16.clicked.connect(self.IM)
        self.ui.pushButton_20.clicked.connect(self.EX)
        self.ui.pushButton_13.clicked.connect(self.change_size)
        self.setAcceptDrops(True)  # 设置接受拖曳动作
        self.file_path = ''  # 定义文件路径
        self.listener()  # 调用监听函数
        self.ui.PrimaryPushButton.clicked.connect(self.records)
        self.selectstodk = SelectStock(self)
        self.ui.stackedWidget.addWidget(self.selectstodk)
        self.putaway = Putaway(self)
        self.ui.stackedWidget.addWidget(self.putaway)
        self.allscaner = AllScaner(self)
        self.ui.stackedWidget.addWidget(self.allscaner)
        self.takestock = TakeStock(self)
        self.ui.stackedWidget.addWidget(self.takestock)
        self.pickinglist = PickingList(self)
        self.ui.stackedWidget.addWidget(self.pickinglist)

    def change_size(self):
        if self.isMaximized():
            self.ui.pushButton_13.setIcon(QtGui.QIcon(u":/icons/icons/max.png"))
            self.showNormal()
        else:
            self.ui.pushButton_13.setIcon(QtGui.QIcon(u":/icons/icons/normal.png"))
            self.showMaximized()

    def listener(self):
        self.ui.PushButton.clicked.connect(self.choice_file)
        self.ui.PushButton_2.clicked.connect(self.open_file)

    def choice_file(self):
        now_path = os.getcwd()
        choice_file_path = QFileDialog.getOpenFileName(self, '选择文件', now_path, 'Excel files(*.xlsx , *.xls)')
        # 如果路径存在，设置文件路径和输入框内容
        if choice_file_path[0]:
            self.file_path = choice_file_path[0]
            self.ui.LineEdit.setText(choice_file_path[0])

    def open_file(self):
        # 判断文件路径是否有值(就是有没有选择了文件)
        if self.file_path:
            # 获取文件后缀
            file_format = self.file_path.split('.')[-1]
            # 判断文件格式是否是'xls'或者'xlsx'
            if file_format == 'xls' or file_format == 'xlsx' or file_format == 'XLS' or file_format == 'XLSX':
                # 使用xlrd提取excel数据
                workbook = xlrd.open_workbook(self.file_path)
                sheet2_name = workbook.sheet_names()[0]
                # 根据sheet索引或者名称获取sheet内容
                sheet1 = workbook.sheet_by_index(0)  # sheet索引从0开始
                # cols = sheet1.col_values(0)  # 获取第1列内容
                rows = sheet1.nrows
                columns = sheet1.ncols
                self.ui.data_table.setRowCount(rows)
                self.ui.data_table.setColumnCount(columns)
                data = sheet1.row_values(0)
                self.ui.data_table.setHorizontalHeaderLabels(data)
                self.ui.data_table.resizeColumnsToContents()
                self.ui.data_table.verticalHeader().setVisible(False)
                for r in range(1, rows):
                    for c in range(columns):
                        # 创建QTableWidgetItem对象，并设置值
                        item = QTableWidgetItem(str(sheet1.cell_value(r, c)))
                        # 设置QTableWidgetItem对象启用、可以选中和可以编辑
                        item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
                        # 设置QTableWidgetItem对象里面的值水平，垂直居中
                        item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                        # 将新建的QTableWidgetItem添加到tableWidget中
                        self.ui.data_table.setItem(r - 1, c, item)
            else:
                QMessageBox.warning(self, '警告', '请选择xlsx或者xls文件！')
        else:
            QMessageBox.warning(self, '警告', '请选择文件！')

    '''重写拖曳方法'''
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls:
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls:
            try:
                event.setDropAction(Qt.CopyAction)
            except Exception as e:
                print(e)
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls:
            event.setDropAction(Qt.CopyAction)
            event.accept()
            links = []
            for url in event.mimeData().urls():
                links.append(str(url.toLocalFile()))
            self.file_path = links[0]  # 获取文件绝对路径
            self.ui.LineEdit.setText(self.file_path)  # 设置文件绝对路径
        else:
            event.ignore()

    def EX(self):
        self.im_sure = im_sure()
        self.im_sure.setWindowTitle("确认出库")
        self.im_sure.StrongBodyLabel.setText("确定要将所扫描的商品出库吗？")
        self.im_sure.dialogSignel.connect(self.slot_emit2)
        self.im_sure.show()

    def IM(self):
        self.im_sure = im_sure()
        self.im_sure.dialogSignel.connect(self.slot_emit)
        self.im_sure.show()

    def slot_emit2(self, flag):
        if flag == 0:
            db = pymysql.connect(host='localhost',user='root',password='123456',charset='utf8',db='wms')
            cursor = db.cursor()
            select_sql = '''
            select ean, material, material_description from material_address where ean=%s
            '''
            insert_sql = '''
            insert into material_address values('cache_1',%s,%s,%s,'1') ON DUPLICATE KEY UPDATE qty=qty-1
            '''
            row = self.ui.listWidget_2.count()
            list = []
            try:
                for i in range(row):
                    ean = self.ui.listWidget_2.item(i).text()
                    cursor.execute(select_sql, ean)
                    all = cursor.fetchall()
                    for i in all:
                        list.append(i)
                cursor.executemany(insert_sql, list)
                cursor.execute('delete from material_address where qty=0')
                db.commit()
            except Exception as e:
                print(e)
            finally:
                cursor.close()
                db.close()
                self.ui.listWidget_2.clear()
                self.ui.label_2.setText("出库成功！")

        else:
            self.im_sure.close()

    def slot_emit(self, flag):
        if flag == 0:
            db = pymysql.connect(host='localhost',user='root',password='123456',charset='utf8',db='wms')
            cursor = db.cursor()
            select_sql = '''
            select ean, material, material_description from material_info where ean=%s
            '''
            insert_sql = '''
            insert into material_address values('cache_1',%s,%s,%s,'1') ON DUPLICATE KEY UPDATE qty = qty + 1
            '''
            row = self.ui.listWidget.count()
            list = []
            try:
                for i in range(row):
                    ean = self.ui.listWidget.item(i).text()
                    cursor.execute(select_sql, ean)
                    all = cursor.fetchall()
                    for i in all:
                        list.append(i)
                cursor.executemany(insert_sql, list)
                db.commit()
            except Exception as e:
                print(e)
            finally:
                cursor.close()
                db.close()
                self.ui.listWidget.clear()
                self.ui.label.setText("入库成功！")

        else:
            self.im_sure.close()

    def menu_chang(self):
        row = self.ui.ListWidget.currentRow()
        self.ui.stackedWidget.setCurrentIndex(row)

    def thread_add(self):  # 创建子线程
        t = threading.Thread(target=self.add)  # 子线程执行 add
        t.start()

    def add(self):  # 将扫描的条码添加到列表
        db = pymysql.connect(host='localhost', user='root', password='123456', charset='utf8', db='wms')
        cursor = db.cursor()
        select_sql = '''
        select * from material_info where ean=%s
        '''
        try:
            ean = self.ui.lineEdit.text()
            cursor.execute(select_sql, ean)
            root_path = os.getcwd()  # 获取play_sound.py的根目录
            music_sure = os.path.join(root_path, 'audio', 'sure.wav')  # 获取音频路径
            music_error = os.path.join(root_path, 'audio', 'error.wav')  # 获取音频路径
            if cursor.fetchall():
                list = self.ui.listWidget
                list.insertItem(0, ean)
                self.ui.lineEdit.setText("")
                playsound(music_sure)
            else:
                playsound(music_error)
                self.ui.lineEdit.setText("")
        except Exception as e:
            print(e)
        finally:
            cursor.close()
            db.close()

    def thread_add2(self):
        t = threading.Thread(target=self.add_2)  # 子线程执行 add_2
        t.start()

    def add_2(self):
        db = pymysql.connect(host='localhost', user='root', password='123456', charset='utf8', db='wms')
        cursor = db.cursor()
        select_sql = '''
                select * from material_address where ean=%s and address='cache_1'
                '''
        try:
            ean = self.ui.lineEdit_2.text()
            cursor.execute(select_sql, ean)
            root_path = os.getcwd()  # 获取play_sound.py的根目录
            music_sure = os.path.join(root_path, 'audio', 'sure.wav')  # 获取音频路径
            music_error = os.path.join(root_path, 'audio', 'error.wav')  # 获取音频路径
            if cursor.fetchall():
                # print(cursor.fetchall())
                list = self.ui.listWidget_2
                ean = self.ui.lineEdit_2.text()
                list.insertItem(0, ean)
                self.ui.lineEdit_2.setText("")
                playsound(music_sure)
            else:
                playsound(music_error)
                self.ui.lineEdit_2.setText("")
        except Exception as e:
            print(e)
        finally:
            cursor.close()
            db.close()

    def delete(self):
        # 删除扫描错误的条码
        selected_row = self.ui.listWidget.currentRow()
        item = self.ui.listWidget.takeItem(selected_row)
        del item

    def delete_2(self):
        selected_row = self.ui.listWidget_2.currentRow()
        item = self.ui.listWidget_2.takeItem(selected_row)
        del item

    def records(self):
        database = pymysql.connect(host='localhost', user='root', password='123456', charset='utf8', db='wms')
        cursor = database.cursor()
        data = []
        rowCount = self.ui.data_table.rowCount()
        rowCount = rowCount - 1
        columnCount = self.ui.data_table.columnCount()
        for row in range(rowCount):
            row_item = []
            for column in range(columnCount):
                row_item.append(str(self.ui.data_table.item(row, column).text()))
            data.append(row_item)
        insert_sql = '''
        insert into material_info values(%s,%s,%s,%s)
        '''
        cursor.executemany(insert_sql, data)
        database.commit()
        cursor.close()
        database.close()

    def mousePressEvent(self, event):  # 窗口拖拽
        if event.button() == QtCore.Qt.LeftButton and self.isMaximized() == False:
            self.m_flag = True
            self.m_Position = event.globalPos() - self.pos()  # 获取鼠标相对窗口的位置
            event.accept()
            self.setCursor(QtGui.QCursor(QtCore.Qt.OpenHandCursor))  # 更改鼠标图标

    def mouseMoveEvent(self, mouse_event):
        if QtCore.Qt.LeftButton and self.m_flag:
            self.move(mouse_event.globalPos() - self.m_Position)  # 更改窗口位置
            mouse_event.accept()

    def mouseReleaseEvent(self, mouse_event):
        self.m_flag = False
        self.setCursor(QtGui.QCursor(QtCore.Qt.ArrowCursor))


class im_sure(QWidget, Ui_Form):
    dialogSignel = pyqtSignal(int,)

    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.PushButton_sure.clicked.connect(self.accept)
        self.PrimaryPushButton_2.clicked.connect(self.reject)

    def accept(self):
        self.dialogSignel.emit(0,)
        self.destroy()

    def reject(self):
        self.dialogSignel.emit(1,)
        self.destroy()


if __name__ == '__main__':
    # enable dpi scale
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps)
    app = QApplication(sys.argv)
    ui = LoginWindow()
    ui.show()
    sys.exit(app.exec_())
