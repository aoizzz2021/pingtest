import subprocess
import sqlite3
import pandas as pd
from PyQt5.QtWidgets import QApplication,QMainWindow,QFileDialog,QMessageBox
from PyQt5.QtCore import QThread, pyqtSignal
import sys,os
import MainWindow
import qt_material
import time

class IPThread(QThread):
    trigger = pyqtSignal(str)
    s_trigger = pyqtSignal(int)
    f_trigger = pyqtSignal(int)
    def __init__(self):
        super().__init__()

    def __del__(self):
        self.wait()

    def run(self):
        with open('log.txt','a') as f:
            f.write(str(time.time())+'\n')
            f.write('-----------------------------------------------------------------------\n')
        # self.trigger.emit('测试开始')
        conn = sqlite3.connect('ip.db')
        cur = conn.cursor()
        sql = 'select * from ipdata'
        cur.execute(sql, )
        res = cur.fetchall()
        record = []
        temp_lst = []
        s_count = 0
        f_count = 0
        for item in res:
            record.append(item[3])
            temp_lst.append([item[1],item[2]])
        for i in range(0, len(record)):
            ip_str = 'ping -n 1 ' + str(record[i])
            p = subprocess.Popen(ip_str, stdout=subprocess.PIPE,shell=True)
            result = p.stdout.read()
            lst = str(result).split(r'\r\n')
            if 'Request timed out.' in lst:
                f_count+=1
                self.f_trigger.emit(f_count)
                with open('log.txt', 'a') as f:
                    f.write(record[i] + '\n')
                self.trigger.emit(record[i]+','+temp_lst[i][0]+','+temp_lst[i][1])
            else:
                s_count+=1
                self.s_trigger.emit(s_count)
        # self.trigger.emit('测试结束')


def text_callback(msg):
    lst = msg.split(',')
    ui.textEdit.append(lst[0])
    ui.textEdit_3.append(lst[1])
    ui.textEdit_2.append(lst[2])
    # ui.textEdit.append(msg)

def s_callback(msg):
    ui.lineEdit_4.setText(str(msg))

def f_callback(msg):
    ui.lineEdit_2.setText(str(msg))
    # ui.textEdit.append(msg)

def get_data():
    path = QFileDialog.getOpenFileName()
    conn = sqlite3.connect('ip.db')
    cur = conn.cursor()
    sql = 'delete from ipdata'
    cur.execute(sql)
    conn.commit()
    df = pd.read_excel(path[0])
    for i in range(df.shape[0]):
        sql = 'insert into ipdata values ('+str(i)+',"'+df.iloc[i][0]+'","'+df.iloc[i][1]+'","'+df.iloc[i][2]+'")'
        cur.execute(sql)
        conn.commit()
    QMessageBox.warning(win,'提示','IP数据录入成功，请开始测试！')

def check_ip_Ping():
    pvthread.start()


def get_style():
    style = ui.comboBox.currentText()
    with open('curstyle.txt','w') as f:
        f.write(style)
    qt_material.apply_stylesheet(app, theme=style + '.xml')


def init_ui(ui):
    win.setWindowTitle('IP批量测试工具')

    comb = ui.comboBox
    comb_lst = []
    with open('style.txt','r') as f:
        for item in f.readlines():
            comb_lst.append(item[:-1])
    comb.addItems(comb_lst)

    comb.activated.connect(lambda :get_style())

    ui.pushButton.clicked.connect(lambda :get_data())
    ui.pushButton_2.clicked.connect(lambda :check_ip_Ping())

    conn = sqlite3.connect('ip.db')
    cur = conn.cursor()
    sql = 'select * from ipdata'
    cur.execute(sql, )
    res = cur.fetchall()
    ui.lineEdit.setText(str(len(res)))
    ui.lineEdit_2.setText('0')
    ui.lineEdit_4.setText('0')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = QMainWindow()
    ui = MainWindow.Ui_MainWindow()
    ui.setupUi(win)
    pvthread = IPThread()
    pvthread.trigger.connect(text_callback)
    pvthread.s_trigger.connect(s_callback)
    pvthread.f_trigger.connect(f_callback)
    init_ui(ui)
    with open('curstyle.txt','r') as f:
        style = f.read()
    qt_material.apply_stylesheet(app, theme=style+'.xml')
    win.show()

    sys.exit(app.exec_())
