# -*- coding:gbk -*-
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.Qt import *
from PyQt5.QtCore import QObject, pyqtSignal
import threading
import sys


class OutputRedirector(QObject):
    # print输出重定向

    update_output = pyqtSignal(str)

    def __init__(self, text_edit):
        super().__init__()
        self.text_edit = text_edit
        self.stdout_backup = sys.stdout
        self.stderr_backup = sys.stderr
        self.lock = threading.Lock()  # 添加锁对象

    def initUI(self):
        self.update_output.connect(self.write)  # 将信号连接到槽函数

    def write(self, message):
        with self.lock:
            cursor = self.text_edit.textCursor()
            cursor.movePosition(QTextCursor.End)
            self.text_edit.setTextCursor(cursor)
            self.text_edit.insertPlainText(message)
            QApplication.processEvents()

    def flush(self):
        pass

    def start_redirect(self):
        sys.stdout = self
        sys.stderr = self

    def stop_redirect(self):
        sys.stdout = self.stdout_backup
        sys.stderr = self.stderr_backupz
