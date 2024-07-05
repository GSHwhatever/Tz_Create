# -*- coding:gbk -*-
from PyQt5.QtWidgets import *
from Window_main import Query_Window
from Out_put import OutputRedirector
from PyQt5.QtCore import QTranslator
import threading
import logging
import sys, hashlib, os


class Activation(QDialog):

    def __init__(self):
        super().__init__()
        self.setWindowTitle('金保综合下载器激活')

        # 创建输入框和按钮
        self.label = QLabel('请联系管理员激活')
        self.line_edit = QLineEdit()
        self.line_edit.setPlaceholderText('输入激活码')
        self.line_edit.setEchoMode(QLineEdit.Password)
        self.button = QPushButton('激活')

        # 占位组件，隐藏
        self.lineedit2 = QTextEdit()
        self.lineedit2.setMaximumWidth(150)
        # self.lineedit2.setMaximumHeight(400)
        self.lineedit2.setReadOnly(True)
        self.lineedit2.setVisible(False)

        # 设置布局
        layout = QGridLayout()
        layout.addWidget(self.label, 0, 3, 1, 3)
        layout.addWidget(self.line_edit, 1, 1, 1, 5)
        layout.addWidget(self.button, 2, 5, 1, 1)
        layout.addWidget(self.lineedit2, 3, 0, 2, 7)

        # 将布局设置为窗口的主布局
        self.setLayout(layout)

        # 连接按钮的点击信号和处理函数
        self.button.clicked.connect(self.submit_text)

        self.resize(400, 300)

    def submit_text(self):
        text = self.line_edit.text()

        text_hash = hashlib.md5(text.encode()).hexdigest()
        with open('Activation.txt', 'r') as f:
            pass_hash = f.read()
        
        if text_hash == pass_hash:
            QMessageBox.information(self, '激活成功', '欢迎进入系统！')
            path = os.path.join(os.environ['LOCALAPPDATA'], 'Glife_TZ.txt')
            with open(path, 'w') as f:
                f.write(text_hash)
            self.close()
            main_window.show()
        else:
            QMessageBox.warning(self, '激活失败', '激活码错误，请重试。')

    def open_second_window(self):
        self.close()
        main_window.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    translator = QTranslator()
    translator.load('./config/qt_zh_CN.qm')
    app.installTranslator(translator)
    logging.basicConfig(filename='./app.log', level=logging.ERROR)
    try:
        main_window = Query_Window()
        redirector = OutputRedirector(main_window.text_edit)
        redirector_thread = threading.Thread(target=redirector.initUI)
        redirector_thread.start()
        sys.stdout = redirector
        sys.stderr = redirector
        Glife = os.path.join(os.environ['LOCALAPPDATA'], 'Glife_TZ.txt')
        if os.path.exists(Glife):
            main_window.show()
        else:
            window = Activation()
            window.show()
        sys.exit(app.exec_())
    except Exception as e:
        print(e)
        logging.exception("An exception occurred")
