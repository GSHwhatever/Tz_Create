# -*- coding:gbk -*-
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.Qt import *
from PyQt5.QtCore import Qt
from SQ_TZ import JCTZ
from Write_BB import Write
import sys, os, traceback


class Query_Window(QMainWindow):

    def __init__(self, parent=None):
        # super(Query_Window, self).__init__(parent)
        super().__init__()
        self.label = None
        self.line_edit = None
        self.btn_push = None
        self.text_edit = QTextEdit()
        self.radiobutton1 = None
        self.radiobutton2 = None
        self.radiobutton3 = None
        self.line_edit5 = None
        self.btn_push5 = None
        self.btn_push6 = None
        self.btn_open_dir = None
        self.btn_open_file = None
        self.is_filter = None
        self.info_edit = None
        self.tab_widget = None
        self.t_layout = None
        self.search_info = None
        self.vertical_layout = None
        self.download_path = os.path.join(os.environ['USERPROFILE'], 'Desktop', '̨�˽��')
        if not os.path.exists(self.download_path):
            os.mkdir(self.download_path)
        self.jctz = JCTZ()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('̨���ۺ�������')
        # ������ QWidget
        main_widget = QWidget()

        self.tab_widget = QTabWidget()
        layout = QVBoxLayout(main_widget)  # ������ֱ���ֹ�����

        tab1 = QWidget()

        tab1.layout = QGridLayout()  # ����դ�񲼾ֹ�����

        self.label = QLabel('������')
        tab1.layout.addWidget(self.label, 0, 0, 1, 1)
        self.line_edit = QLineEdit()
        # self.line_edit.editingFinished.connect(self.input_check)
        self.line_edit.setPlaceholderText('ѡ��ʵ�����ļ���̨��4��̨��7')
        # print(self.line_edit.width())
        # print(self.line_edit.height())
        tab1.layout.addWidget(self.line_edit, 0, 1, 1, 1)   

        self.btn_push2 = QPushButton('ѡ���ļ�')
        self.btn_push2.clicked.connect(self.search_file)
        tab1.layout.addWidget(self.btn_push2, 0, 2, 1, 1)

        self.btn_login = QPushButton('����·��')
        self.btn_login.setEnabled(False)
        self.btn_login.clicked.connect(self.modify_path)
        tab1.layout.addWidget(self.btn_login, 0, 3, 1, 1)

        # ������
        self.text_edit.setReadOnly(True)
        self.text_edit.setStyleSheet("background-image: url(./background.png); background-attachment: fixed; background-repeat: no-repeat; background-position: center;")
        tab1.layout.addWidget(self.text_edit, 1, 0, 8, 7)

        self.btn_open_dir = QPushButton('һ������\n345612')
        self.btn_open_dir.clicked.connect(self.work_all)
        tab1.layout.addWidget(self.btn_open_dir, 1, 8, 1, 1)

        self.btn_open_dir = QPushButton('����3456')
        self.btn_open_dir.clicked.connect(self.work_to_3456)
        tab1.layout.addWidget(self.btn_open_dir, 2, 8, 1, 1)

        self.btn_open_dir = QPushButton('4����12')
        self.btn_open_dir.clicked.connect(self.work_to_12)
        tab1.layout.addWidget(self.btn_open_dir, 3, 8, 1, 1)

        self.btn_open_dir = QPushButton('7��ӵ�15')
        self.btn_open_dir.clicked.connect(self.work_to_15)
        tab1.layout.addWidget(self.btn_open_dir, 4, 8, 1, 1)

        self.btn_open_dir = QPushButton('ͳ�Ʊ���')
        # self.btn_open_dir.setEnabled(False)
        self.btn_open_dir.clicked.connect(self.work_to_gb)
        tab1.layout.addWidget(self.btn_open_dir, 5, 8, 1, 1)

        self.btn_open_dir = QPushButton('������')
        self.btn_open_dir.clicked.connect(self.clear_edit)
        tab1.layout.addWidget(self.btn_open_dir, 6, 8, 1, 1)

        tab1.setLayout(tab1.layout)

        self.tab_widget.addTab(tab1, '������')

        layout.addWidget(self.tab_widget)

        main_widget.setLayout(layout)
        self.setCentralWidget(main_widget)  # ���� QWidget ����Ϊ QMainWindow �����봰�ڲ���
        # self.show()

        self.resize(800, 600)

    def search_file(self):
        # ѡ�������ļ�
        file_dialog = QFileDialog()
        file_dialog.setDirectory(os.path.join(os.environ['USERPROFILE'], 'Desktop'))
        file_path = file_dialog.getOpenFileName(self, 'ѡ���ļ�')
        if file_path[0]:
            self.line_edit.setText(file_path[0])
            self.do_tag = True
    
    def modify_path(self):
        file_dialog = QFileDialog()
        file_dialog.setDirectory(self.download_path)
        directory_path = file_dialog.getExistingDirectory(self, 'ѡ��Ŀ¼')
        print(f'��������·��Ϊ:{directory_path}')
        if directory_path:
            self.download_path = directory_path

    def clear_edit(self):
        if self.text_edit.toPlainText():
            self.text_edit.blockSignals(True)
            self.text_edit.clear()
            self.text_edit.blockSignals(False)

    def work_all(self):
        try:
            text = self.line_edit.text()
            if text == '':
                print('�ڴ�֮ǰ������ѡ���ļ�')
            else:
                self.jctz.run_smz(text, self.download_path)
                if self.jctz.run_smz_status:
                    self.jctz.run_4to12(self.download_path, self.download_path)
        except Exception as e:
            traceback.print_exc()
            print(e)

    def work_to_3456(self):
        try:
            text = self.line_edit.text()
            if text == '':
                print('�ڴ�֮ǰ������ѡ���ļ�')
            else:
                # print(f'text:{text},{len(text)}')
                self.jctz.run_smz(text, self.download_path)
        except Exception as e:
            traceback.print_exc()
            print(e)
    
    def work_to_12(self):
        try:
            text = self.line_edit.text()
            if text == '':
                print('�ڴ�֮ǰ������ѡ���ļ�')
            else:
                # print(f'text:{text},{len(text)}')
                self.jctz.run_4to12(text, self.download_path)
        except Exception as e:
            traceback.print_exc()
            print(e)

    def work_to_15(self):
        try:
            text = self.line_edit.text()
            if text == '':
                print('�ڴ�֮ǰ������ѡ���ļ�')
            else:
                # print(f'text:{text},{len(text)}')
                self.jctz.run_7to15(text, self.download_path)
        except Exception as e:
            traceback.print_exc()
            print(e)

    def work_to_gb(self):
        try:
            text = self.line_edit.text()
            if text == '':
                print('�ڴ�֮ǰ������ѡ���ļ�')
            else:
                # print(f'text:{text},{len(text)}')
                self.write = Write()
                self.write.run(text, self.download_path)
        except Exception as e:
            traceback.print_exc()
            print(e)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Query_Window()
    window.show()
    sys.exit(app.exec_())
