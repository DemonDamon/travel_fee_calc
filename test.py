import sys
from PyQt5.QtWidgets import QWidget, QApplication, QDialog, QPushButton, QHBoxLayout


class Example(QWidget):
    def __init__(self):
        super().__init__()
        self.button = QPushButton("打开对话框", self)
        self.show()

class MyDialog(QDialog):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        ok_button = QPushButton("确认", self)
        ok_button.clicked.connect(self.ok)
        # 问题一：点击“取消”按钮后，主窗口也关闭了
        # 问题二：注释掉下面第二句时，窗口只是一闪而过
        # 问题三：使用connect连接的函数中出现异常时，不显示异常信息，只是退出程序
        cancel_button = QPushButton("取消", self)
        cancel_button.clicked.connect(self.close)

        # 设置水平布局
        hbox = QHBoxLayout()
        hbox.addWidget(ok_button)
        hbox.addWidget(cancel_button)
        self.setLayout(hbox)


    def ok(self):
        print("ok")


app = QApplication(sys.argv)
ex = Example()
a=MyDialog()
ex.button.clicked.connect(a.show)
sys.exit(app.exec_())