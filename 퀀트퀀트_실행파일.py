import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QProgressBar
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QCoreApplication


class MyApp(QWidget):

  def __init__(self):
        super().__init__()
        self.initUI()

  def initUI(self):
        self.pbar = QProgressBar(self)
        self.btn = QPushButton('모든 종목 분석', self)
        self.btn.clicked.connect(execfile('퀀트퀀트1.py'))
        self.btn.clicked.connect(self.doAction)

        self.vbox = QVBoxLayout()
        self.vbox.addWidget(btn)
        self.vbox.addWidget(pbar)

        self.setLayout(vbox)  
        self.setWindowTitle('혁주이 투자비법')
        # self.setWindowIcon(QIcon('web.png'))
        self.setGeometry(300, 300, 300, 200)
        self.show()


if __name__ == '__main__':
  app = QApplication(sys.argv)
  ex = MyApp()
  sys.exit(app.exec_())