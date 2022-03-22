import sys
import os
import openpyxl
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5 import uic, QtWidgets
from PyQt5.QtGui import *

form_class_MainWindow = uic.loadUiType("UI/naejeon.ui")[0]
form_class_Loading = uic.loadUiType("UI/loading.ui")[0]

class MainWindow(QMainWindow, form_class_MainWindow) :
    filename = ""
    excel = ""
    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        self.FindExcelButton.clicked.connect(self.OnClickFindExcelButton)
        self.FindNicknameButton.clicked.connect(self.OnClickFindNicknameButton)

    def OnClickFindExcelButton(self):
        filename = QtWidgets.QFileDialog.getOpenFileName(self, "Select Excel", filter='Excel (*.xlsx *.xls);;모든 파일 (*.*)')
        print(filename)
        extension = os.path.splitext(filename[0])[1]
        print(extension)
        if extension == '.xlsx' or extension == 'xls':
            print("Success File Search")
            self.FilePath.setText(filename[0])
            self.StartLoading()
            QApplication.processEvents()
            excel = openpyxl.load_workbook(filename=filename[0])
            self.StopLoading()
            print("Success Load Excel")
        else:
            print("Fail File Search")

    def OnClickFindNicknameButton(self):
        print("FindNickname")

    @pyqtSlot()
    def StartLoading(self):
        self.Loading = Loading(self)

    @pyqtSlot()
    def StopLoading(self):
        self.Loading.deleteLater()




class Loading(QWidget, form_class_Loading):

    def __init__(self, parent):
        self.running = False
        super(Loading, self).__init__(parent)
        self.setupUi(self)
        self.center()
        self.show()

        # 동적 이미지 추가
        self.movie = QMovie('Resources/Spinner2.gif', QByteArray(), self)
        self.movie.setCacheMode(QMovie.CacheAll)
        # QLabel에 동적 이미지 삽입
        self.label.setMovie(self.movie)
        # 윈도우 해더 숨기기
        self.movie.start()
        self.setWindowFlags(Qt.FramelessWindowHint)

    # 위젯 정중앙 위치
    def center(self):
        size = self.size()
        ph = self.parent().geometry().height()
        pw = self.parent().geometry().width()
        self.geometry().x = 0
        self.geometry().y = 0
        self.move(int(pw / 2 - size.width() / 2), int(ph / 2 - size.height() / 2))

if __name__ == "__main__":
    # QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv)

    # WindowClass의 인스턴스 생성
    myWindow = MainWindow()

    # 프로그램 화면을 보여주는 코드
    myWindow.show()
    # 프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()

