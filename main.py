import sys
import os
import openpyxl
from PyQt5.QtWidgets import *
from PyQt5 import uic, QtWidgets

form_class = uic.loadUiType("UI/naejeon.ui")[0]

class WindowClass(QMainWindow, form_class) :
    filename = ""
    excel = ""
    def __init__(self) :
        super().__init__()
        self.setupUi(self)
        self.FindExcelButton.clicked.connect(self.OnClickFindExcelButton)

    def OnClickFindExcelButton(self):
        filename = QtWidgets.QFileDialog.getOpenFileName(self, "Select Excel", filter='Excel (*.xlsx *.xls);;모든 파일 (*.*)')
        print(filename)
        extension = os.path.splitext(filename[0])[1]
        print(extension)
        if extension == '.xlsx' or extension == 'xls':
            print("Success File Search")
            self.FilePath.setText(filename[0])
            excel = openpyxl.load_workbook(filename=filename[0])
            print("Success Load Excel")
        else:
            print("Fail File Search")

if __name__ == "__main__":
    # QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv)

    # WindowClass의 인스턴스 생성
    myWindow = WindowClass()

    # 프로그램 화면을 보여주는 코드
    myWindow.show()
    # 프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()

