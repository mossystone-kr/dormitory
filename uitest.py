import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5 import uic

#UI파일 연결
form_class = uic.loadUiType("untitled.ui")[0]

#화면을 띄우는데 사용되는 Class 선언
class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        self.save=[False]*7
        self.btn_2f.clicked.connect(lambda: self.roomPlan2(self.save[2]))
        self.btn_3f.clicked.connect(lambda: self.roomPlan3(self.save[3]))
        self.btn_4f.clicked.connect(lambda: self.roomPlan4(self.save[4]))
        self.btn_sm.clicked.connect(lambda: self.studyPlanM(self.save[5]))
        self.btn_sf.clicked.connect(lambda: self.studyPlanF(self.save[6]))
        self.btn_xlsx.clicked.connect(self.exportXLSX)
        self.btn_png.clicked.connect(self.exportPNG)

        self.check_2f.stateChanged.connect(lambda: self.checkFunction(2))
        self.check_3f.stateChanged.connect(lambda: self.checkFunction(3))
        self.check_4f.stateChanged.connect(lambda: self.checkFunction(4))
        self.check_sm.stateChanged.connect(lambda: self.checkFunction(5))
        self.check_sf.stateChanged.connect(lambda: self.checkFunction(6))

        self.actionExit.triggered.connect(qApp.quit)
        self.studentFileName=''
        self.actionUpload.triggered.connect(self.openFile)
        self.visualOn=False
        self.actionVisual.triggered.connect(self.changeMode)
        self.plan=[[] for i in range(5)]
        self.tmpPlan=[[] for i in range(5)]

    def roomPlan2(self, ifsave):
        if self.studentFileName=='': QMessageBox.about(self, 'Warning', 'No File Selected')
        if ifsave:
            self.plan[0]=self.tmpPlan[0]

        print("2층 배치")

    def roomPlan3(self, ifsave):
        if self.studentFileName == '': QMessageBox.about(self, 'Warning', 'No File Selected')
        if ifsave:
            self.plan[1]=self.tmpPlan[1]
        print("3층 배치")

    def roomPlan4(self, ifsave):
        if self.studentFileName == '': QMessageBox.about(self, 'Warning', 'No File Selected')
        if ifsave:
            self.plan[2]=self.tmpPlan[2]
        print("4층 배치")

    def studyPlanM(self, ifsave):
        if self.studentFileName == '': QMessageBox.about(self, 'Warning', 'No File Selected')
        if ifsave:
            self.plan[3]=self.tmpPlan[3]
        print("자습실(남) 배치")

    def studyPlanF(self, ifsave):
        if self.studentFileName == '': QMessageBox.about(self, 'Warning', 'No File Selected')
        if ifsave:
            self.plan[4]=self.tmpPlan[4]
        print("자습실(여) 배치")

    def exportXLSX(self):
        print("엑셀로 내보내기")

    def exportPNG(self):
        print("사진으로 내보내기")

    def checkFunction(self, num): # 2, 3, 4: 층 / 5: 자습실 남 / 6: 자습실 여
        if num==2:
            if self.check_2f.isChecked(): self.save[2]=True
            else: self.save[2]=False
        elif num==3:
            if self.check_3f.isChecked(): self.save[3]=True
            else: self.save[3]=False
        elif num == 4:
            if self.check_4f.isChecked(): self.save[4] = True
            else: self.save[4] = False
        elif num == 5:
            if self.check_sm.isChecked(): self.save[5] = True
            else: self.save[5] = False
        elif num == 6:
            if self.check_sf.isChecked(): self.save[6] = True
            else: self.save[6] = False

    def openFile(self):
        filename = QFileDialog.getOpenFileName(self, 'Open File')
        if filename[0] != '':
            self.studentFileName = filename[0]
        else:
            QMessageBox.about(self, 'Warning', 'No File Selected')

    def changeMode(self):
        self.visualOn=not self.visualOn

if __name__ == "__main__" :
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec_()