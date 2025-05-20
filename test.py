import win32com.client
import sys, os, time, datetime
from playsound3 import playsound
import threading  # 추가
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox  # 추가

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True  # 엑셀 프로그램 보이게 하기

file_path = '1.XLS'
if not os.path.exists(file_path):
    app = QtWidgets.QApplication(sys.argv)
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Critical)
    msg.setWindowTitle("파일 오류")
    msg.setText(f"File not found: {file_path}")
    msg.setStandardButtons(QMessageBox.Ok)
    msg.exec_()
    sys.exit(0)  # 정상 종료

# 음성 파일 경로 설정
audio_file_path = 'ya.m4a'
bgm_file_path = '02 - DARK SOULS Ⅲ - Yuka Kitamura.mp3'  # 배경음 파일명(같은 폴더에 bgm.mp3 필요)


# 배경음 반복 재생 함수
def play_bgm():
    while True:
        playsound(os.path.abspath(bgm_file_path))

class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(400, 368)
        self.pushButton = QtWidgets.QPushButton(Dialog)
        self.pushButton.setGeometry(QtCore.QRect(10, 10, 381, 251))
        self.pushButton.setIconSize(QtCore.QSize(30, 50))
        self.pushButton.setCheckable(False)
        self.pushButton.setObjectName("pushButton")
        self.textBrowser = QtWidgets.QTextBrowser(Dialog)
        self.textBrowser.setGeometry(QtCore.QRect(20, 280, 251, 81))
        self.textBrowser.setObjectName("textBrowser")

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

        # 버튼 클릭 시 동작을 연결합니다.
        self.pushButton.clicked.connect(self.on_pushButton_clicked)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "업무 간단히"))
        self.pushButton.setText(_translate("Dialog", "딱!"))
        self.pushButton.setFont(QtGui.QFont("Arial", 40, QtGui.QFont.Bold))
        
        
    # 버튼 클릭 시 실행되는 함수입니다.
    # 엑셀 파일을 열고 데이터를 처리합니다.
    def on_pushButton_clicked(self):
        # 음성 파일 재생
        sound = playsound(os.path.abspath(audio_file_path), block=False)
        print(os.path.abspath(audio_file_path))
        self.textBrowser.setText("버튼이 클릭되었습니다!")
        wb = excel.Workbooks.Open(os.path.abspath(file_path))
        ws = wb.Sheets('Sheet1')  # 시트 가져오기
        ws2 = wb.Sheets('Sheet')  # 시트 가져오기

        new_file_path = '오늘의엑셀.XLS'
        new_wb = excel.Workbooks.Add()
        new_ws = new_wb.Sheets('Sheet1')  # 새로운 시트 가져오기
        new_ws.Name = "대응완료"  # 시트 이름 변경

        self.textBrowser.append("스노우볼 해킹중!")
        time_now = datetime.datetime.now()
        new_ws.Cells(1, 1).value = "상태"
        new_ws.Cells(1, 2).value = "접수일"
        new_ws.Cells(1, 3).value = "고객사"
        new_ws.Cells(1, 4).value = "접수내용"
        new_ws.Cells(1, 5).value = "처리내용"
        cell = 2
        cell2 = 2
        print(time_now.strftime("%Y-%m-%d"))
        while True:  # 대응
            value = ws.Cells(cell, 18).value  # 날짜
            if value is None:
                break
            if time_now.strftime("%Y-%m-%d") == value.strftime("%Y-%m-%d"):
                print("True")
                register = ws.Cells(cell, 3).value
                customer = ws.Cells(cell, 9).value
                info = ws.Cells(cell, 17).value
                chat = ws.Cells(cell, 19).value
                new_ws.Cells(cell2, 1).value = "대응"
                new_ws.Cells(cell2, 1).Interior.Color = 65535
                new_ws.Cells(cell2, 2).value = register
                new_ws.Cells(cell2, 3).value = customer
                new_ws.Cells(cell2, 4).value = info
                new_ws.Cells(cell2, 5).value = chat
                print(register)
                cell2 += 1
            else:
                print("False")
            cell += 1

        cell = 2
        while True:  # 완료
            value = ws2.Cells(cell, 18).value  # 날짜
            if value is None:
                break
            if time_now.strftime("%Y-%m-%d") == value.strftime("%Y-%m-%d"):
                print("True")
                register = ws2.Cells(cell, 3).value
                customer = ws2.Cells(cell, 9).value
                info = ws2.Cells(cell, 17).value
                chat = ws2.Cells(cell, 19).value
                new_ws.Cells(cell2, 1).value = "완료"
                new_ws.Cells(cell2, 1).Interior.Color = 14277081  # 연한 녹색으로 변경
                new_ws.Cells(cell2, 2).value = register
                new_ws.Cells(cell2, 3).value = customer
                new_ws.Cells(cell2, 4).value = info
                new_ws.Cells(cell2, 5).value = chat
                print(register)
                cell2 += 1
            else:
                print("False")
            cell += 1

        new_ws.Columns("A:E").AutoFit()

        # 모든 셀에 테두리 추가
        last_row = new_ws.Cells(new_ws.Rows.Count, 1).End(-4162).Row  # -4162는 xlUp 상수
        last_col = new_ws.Cells(1, new_ws.Columns.Count).End(-4159).Column  # -4159는 xlToLeft 상수
        for row in range(1, last_row + 1):
            for col in range(1, last_col + 1):
                cell = new_ws.Cells(row, col)
                borders = cell.Borders
                borders.LineStyle = 1  # xlContinuous
                borders.Weight = 2  # xlThin

        # 새로운 엑셀 파일 저장
        new_wb.SaveAs(os.path.abspath(new_file_path))
        self.textBrowser.append("구글 시트 연동중!")
        self.textBrowser.append("작업완료!")

if __name__ == "__main__":
    # 배경음 스레드 시작
    bgm_thread = threading.Thread(target=play_bgm, daemon=True)
    bgm_thread.start()

    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())
