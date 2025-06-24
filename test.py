import win32com.client
import sys, os, time, datetime
from playsound3 import playsound
import threading  # 추가
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox  # 추가
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

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
        self.emailLabel = QtWidgets.QLabel(Dialog)
        self.emailLabel.setGeometry(QtCore.QRect(280, 280, 100, 20))
        self.emailLabel.setText("수신자 이메일(,로 구분)")
        self.emailEdit = QtWidgets.QLineEdit(Dialog)
        self.emailEdit.setGeometry(QtCore.QRect(280, 300, 110, 20))
        self.emailEdit.setPlaceholderText("aaa@bbb.com,ccc@ddd.com")
        self.sendButton = QtWidgets.QPushButton(Dialog)
        self.sendButton.setGeometry(QtCore.QRect(280, 330, 110, 30))
        self.sendButton.setText("이메일 전송")
        self.sendButton.clicked.connect(self.send_email_with_table)

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
                ws.Cells(cell,17).value
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
                ws.Cells(cell,17).value
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
        
         # 4열(접수내용) 기준 중복값이 있으면, 5열(처리내용)에 '트래킹'이 포함된 행만 남기고 나머지 삭제
        last_row = new_ws.Cells(new_ws.Rows.Count, 1).End(-4162).Row  # -4162는 xlUp
        seen = dict()  # {접수내용: (row, content)}
        for row in range(last_row, 1, -1):  # 2행부터 시작, 역순
            key = new_ws.Cells(row, 4).value  # 4번째 컬럼(접수내용)
            content = new_ws.Cells(row, 5).value  # 5번째 컬럼(처리내용)
            if key in seen:
                # 이미 같은 접수내용이 있으면, 둘 중 '트래킹'이 포함된 것만 남김
                prev_row, prev_content = seen[key]
                # 현재 행에 '트래킹'이 있으면 이전 행 삭제, 아니면 현재 행 삭제
                if content and "트래킹" in str(content):
                    new_ws.Rows(prev_row).Delete()
                    seen[key] = (row, content)
                else:
                    new_ws.Rows(row).Delete()
            else:
                seen[key] = (row, content)

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
        
    def send_email_with_table(self):
        # 이메일 입력값 가져오기
        to_emails = self.emailEdit.text().strip()
        if not to_emails:
            self.textBrowser.append("수신자 이메일을 입력하세요.")
            return
        to_emails = [e.strip() for e in to_emails.split(",") if e.strip()]

        # new_ws의 셀 내용 HTML 테이블로 변환
        new_file_path = '오늘의엑셀.XLS'
        if not os.path.exists(new_file_path):
            self.textBrowser.append("오늘의엑셀.XLS 파일이 존재하지 않습니다.")
            return
        new_wb = excel.Workbooks.Open(os.path.abspath(new_file_path))
        new_ws = new_wb.Sheets("대응완료")
        last_row = new_ws.Cells(new_ws.Rows.Count, 1).End(-4162).Row
        last_col = new_ws.Cells(1, new_ws.Columns.Count).End(-4159).Column
        
        def excel_color_to_hex(color):
            # 엑셀 색상값(정수)을 #RRGGBB로 변환
            if color is None or color == 16777215:  # 기본 흰색
                return None
            try:
                color = int(color)
                r = color & 0xFF
                g = (color >> 8) & 0xFF
                b = (color >> 16) & 0xFF
                return f"#{r:02X}{g:02X}{b:02X}"
            except Exception:
                return None

        html = "<table border='1' cellpadding='4' cellspacing='0' style='border-collapse:collapse;'>"
        for row in range(1, last_row + 1):
            html += "<tr>"
            for col in range(1, last_col + 1):
                cell = new_ws.Cells(row, col)
                value = cell.Value if cell.Value is not None else ""
                color = cell.Interior.Color
                bgcolor = excel_color_to_hex(color)
                if bgcolor:
                    html += f"<td style='background-color:{bgcolor}'>{value}</td>"
                else:
                    html += f"<td>{value}</td>"
            html += "</tr>"
        html += "</table>"

        # 이메일 전송
        sender_email = "hsyoon@sinsungcns.com"  # 본인 이메일로 수정
        sender_pw = ""        # 앱 비밀번호 등

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = ", ".join(to_emails)
        msg['Subject'] = "엑셀 처리 결과 자동 발송"
        msg.attach(MIMEText(html, 'html'))

        try:
            # Daou Office SMTP 서버 정보
            smtp_server = "outbound.daouoffice.com"
            smtp_port = 25  # 또는 587
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.ehlo()
            server.starttls()
            server.login(sender_email, sender_pw)
            server.sendmail(sender_email, to_emails, msg.as_string())
            server.quit()
            self.textBrowser.append("이메일 전송 완료!")
        except Exception as e:
            self.textBrowser.append(f"이메일 전송 실패: {e}")

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
