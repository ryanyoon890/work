# -*- coding: utf-8 -*-

from PyQt5 import QtCore, QtGui, QtWidgets
import winsound
import time
import ctypes
from ctypes import wintypes
import subprocess

# Windows API 설정을 위한 구조체 정의
class WAVEOUTCAPS(ctypes.Structure):
    _fields_ = [
        ("wMid", wintypes.WORD),
        ("wPid", wintypes.WORD),
        ("vDriverVersion", wintypes.UINT),
        ("szPname", wintypes.WCHAR * 32), # 장치 이름이 저장되는 곳
        ("dwFormats", wintypes.DWORD),
        ("wChannels", wintypes.WORD),
        ("wReserved1", wintypes.WORD),
        ("dwSupport", wintypes.DWORD),
    ]

def get_audio_device_name():
    """현재 시스템의 기본 오디오 출력 장치 이름을 가져옵니다."""
    try:
        wave_caps = WAVEOUTCAPS()
        res = ctypes.windll.winmm.waveOutGetDevCapsW(0, ctypes.byref(wave_caps), ctypes.sizeof(wave_caps))
        if res == 0:
            return wave_caps.szPname
        return "장치 이름을 찾을 수 없음"
    except:
        return "Windows 오디오 서비스 접근 불가"

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(700, 850)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton_sound = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_sound.setGeometry(QtCore.QRect(50, 30, 270, 60))
        self.pushButton_sound.setObjectName("pushButton_sound")
        self.pushButton_activation = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_activation.setGeometry(QtCore.QRect(340, 30, 270, 60))
        self.pushButton_activation.setObjectName("pushButton_activation")
        self.pushButton_specs = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_specs.setGeometry(QtCore.QRect(50, 110, 270, 60))
        self.pushButton_specs.setObjectName("pushButton_specs")
        self.pushButton_faulty = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_faulty.setGeometry(QtCore.QRect(340, 110, 270, 60))
        self.pushButton_faulty.setObjectName("pushButton_faulty")
        self.textBrowser = QtWidgets.QTextBrowser(self.centralwidget)
        self.textBrowser.setGeometry(QtCore.QRect(50, 200, 561, 600))
        self.textBrowser.setObjectName("textBrowser")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 700, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "TAK 진단 도구"))
        self.pushButton_sound.setText(_translate("MainWindow", "윈도우 사운드 테스트"))
        self.pushButton_activation.setText(_translate("MainWindow", "윈도우 정품 인증 확인"))
        self.pushButton_specs.setText(_translate("MainWindow", "시스템 사양 확인"))
        self.pushButton_faulty.setText(_translate("MainWindow", "장치 관리자 오류 점검"))

class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.pushButton_sound.clicked.connect(self.test_windows_sound)
        self.pushButton_activation.clicked.connect(self.check_windows_activation)
        self.pushButton_specs.clicked.connect(self.get_system_specs)
        self.pushButton_faulty.clicked.connect(self.check_faulty_devices)

    def append_text(self, text):
        self.textBrowser.append(text)

    def test_windows_sound(self):
        device_name = get_audio_device_name()
        self.append_text("="*40)
        self.append_text(f"📡 감지된 장치: {device_name}")
        self.append_text("="*40)
        self.append_text("=== 윈도우 사운드 장치 테스트를 시작합니다 ===")

        # 1. 기본 비프음 테스트
        self.append_text("\n[1단계] 기본 비프음(Beep) 테스트 중...")
        try:
            winsound.Beep(440, 500)
            time.sleep(0.5)
            winsound.Beep(880, 500)
            self.append_text("-> 비프음 재생 성공")
        except Exception as e:
            self.append_text(f"-> 비프음 오류: {e}")

        # 2. 윈도우 시스템 사운드 테스트
        self.append_text("\n[2단계] 시스템 경고음(System Hand) 테스트 중...")
        try:
            winsound.MessageBeep(winsound.MB_ICONHAND)
            time.sleep(1)
            self.append_text("-> 시스템 경고음 재생 요청 완료")
        except Exception as e:
            self.append_text(f"-> 시스템 소리 오류: {e}")

        # 3. 주파수 스윕 테스트
        self.append_text("\n[3단계] 주파수 상승 테스트 (500Hz -> 1500Hz)")
        try:
            for freq in range(500, 1501, 250):
                self.append_text(f"현재 주파수: {freq}Hz")
                winsound.Beep(freq, 300)
                time.sleep(0.1)
            self.append_text("-> 주파수 테스트 완료")
        except Exception as e:
            self.append_text(f"-> 주파수 테스트 오류: {e}")

        self.append_text("\n=== 모든 테스트가 완료되었습니다 ===")

    def check_windows_activation(self):
        self.append_text("🔎 윈도우 정품 인증 상태를 확인하는 중...")
        try:
            result = subprocess.run(
                ["cscript", "//nologo", "C:\\Windows\\System32\\slmgr.vbs", "/xpr"],
                capture_output=True,
                text=True,
                encoding='cp949'
            )
            if result.returncode == 0:
                status_info = result.stdout.strip()
                self.append_text("-" * 40)
                self.append_text(f"인증 상태 결과:\n{status_info}")
                self.append_text("-" * 40)
                if "영구적" in status_info or "permanently" in status_info.lower():
                    self.append_text("✅ 이 윈도우는 영구적으로 인증되었습니다.")
                else:
                    self.append_text("⚠️ 정품 인증이 필요하거나 기간 제한 라이선스입니다.")
            else:
                self.append_text("❌ 상태 정보를 가져오지 못했습니다.")
        except Exception as e:
            self.append_text(f"❌ 오류 발생: {e}")

    def get_system_specs(self):
        self.append_text("="*60)
        self.append_text("      [ 현재 컴퓨터 하드웨어 사양 ]")
        self.append_text("="*60)
        try:
            cpu_cmd = "wmic cpu get name"
            cpu_name = subprocess.check_output(cpu_cmd, shell=True).decode('cp949').split('\n')[1].strip()
            self.append_text(f"🖥️  CPU: {cpu_name}")

            ram_cmd = "wmic computersystem get totalphysicalmemory"
            ram_raw = subprocess.check_output(ram_cmd, shell=True).decode('cp949').split('\n')[1].strip()
            ram_gb = round(int(ram_raw) / (1024**3), 1)
            self.append_text(f"Memory: {ram_gb} GB")

            vga_cmd = "wmic path win32_VideoController get name"
            vga_output = subprocess.check_output(vga_cmd, shell=True).decode('cp949').split('\n')
            vga_names = [line.strip() for line in vga_output[1:] if line.strip()]
            for i, name in enumerate(vga_names, 1):
                self.append_text(f"🎬  VGA {i}: {name}")

            disk_output = subprocess.check_output("wmic diskdrive get model, size", shell=True).decode('cp949').split('\n')
            self.append_text("💾  Storage (SSD/HDD):")
            for line in disk_output[1:]:
                if line.strip():
                    parts = line.strip().rsplit(None, 1)
                    if len(parts) == 2:
                        model = parts[0]
                        size_gb = round(int(parts[1]) / (1024**3), 0)
                        self.append_text(f"    - {model} ({size_gb} GB)")
        except Exception as e:
            self.append_text(f"❌ 사양을 가져오는 중 오류가 발생했습니다: {e}")
        self.append_text("="*60)

    def check_faulty_devices(self):
        self.append_text("="*60)
        self.append_text("      [ 장치 관리자 오류 장치 점검 (! 또는 ? 상태) ]")
        self.append_text("="*60)
        try:
            cmd = 'wmic path Win32_PnPEntity where "ConfigManagerErrorCode <> 0" get Name, ConfigManagerErrorCode'
            result = subprocess.check_output(cmd, shell=True).decode('cp949', errors='ignore')
            lines = result.strip().split('\n')[1:]
            faulty_list = [line.strip() for line in lines if line.strip()]
            if not faulty_list:
                self.append_text("✅ 확인 완료: 문제 있는 장치가 없습니다. (모든 드라이버 정상)")
            else:
                self.append_text(f"⚠️ 총 {len(faulty_list)}개의 문제가 있는 장치를 발견했습니다:\n")
                self.append_text(f"{'장치 이름':<40} | {'오류 코드'}")
                self.append_text("-" * 60)
                for item in faulty_list:
                    parts = item.rsplit(None, 1)
                    if len(parts) == 2:
                        name, code = parts
                        self.append_text(f"{name[:40]:<40} | {code}")
                self.append_text("\n💡 팁: 오류 코드 28은 드라이버 미설치, 43은 장치 연결 오류인 경우가 많습니다.")
        except Exception as e:
            self.append_text(f"❌ 장치 정보를 읽어오는 중 오류 발생: {e}")

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())