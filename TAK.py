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
        # waveOutGetDevCaps: 오디오 출력 장치의 성능과 이름을 가져오는 WinAPI
        res = ctypes.windll.winmm.waveOutGetDevCapsW(0, ctypes.byref(wave_caps), ctypes.sizeof(wave_caps))
        if res == 0:
            return wave_caps.szPname
        return "장치 이름을 찾을 수 없음"
    except:
        return "Windows 오디오 서비스 접근 불가"
    
    
def test_windows_sound():
    device_name = get_audio_device_name()
    
    print("="*40)
    print(f"📡 감지된 장치: {device_name}")
    print("="*40)
    print("=== 윈도우 사운드 장치 테스트를 시작합니다 ===")

    # 1. 기본 비프음 테스트 (가장 낮은 단계의 사운드 출력)
    print("\n[1단계] 기본 비프음(Beep) 테스트 중...")
    try:
        # Beep(주파수, 지속시간)
        winsound.Beep(440, 500)  # '라' 음, 0.5초
        time.sleep(0.5)
        winsound.Beep(880, 500)  # 한 옥타브 높은 '라'
        print("-> 비프음 재생 성공")
    except Exception as e:
        print(f"-> 비프음 오류: {e}")

    # 2. 윈도우 시스템 사운드 테스트
    # 실제 윈도우 설정에 지정된 시스템 소리를 재생합니다.
    print("\n[2단계] 시스템 경고음(System Hand) 테스트 중...")
    try:
        # winsound.MessageBeep()은 윈도우 스피커 설정에 따라 소리가 다를 수 있습니다.
        winsound.MessageBeep(winsound.MB_ICONHAND)
        time.sleep(1)
        print("-> 시스템 경고음 재생 요청 완료")
    except Exception as e:
        print(f"-> 시스템 소리 오류: {e}")

    # 3. 주파수 스윕(Sweep) 테스트 (스피커 가청 범위 확인)
    print("\n[3단계] 주파수 상승 테스트 (500Hz -> 1500Hz)")
    try:
        for freq in range(500, 1501, 250):
            print(f"현재 주파수: {freq}Hz")
            winsound.Beep(freq, 300)
            time.sleep(0.1)
        print("-> 주파수 테스트 완료")
    except Exception as e:
        print(f"-> 주파수 테스트 오류: {e}")

    print("\n=== 모든 테스트가 완료되었습니다 ===")
    
def check_windows_activation():
    print("🔎 윈도우 정품 인증 상태를 확인하는 중...")
    
    try:
        # /xpr: 라이선스 만료 날짜 및 현재 인증 상태 확인 (팝업창 출력)
        # /dli: 더 자세한 라이선스 정보 확인
        result = subprocess.run(
            ["cscript", "//nologo", "C:\\Windows\\System32\\slmgr.vbs", "/xpr"], 
            capture_output=True, 
            text=True, 
            encoding='cp949' # 한글 윈도우 인코딩 대응
        )
        
        if result.returncode == 0:
            status_info = result.stdout.strip()
            print("-" * 40)
            print(f"인증 상태 결과:\n{status_info}")
            print("-" * 40)
            
            if "영구적" in status_info or "permanently" in status_info.lower():
                print("✅ 이 윈도우는 영구적으로 인증되었습니다.")
            else:
                print("⚠️ 정품 인증이 필요하거나 기간 제한 라이선스입니다.")
        else:
            print("❌ 상태 정보를 가져오지 못했습니다.")
            
    except Exception as e:
        print(f"❌ 오류 발생: {e}")

def get_system_specs():
    print("="*60)
    print("      [ 현재 컴퓨터 하드웨어 사양 ]")
    print("="*60)

    try:
        # 1. CPU 사양 가져오기
        cpu_cmd = "wmic cpu get name"
        cpu_name = subprocess.check_output(cpu_cmd, shell=True).decode('cp949').split('\n')[1].strip()
        print(f"🖥️  CPU: {cpu_name}")

        # 2. RAM 사양 가져오기 (Bytes -> GB 변환)
        ram_cmd = "wmic computersystem get totalphysicalmemory"
        ram_raw = subprocess.check_output(ram_cmd, shell=True).decode('cp949').split('\n')[1].strip()
        ram_gb = round(int(ram_raw) / (1024**3), 1)
        print(f"Memory: {ram_gb} GB")

        # 3. VGA(그래픽 카드) 사양 가져오기
        vga_cmd = "wmic path win32_VideoController get name"
        vga_output = subprocess.check_output(vga_cmd, shell=True).decode('cp949').split('\n')
        # 그래픽 카드가 여러 개인 경우(내장/외장) 모두 출력
        vga_names = [line.strip() for line in vga_output[1:] if line.strip()]
        for i, name in enumerate(vga_names, 1):
            print(f"🎬  VGA {i}: {name}")
            
        # 4. SSD/HDD(저장 장치) 사양 추가
        # 모델명과 크기(Size)를 가져옵니다.
        disk_output = subprocess.check_output("wmic diskdrive get model, size", shell=True).decode('cp949').split('\n')
        print("💾  Storage (SSD/HDD):")
        for line in disk_output[1:]:
            if line.strip():
                # 뒤쪽 숫자가 용량(Byte), 앞쪽이 모델명
                parts = line.strip().rsplit(None, 1)
                if len(parts) == 2:
                    model = parts[0]
                    size_gb = round(int(parts[1]) / (1024**3), 0)
                    print(f"    - {model} ({size_gb} GB)")

    except Exception as e:
        print(f"❌ 사양을 가져오는 중 오류가 발생했습니다: {e}")

    print("="*60)

def check_faulty_devices():
    print("="*60)
    print("      [ 장치 관리자 오류 장치 점검 (! 또는 ? 상태) ]")
    print("="*60)
    
    try:
        # ConfigManagerErrorCode가 0이 아니면 장치 관리자에서 경고가 뜬 장치입니다.
        # 0: 정상 작동
        # 28: 드라이버 미설치 (물음표)
        # 10, 43: 장치 오류/중지 (느낌표) 등
        cmd = 'wmic path Win32_PnPEntity where "ConfigManagerErrorCode <> 0" get Name, ConfigManagerErrorCode'
        result = subprocess.check_output(cmd, shell=True).decode('cp949', errors='ignore')
        
        lines = result.strip().split('\n')[1:] # 헤더 제외
        
        faulty_list = [line.strip() for line in lines if line.strip()]
        
        if not faulty_list:
            print("✅ 확인 완료: 문제 있는 장치가 없습니다. (모든 드라이버 정상)")
        else:
            print(f"⚠️ 총 {len(faulty_list)}개의 문제가 있는 장치를 발견했습니다:\n")
            print(f"{'장치 이름':<40} | {'오류 코드'}")
            print("-" * 60)
            for item in faulty_list:
                # 마지막 공백을 기준으로 이름과 코드를 분리
                parts = item.rsplit(None, 1)
                if len(parts) == 2:
                    name, code = parts
                    print(f"{name[:40]:<40} | {code}")
            
            print("\n💡 팁: 오류 코드 28은 드라이버 미설치, 43은 장치 연결 오류인 경우가 많습니다.")
            
    except Exception as e:
        print(f"❌ 장치 정보를 읽어오는 중 오류 발생: {e}")

if __name__ == "__main__":
    test_windows_sound()
    check_windows_activation()
    get_system_specs()
    check_faulty_devices()
