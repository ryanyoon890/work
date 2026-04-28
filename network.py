
import subprocess
import ctypes
import sys

def is_admin():
    """관리자 권한 여부를 확인합니다."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def show_network_interfaces():
    """현재 시스템에 있는 네트워크 인터페이스 목록을 보여줍니다."""
    print("\n--- 현재 네트워크 어댑터 목록 ---")
    try:
        # netsh를 이용해 인터페이스 이름만 추출
        result = subprocess.check_output("netsh interface show interface", shell=True).decode('cp949')
        print(result)
    except Exception as e:
        print(f"목록을 가져오는 중 오류 발생: {e}")

def change_interface_name():
    if not is_admin():
        print("\n❌ 오류: 이 코드는 반드시 '관리자 권한으로 실행'해야 합니다.")
        print("팁: 파이썬이나 터미널을 '관리자 권한으로 실행' 후 다시 시도해 주세요.")
        return

    # 1. 목록 확인
    show_network_interfaces()

    # 2. 사용자 입력 받기
    print("변경을 원치 않으시면 Ctrl+C를 눌러 종료하세요.")
    old_name = input("\n[1] 변경할 현재 이름을 입력하세요 (예: 이더넷): ").strip()
    new_name = input("[2] 새로 바꿀 이름을 입력하세요: ").strip()

    if not old_name or not new_name:
        print("❌ 이름은 빈 칸일 수 없습니다.")
        return

    # 3. 변경 실행
    print(f"\n🔄 실행 중: [{old_name}] -> [{new_name}]...")
    cmd = f'netsh interface set interface name="{old_name}" newname="{new_name}"'
    
    try:
        # 명령 실행
        result = subprocess.run(["powershell", "-Command", cmd], capture_output=True, text=True)
        
        if result.returncode == 0:
            print(f"✅ 성공: 네트워크 이름이 '{new_name}'(으)로 변경되었습니다.")
        else:
            print(f"❌ 실패: 이름을 변경할 수 없습니다.")
            print(f"이유: {result.stderr if result.stderr else '이름이 틀렸거나 이미 존재하는 이름일 수 있습니다.'}")
            
    except Exception as e:
        print(f"❌ 예외 발생: {e}")

if __name__ == "__main__":
    change_interface_name()