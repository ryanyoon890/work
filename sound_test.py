import winsound
import time

print("윈도우 기본 소리(경고음) 3회 재생 테스트")
for i in range(3):
    print(f"{i+1}번째 소리 재생...")
    winsound.MessageBeep()
    time.sleep(1)
print("테스트 완료!")
