import os
import subprocess
import pyautogui
import time
from pathlib import Path
import pygetwindow as gw

# 사용자 홈 디렉터리의 바탕화면 경로
user_profile = Path.home()
chrome_extension_path = os.path.join(user_profile, 'Desktop')  # 바탕화면 경로
zip_file_name = '강종.bat'  # 복사될 파일 이름
zip_file_path = os.path.abspath('./zip/강종.bat')  # 절대 경로로 변환
desktop_file_path = os.path.join(chrome_extension_path, zip_file_name)  # 바탕화면에 복사될 파일 경로

# 바탕화면에 이미 복사된 파일이 있다면 삭제하는 함수
def clear_existing_file(file_path):
    subprocess.run('echo off | clip', shell=True) # 클립보드 초기화
    if os.path.exists(file_path):
        os.remove(file_path)
        print(f"{file_path} 파일이 삭제되었습니다.")
    else:
        print(f"{file_path} 파일이 존재하지 않습니다. 새로 복사 작업을 시작합니다.")

# 파일을 탐색기에서 선택하기 위한 함수 (subprocess 사용)
def select_file_in_explorer(file_path):
    # 'explorer' 명령어로 파일 선택 (탐색기에서 해당 파일을 열고 선택)
    pyautogui.hotkey('win', 'd')  # 바탕화면으로 이동 (Windows 키 + D)
    subprocess.run(['explorer', '/select,', file_path])
    time.sleep(2)  # 탐색기가 열릴 시간을 대기

# 탐색기의 보기 방식을 "자세히 보기"로 변경하는 함수
def set_view_to_details():
    pyautogui.click(button='right')  # 파일 위에서 마우스 우클릭
    time.sleep(1)
    pyautogui.press('v')  # 보기 메뉴 열기
    time.sleep(0.5)
    pyautogui.press('d')  # "자세히 보기" 선택
    time.sleep(1)


# 기존 파일 삭제
clear_existing_file(desktop_file_path)

# 파일을 선택
select_file_in_explorer(zip_file_path)

windows = gw.getWindowsWithTitle("zip")  # 'zip' 제목을 가진 창 찾기

if windows:
        # 첫 번째 탐색기 창을 활성화 및 최상위로 전환
        explorer_window = windows[0]
        explorer_window.activate()

pyautogui.hotkey('shift', 'f10')

# 우클릭 메뉴에서 '복사' 선택
time.sleep(1)
for _ in range(5):  # '복사' 항목으로 이동하기 위해 5번 'up' 키 누름
    pyautogui.press('up')

pyautogui.press('enter')  # '복사' 선택

# 탐색기 창 닫기
for window in windows:
    window.close()

# 바탕화면으로 이동 후 붙여넣기
time.sleep(1)
pyautogui.hotkey('win', 'd')  # 바탕화면으로 이동 (Windows 키 + D)
time.sleep(1)

# 바탕화면에서 우클릭하여 붙여넣기
pyautogui.rightClick(pyautogui.size().width // 2 - 350, pyautogui.size().height // 2 + 150)  # 화면 중앙에서 우클릭
time.sleep(1)
for _ in range(5):  # 붙여넣기 메뉴로 이동
    pyautogui.press('down')

pyautogui.press('enter')  # 붙여넣기 선택

print("파일이 바탕화면에 복사되었습니다.")
