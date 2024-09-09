import os
import time
import pyperclip
import pyautogui
from shutil import which
import shutil
import zipfile
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
import subprocess
import pygetwindow as gw


######################
# 확장 프로그램 다운로드 #
######################
def setup_chrome_extension(target_path, zip_file_path):
    """Chrome 확장 프로그램을 다운로드"""
    try:
        # 지정된 경로가 없으면 생성
        os.makedirs(target_path, exist_ok=True)
        
        # ZIP 파일을 대상 경로로 복사
        destination_zip = os.path.join(target_path, os.path.basename(zip_file_path))
        shutil.copy2(zip_file_path, destination_zip)

        # ZIP 파일 압축 해제
        with zipfile.ZipFile(destination_zip, 'r') as zip_ref:
            zip_ref.extractall(target_path)

        # 압축 해제 후 클립보드에 경로 복사 (Windows 전용)
        if os.name == 'nt':
            os.system(f'echo {target_path + '\\2.1_0'} | clip')

        print(f'Chrome 확장 프로그램이 {target_path}에 설치되었습니다.')

    except PermissionError:
        print(f"PermissionError: {target_path}에 대한 권한이 없습니다.")
    except FileNotFoundError:
        print(f"FileNotFoundError: {zip_file_path} 파일을 찾을 수 없습니다.")
    except zipfile.BadZipFile:
        print(f"BadZipFileError: {zip_file_path} 파일이 유효한 ZIP 파일이 아닙니다.")
    except Exception as e:
        print(f"오류 발생: {str(e)}")


######################
# 기본 Chrome 설정 및 확장 프로그램 설치 #
######################

def setup_chrome_options(user_profile='Default'):
    """Chrome 옵션을 설정하고 프로필을 선택하여 반환"""
    user_name = os.getlogin()
    chrome_options = Options()
    chrome_options.add_argument(f"--user-data-dir=C:\\Users\\{user_name}\\AppData\\Local\\Google\\Chrome\\User Data")
    chrome_options.add_argument(f"--profile-directory={user_profile}")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    return chrome_options

def enable_developer_mode(driver):
    """크롬 확장 프로그램 개발자 모드 활성화"""
    try:
        dev_mode_toggle = driver.execute_script('''
            return document.querySelector('extensions-manager').shadowRoot
            .querySelector('extensions-toolbar').shadowRoot
            .querySelector('#devMode');
        ''')
        if dev_mode_toggle and dev_mode_toggle.get_attribute("aria-pressed") == "false":
            dev_mode_toggle.click()
            print("개발자 모드를 활성화했습니다.")
        else:
            print("개발자 모드는 이미 활성화되어 있습니다.")
    except Exception as e:
        print(f"개발자 모드를 설정하는 도중 오류 발생: {e}")

def load_extension(driver):
    """압축해제된 확장 프로그램을 로드하는 버튼 클릭"""
    wait = WebDriverWait(driver, 10)
    try:
        load_unpacked_button = wait.until(lambda d: driver.execute_script('''
            return document.querySelector('extensions-manager').shadowRoot
            .querySelector('extensions-toolbar').shadowRoot
            .querySelector('#loadUnpacked');
        '''))
        load_unpacked_button.click()
        print("압축해제된 확장 프로그램을 로드합니다 버튼을 클릭했습니다.")
    except Exception as e:
        print(f"압축해제된 확장 프로그램을 로드합니다 버튼을 클릭하는 도중 오류 발생: {e}")

def select_extension_folder():
    """PyAutoGUI를 사용하여 파일 탐색기에서 확장 프로그램 경로 선택"""
    directory_path = pyperclip.paste()
    
    time.sleep(1)
    # pyautogui.hotkey('alt', 'tab')  # 탐색기 창으로 전환
    # 주소 입력란에 포커스 맞추고 경로 복사 후 붙여넣기
    pyautogui.hotkey('ctrl', 'l')  # 주소 입력란으로 이동
    pyperclip.copy(directory_path)  # 클립보드에 경로 복사
    pyautogui.hotkey('ctrl', 'v')   # 경로 붙여넣기
    time.sleep(1)
    # 경로 입력 후 엔터
    pyautogui.press('enter')
    time.sleep(0.5)
    # 두 번째 확인 창에서 엔터 (디렉토리 선택)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.press('enter')


def execute_script_in_extension(driver, script):
    """확장 프로그램 내에서 스크립트 inject"""
    input_element = driver.find_element(By.ID, 'AddingTitle')
    input_element.send_keys('any')
    select_element = Select(driver.find_element(By.ID, 'AddingType'))
    select_element.select_by_index(1)
    driver.execute_script(script)
    time.sleep(2)
    driver.find_element(By.ID, 'AddEditButton').click()

####################
# 강종 bat파일 바탕화면에 복사 #
####################

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


####################
# Chrome 창 관리 함수 #
####################

def find_chrome_path():
    """Chrome 경로를 검색하여 반환"""
    potential_paths = [
        r'C:\Program Files\Google\Chrome\Application\chrome.exe',
        r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
    ]
    chrome_path = which("chrome")
    if chrome_path:
        return chrome_path
    for path in potential_paths:
        if os.path.exists(path):
            return path
    raise FileNotFoundError("Chrome executable not found.")

def open_chrome_windows(chrome_path, user_data_dir, urls_window_sets):
    """Chrome 창을 여러 개 열기"""
    for urls in urls_window_sets:
        subprocess.Popen([chrome_path, '--new-window', f'--user-data-dir={user_data_dir}', '--profile-directory=Default'] + urls)

def set_window_positions(window_titles):
    """각 Chrome 창의 위치 및 크기 설정"""
    windows = gw.getWindowsWithTitle(window_titles)
    if len(windows) >= 1:
        windows[0].restore()
        windows[0].resizeTo(560, 1050)
        windows[0].moveTo(-7, 0)
    if len(windows) >= 2:
        windows[1].restore()
        windows[1].resizeTo(650, 450)
        windows[1].moveTo(538, 0)
    # if len(windows) >= 3:
    #     windows[2].restore()
    #     windows[2].resizeTo(650, 610)
    #     windows[2].moveTo(538, 440)

##################
# 전체 실행 프로세스 #
##################

def main():
    os.system("taskkill /F /IM chrome.exe")
    user_name = os.getlogin()
    user_data_dir = f"C:\\Users\\{user_name}\\AppData\\Local\\Google\\Chrome\\User Data"
    extension_path = f"C:\\Users\\{user_name}\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Extensions\\dnoagfebjndkhkabjkkoeeijnjpmbimj"
    zip_file_path = './zip/dnoagfebjndkhkabjkkoeeijnjpmbimj.zip'
    bat_file_path = os.path.abspath('./zip/강종.bat')
    desktop_file_path = f"C:\\Users\\{user_name}\\Desktop\\강종.bat"


    setup_chrome_extension(extension_path, zip_file_path)

    chrome_options = setup_chrome_options()
    
    # ChromeDriver 서비스 자동 설정
    service = Service(ChromeDriverManager().install())

    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    driver.get('chrome://extensions')
    enable_developer_mode(driver)
    time.sleep(2)
    
    load_extension(driver)
    select_extension_folder()
    time.sleep(3)
    
    # 스크립트 실행 (사용자 정의 script는 외부에서 전달받을 수 있음)
    script = """
    var editor = document.querySelector('.CodeMirror').CodeMirror;
    editor.setValue(`

    window.addEventListener('message', function (e) {
        if (e.data.indexOf("e4setting") !== -1) {
            // handle e4setting case
        } else if (e.data.indexOf("e1nexon_") !== -1) {
            document.querySelector("#txtNexonID").value = e.data.replace("e1nexon_", "");
        } else if (e.data.indexOf("e2nexon_") !== -1) {
            document.querySelector("#txtPWD").value = e.data.replace("e2nexon_", "");
            document.querySelector(".button01").click();
        } else if (e.data == "e3nexon_") {
            GnxGameStartOnClick();
        } else if (e.data.indexOf("e0naver_") !== -1) {
            document.querySelector(".btNaver").click();
        } else if (e.data.indexOf("e3naver_") !== -1) {
            document.querySelector("#qrlog\\.cancel").click();
        } else if (e.data.indexOf("e1naver_") !== -1) {
            document.querySelector("#id").value = e.data.replace("e1naver_", "");
        } else if (e.data.indexOf("e2naver_") !== -1) {
            document.querySelector("#pw").value = e.data.replace("e2naver_", "");
        } else if (e.data.indexOf("e4naver_") !== -1) {
            document.querySelector("#log\\.login").click();
        } else if (e.data.indexOf("e0google_") !== -1) {
            document.querySelector(".btGoogle").click();
        } else if (e.data.indexOf("e1google_") !== -1) {
            try {
                document.querySelector(".UXurCe").click();
            } catch {}
            try {
                document.querySelector(".jbiPi").click();
            } catch {}
        } else if (e.data.indexOf("e2google_") !== -1) {
            document.querySelector("#identifierId").value = e.data.replace("e2google_", "");
            document.querySelector(".VfPpkd-LgbsSe.VfPpkd-LgbsSe-OWXEXe-k8QpJ").click();
        } else if (e.data.indexOf("e3google_") !== -1) {
            document.querySelector(".whsOnd.zHQkBf").value = e.data.replace("e3google_", "");
            document.querySelector(".VfPpkd-LgbsSe.VfPpkd-LgbsSe-OWXEXe-k8QpJ").click();
        } else if (e.data.indexOf("e4google_") !== -1) {
            if (window.location.hostname == "accounts.google.com") {
                return;
            } else {
                location.href = "http://fifaonline4.nexon.com";
            }
        } else if (e.data.indexOf("e7facebook_") !== -1) {
            document.querySelector('.hu5pjgll.lzf7d6o1.sp_AvOBU6yv9AO.sx_a06bd2').click();
        } else if (e.data.indexOf("e8facebook_") !== -1) {
            document.querySelector('.hu5pjgll.lzf7d6o1.sp_H7axIjqrTST.sx_332863').click();
        } else if (e.data.indexOf("e4facebook_") !== -1) {
            document.querySelector("#mbasic_logout_button").click();
        } else if (e.data.indexOf("e5facebook_") !== -1) {
            location.reload();
        } else if (e.data.indexOf("e6facebook_") !== -1) {
            document.querySelector(".q.r.s.t.u.y.w.z").click();
        } else if (e.data.indexOf("e0facebook_") !== -1) {
            document.querySelector(".btFacebook").click();
        } else if (e.data.indexOf("e2facebook_") !== -1) {
            document.querySelector("#email").value = e.data.replace("e2facebook_", "");
        } else if (e.data.indexOf("e3facebook_") !== -1) {
            document.querySelector("#pass").value = e.data.replace("e3facebook_", "");
            document.querySelector("#loginbutton").click();
        } else if (e.data.indexOf("e0apple") !== -1) {
            document.querySelector(".btApple").click();
        }
    });
    `);
    """
    driver.get('chrome-extension://dnoagfebjndkhkabjkkoeeijnjpmbimj/options.html')
    execute_script_in_extension(driver, script)
    time.sleep(3)
    driver.quit()

    # 기존 파일 삭제
    clear_existing_file(desktop_file_path)

    # 파일을 선택
    select_file_in_explorer(bat_file_path)

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
    pyautogui.rightClick(pyautogui.size().width // 2 - 350, pyautogui.size().height // 2 + 50)  # 화면 중앙에서 우클릭
    time.sleep(1)
    for _ in range(5):  # 붙여넣기 메뉴로 이동
        pyautogui.press('down')

    pyautogui.press('enter')  # 붙여넣기 선택

    print("파일이 바탕화면에 복사되었습니다.")


    chrome_path = find_chrome_path()
    urls_window_sets = [
        ['https://피파대낙.com/main.jsp', 'https://심플대낙.com/worker'],
        ['https://피파대낙.com/main.jsp', 'https://심플대낙.com/worker'],
        # ['https://피파대낙.com/main.jsp', 'https://심플대낙.com/worker']
    ]

    open_chrome_windows(chrome_path, user_data_dir, urls_window_sets)
    time.sleep(2)
    set_window_positions('인터넷')

if __name__ == "__main__":
    main()
