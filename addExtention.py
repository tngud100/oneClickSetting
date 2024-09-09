from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from pywinauto import Application
import pyautogui
import pyperclip
import time
import os


################ 익스텐션 설치 ################


# Chrome 옵션 설정
os.system("taskkill /F /IM chrome.exe")
user_name = os.getlogin()
chrome_options = Options()
chrome_options.add_argument(rf"--user-data-dir=C:\\Users\\{user_name}\\AppData\\Local\\Google\\Chrome\\User Data")  # 기본 사용자 프로필 경로 지정
chrome_options.add_argument("--profile-directory=Default")  # 기본 프로필 지정
chrome_options.add_argument("--start-maximized")  # 창 최대화
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

# Selenium으로 Chrome 실행
driver = webdriver.Chrome(options=chrome_options)

# 'chrome://extensions' 페이지 열기
driver.get('chrome://extensions')

# 페이지가 로드될 때까지 대기 (최대 10초)
wait = WebDriverWait(driver, 10)

# Shadow DOM에서 개발자 모드 토글 접근 및 상태 확인 후 클릭
try:
    # Shadow DOM 접근을 위해 JavaScript 실행 (devMode 토글을 찾음)
    dev_mode_toggle = driver.execute_script('''
        return document.querySelector('extensions-manager').shadowRoot
        .querySelector('extensions-toolbar').shadowRoot
        .querySelector('#devMode');
    ''')

    # aria-pressed 속성이 false일 때 클릭
    if dev_mode_toggle and dev_mode_toggle.get_attribute("aria-pressed") == "false":
        dev_mode_toggle.click()
        print("개발자 모드를 활성화했습니다.")
    else:
        print("개발자 모드는 이미 활성화되어 있습니다.")

except Exception as e:
    print(f"개발자 모드를 설정하는 도중 오류 발생: {e}")

time.sleep(1)
# 개발자 모드가 활성화된 이후 '압축해제된 확장 프로그램을 로드합니다' 버튼을 클릭
try:
    # "압축해제된 확장 프로그램을 로드합니다" 버튼이 나타날 때까지 대기
    load_unpacked_button = wait.until(lambda d: driver.execute_script('''
        return document.querySelector('extensions-manager').shadowRoot
        .querySelector('extensions-toolbar').shadowRoot
        .querySelector('#loadUnpacked');
    '''))

    # 버튼이 있으면 클릭
    if load_unpacked_button:
        load_unpacked_button.click()
        print("압축해제된 확장 프로그램을 로드합니다 버튼을 클릭했습니다.")
    else:
        print("압축해제된 확장 프로그램을 로드합니다 버튼을 찾을 수 없습니다.")

except Exception as e:
    print(f"압축해제된 확장 프로그램을 로드합니다 버튼을 클릭하는 도중 오류 발생: {e}")

# 클립보드에서 경로 읽기
directory_path = pyperclip.paste()+"\\2.1_0"

# 파일 탐색기에서 경로 입력
pyautogui.hotkey('alt', 'tab')  # 주소창으로 이동
pyautogui.hotkey('ctrl', 'l')  # 주소창으로 이동
pyperclip.copy(directory_path)  # 클립보드에 있는 경로 사용
pyautogui.hotkey('ctrl', 'v')  # 붙여넣기

time.sleep(2)  # 경로가 붙여넣기될 시간을 기다림
pyautogui.press('enter')  # Enter 키를 눌러 디렉토리 선택
time.sleep(0.5)  # 경로가 붙여넣기될 시간을 기다림
pyautogui.press('enter')  # Enter 키를 눌러 디렉토리 선택
time.sleep(0.5)  # 경로가 붙여넣기될 시간을 기다림
pyautogui.press('enter')  # Enter 키를 눌러 디렉토리 선택

# input("Press Enter to exit and close the browser...")
time.sleep(3)  # 익스테션 로드를 기다림

driver.get('chrome-extension://dnoagfebjndkhkabjkkoeeijnjpmbimj/options.html')

time.sleep(1)

# input 필드에 값 입력 (id="AddingTitle")
input_element = driver.find_element(By.ID, 'AddingTitle')
input_element.send_keys('any')

# select 필드에서 값 선택 (id="AddingType")
select_element = Select(driver.find_element(By.ID, 'AddingType'))
select_element.select_by_index(1)  # 인덱스 1에 위치한 옵션 선택

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

driver.execute_script(script)

add_edit_button = driver.find_element(By.ID, 'AddEditButton')
add_edit_button.click()

# 작업이 끝나면 브라우저 종료 대기
input("Press Enter to exit and close the browser...")
driver.quit()