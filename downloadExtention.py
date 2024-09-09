import os
import shutil
import zipfile
from pathlib import Path

def setup_chrome_extension(target_path, zip_file_path):
    """Chrome 확장 프로그램을 다운로드 및 설치"""
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
            os.system(f'echo {target_path} | clip')

        print(f'Chrome 확장 프로그램이 {target_path}에 설치되었습니다.')

    except PermissionError:
        print(f"PermissionError: {target_path}에 대한 권한이 없습니다.")
    except FileNotFoundError:
        print(f"FileNotFoundError: {zip_file_path} 파일을 찾을 수 없습니다.")
    except zipfile.BadZipFile:
        print(f"BadZipFileError: {zip_file_path} 파일이 유효한 ZIP 파일이 아닙니다.")
    except Exception as e:
        print(f"오류 발생: {str(e)}")

def main():
    user_profile = Path.home()
    chrome_extension_path = os.path.join(user_profile, r'AppData\Local\Google\Chrome\User Data\Default\Extensions\dnoagfebjndkhkabjkkoeeijnjpmbimj')
    zip_file_path = os.path.abspath('./zip/dnoagfebjndkhkabjkkoeeijnjpmbimj.zip')  # 절대 경로로 변환
    
    if not os.path.exists(zip_file_path):
        print(f"지정한 경로에 ZIP 파일이 없습니다: {zip_file_path}")
        return
    
    setup_chrome_extension(chrome_extension_path, zip_file_path)

if __name__ == "__main__":
    main()
