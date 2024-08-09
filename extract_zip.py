import os
import zipfile
import subprocess
import pandas as pd


def extract_zip(zip_path, extract_to_folder):
    """주어진 zip 파일을 지정된 폴더로 압축 해제합니다."""
    try:
        # 압축 해제할 폴더 생성
        zip_folder = os.path.join(
            extract_to_folder, os.path.splitext(os.path.basename(zip_path))[0])
        os.makedirs(zip_folder, exist_ok=True)

        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(zip_folder)
            # 파일 리스트 수집
            file_list = zip_ref.namelist()

        print(f"{zip_path}을/를 {zip_folder}로 압축 해제했습니다.")

        # 압축 해제 후 zip 파일 삭제
        os.remove(zip_path)

        # 압축 해제 후 빈 폴더 삭제
        if not os.listdir(zip_folder):  # 폴더가 비어있으면
            os.rmdir(zip_folder)

        return file_list, zip_folder
    except Exception as e:
        print(f"zip 압축해제 오류 : {e}")
        except_log(zip_path, e)
        return [], ""


def extract_egg(egg_path):
    """주어진 egg 파일을 Bandizip을 사용하여 압축 해제합니다."""
    try:
        # 압축 해제할 폴더 생성
        egg_folder = os.path.join(os.path.dirname(
            egg_path), os.path.splitext(os.path.basename(egg_path))[0])
        os.makedirs(egg_folder, exist_ok=True)

        egg_path = os.path.join('\\\\?\\', egg_path)
        subprocess.run(
            ["C:\\Program Files\\Bandizip\\Bandizip.exe", "bx", "-y",
                f"-o:{egg_folder}", egg_path],
            check=True
        )

        print(f"{egg_path}을/를 {egg_folder}로 압축 해제했습니다.")

        os.remove(egg_path)
        return [os.path.join(egg_folder, f) for f in os.listdir(egg_folder)], egg_folder
    except Exception as e:
        print(f"압축 해제 중 오류 발생: {e}")
        except_log(egg_path, e)
        return [], ""


def process_folder(folder_path):
    """지정된 폴더를 순회하면서 zip 파일 조각을 결합하고 처리하며 egg 파일도 처리합니다."""
    results = []
    for root, _, files in os.walk(folder_path):
        zip_files = [f for f in files if f.lower().endswith('.zip')]
        egg_files = [f for f in files if f.lower().endswith('.egg')]
        sevenzip_files = [f for f in files if f.lower().endswith('.7z')]
        alz_files = [f for f in files if f.lower().endswith('.alz')]

        for zip_file in zip_files:
            zip_path = os.path.join(root, zip_file)
            file_list, folder_path = extract_egg(zip_path)
            for file in file_list:
                results.append({'File Name': zip_path, 'Extracted File': file})

        for egg_file in egg_files:
            egg_path = os.path.join(root, egg_file)
            file_list, folder_path = extract_egg(egg_path)
            for file in file_list:
                results.append({'File Name': egg_path, 'Extracted File': file})

        for sevenzip_file in sevenzip_files:
            sevenzip_path = os.path.join(root, sevenzip_file)
            file_list, folder_path = extract_egg(sevenzip_path)
            for file in file_list:
                results.append(
                    {'File Name': sevenzip_path, 'Extracted File': file})

        for alz_file in alz_files:
            alz_path = os.path.join(root, alz_file)
            file_list, folder_path = extract_egg(alz_path)
            for file in file_list:
                results.append({'File Name': alz_path, 'Extracted File': file})

    # 결과를 엑셀 파일로 저장
    df = pd.DataFrame(results)
    df.to_excel(os.path.join(
        folder_path, 'extraction_report.xlsx'), index=False)
    print("엑셀 파일이 생성되었습니다: extraction_report.xlsx")


def except_log(dst, e):
    log_dir = os.path.join(folder_path, "log.txt")
    with open(log_dir, 'a') as file:
        file.write(f'Error ({e}) , Source File ({dst})\n')


if __name__ == "__main__":
    folder_path = input("폴더 경로를 입력하세요 : ")
    folder_path = os.path.join("\\\\?\\", folder_path)
    print("진행중...")
    process_folder(folder_path)
    print("모든 압축파일 해제가 완료되었습니다")
