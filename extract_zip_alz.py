import os
import zipfile
import subprocess
import shutil


def combine_zip_parts(folder_path, base_name):
    """분할된 zip 파일 조각들을 하나로 결합합니다."""
    combined_zip_path = os.path.join(folder_path, base_name + '.zip')

    try:
        with open(combined_zip_path, 'wb') as combined_zip:
            for i in range(1, 10):
                part_file = f"{base_name}.z{i:02d}"
                part_path = os.path.join(folder_path, part_file)
                if not os.path.exists(part_path):
                    break
                with open(part_path, 'rb') as part:
                    combined_zip.write(part.read())
    except Exception as e:
        except_log(folder_path, e)

    print(f"분할된 zip 파일 조각들을 {combined_zip_path}로 결합했습니다.")
    return combined_zip_path


def extract_zip(zip_path, extract_to_folder):
    """주어진 zip 파일을 지정된 폴더로 압축 해제합니다."""
    try:
        # 압축 해제할 폴더 생성
        zip_folder = os.path.join(
            extract_to_folder, os.path.splitext(os.path.basename(zip_path))[0])
        os.makedirs(zip_folder, exist_ok=True)

        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(zip_folder)

        print(f"{zip_path}을/를 {zip_folder}로 압축 해제했습니다.")

        # 압축 해제 후 zip 파일 삭제
        os.remove(zip_path)

        # 압축 해제 후 빈 폴더 삭제
        if not os.listdir(zip_folder):  # 폴더가 비어있으면
            os.rmdir(zip_folder)
    except Exception as e:
        print(f"zip 압축해제 오류 : {e}")
        except_log(zip_path, e)


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

        # 압축 해제 후 egg 파일 삭제
        os.remove(egg_path)
    except Exception as e:
        print(f"압축 해제 중 오류 발생: {e}")
        except_log(egg_path, e)


def process_folder(folder_path):
    """지정된 폴더를 순회하면서 zip 파일 조각을 결합하고 처리하며 egg 파일도 처리합니다."""
    for root, _, files in os.walk(folder_path):
        zip_parts = [f for f in files if f.endswith('.z01')]
        zip_files = [f for f in files if f.endswith('.zip')]
        egg_files = [f for f in files if f.endswith('.egg')]

        if zip_parts:
            base_name = zip_parts[0].split('.z01')[0]
            combined_zip_path = combine_zip_parts(root, base_name)
            extract_zip(combined_zip_path, root)

        for zip_file in zip_files:
            zip_path = os.path.join(root, zip_file)
            extract_zip(zip_path, root)

        for egg_file in egg_files:
            egg_path = os.path.join(root, egg_file)
            extract_egg(egg_path)


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
