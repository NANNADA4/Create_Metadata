import os
import shutil


def copy_files_with_extensions(src, dst, extensions):
    """
    특정 확장자를 가진 파일만을 복사합니다.

    :param src: 원본 디렉터리 경로
    :param dst: 대상 디렉터리 경로
    :param extensions: 복사할 파일 확장자의 리스트 (예: ['.zip', '.egg'])
    """
    if not os.path.exists(dst):
        os.makedirs(dst)  # 대상 디렉터리가 없으면 생성합니다.

    for dirpath, _, filenames in os.walk(src):
        # 현재 디렉터리에서 상대 경로를 계산합니다.
        relative_path = os.path.relpath(dirpath, src)
        target_dir = os.path.join(dst, relative_path)

        # 특정 확장자를 가진 파일이 있는 경우만 폴더를 생성합니다.
        files_to_copy = [filename for filename in filenames if any(
            filename.endswith(ext) for ext in extensions)]
        if files_to_copy:
            if not os.path.exists(target_dir):
                os.makedirs(target_dir)

            # 특정 확장자를 가진 파일만 복사합니다.
            try:
                for filename in files_to_copy:
                    src_file = os.path.join(dirpath, filename)
                    dst_file = os.path.join(target_dir, filename)
                    shutil.copy2(src_file, dst_file)
                    print(f"{src_file} - 복사 완료")
            except Exception as e:
                log_dir = os.path.join(dst, "log.txt")
                with open(log_dir, 'a') as file:
                    file.write(f'Error ({e}) , Source File ({
                               src_file}) , Destination File ({dst_file})\n')


# 사용 예제
source_directory = '\\\\?\\F:\\2023년도 국정감사 자료\\05. 위원회 요구 제출 자료_서면질의 답변 자료'
destination_directory = '\\\\?\\D:\\모든압축파일'
extensions_to_copy = ['.zip', '.egg', '.alz',
                      '.z01', '.z02', '.z03', '.z04', '.z05', '.z06']

copy_files_with_extensions(
    source_directory, destination_directory, extensions_to_copy)
