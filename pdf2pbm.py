import os
import fitz


def extract_bookmarks(pdf_path):
    doc = fitz.open(pdf_path)
    toc = doc.get_toc(simple=False)
    return toc


def convert_to_html(bookmarks):
    html_content = '<HTML><HEAD>\n<META NAME="PBM Ver 1.0", CONTENT="Bookmark exported by N2M">\n</HEAD>\n'

    for item in bookmarks:
        level, title, page, _ = item
        blank = "    "*(level-1)
        html_content += f'{blank}<Level ID="{level}", Page="{page -
                                                             1}">{title}</Level>\n'
    html_content += '</HTML>'
    return html_content


def save_pbm(html_content, output_path):
    with open(output_path, 'w', encoding='utf-8') as file:
        file.write(html_content)


def process_files(src_dir, dst_dir):
    try:
        for root, _, files in os.walk(src_dir):
            for file in files:
                file_path = os.path.join(root, file)
                relative_path = os.path.relpath(file_path, src_dir)
                new_file_path = os.path.join(dst_dir, relative_path)

                os.makedirs(os.path.dirname(new_file_path), exist_ok=True)

                if file.lower().endswith('.pdf'):
                    pdf_file_path = new_file_path.rsplit('.', 1)[0] + '.PBM'
                    save_pbm(convert_to_html(
                        extract_bookmarks(file_path)), pdf_file_path)
                else:
                    continue
    except:
        log_missing_file(file, dst_dir)


def log_missing_file(file_name, output_path):
    log_file_path = os.path.join(output_path, 'log.txt')
    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        log_file.write('[PBM 변환 실패] ' + file_name + '\n')


def main():
    pdf_path = input("pdf 파일이 존재하는 폴더를 입력하세요 : ").strip()
    output_path = input("저장할 파일 폴더를 입력하세요 : ").strip()

    if not os.path.isdir(pdf_path):
        print("입력 폴더의 경로를 다시 한번 확인해주세요")
        return
    if not os.path.exists(output_path):
        os.makedirs(output_path)
    process_files(pdf_path, output_path)
    print("모든 작업이 정상적으로 완료되었습니다.")


if __name__ == "__main__":
    main()
