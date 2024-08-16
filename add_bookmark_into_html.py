import os
import fitz
import subprocess


def extract_bookmarks(pdf_path):
    doc = fitz.open(pdf_path)
    toc = doc.get_toc(simple=False)
    return toc


# def pdf2html()
# wsl ~/pdf2htmlEX-0.18.8.rc1-master-20200630-Ubuntu-focal-x86_64.AppImage 커맨드 사용
def pdf2html(pdf_path, html_path, pdf_alphabet, html_alphabet):
    pdf_path = pdf_path.replace(f'{pdf_alphabet}:\\', '')
    html_path = html_path.replace(f'{html_alphabet}:\\', '')
    pdf_path = pdf_path.replace('\\', '/')
    html_path = html_path.replace('\\', '/')
    pdf_alphabet = pdf_alphabet.lower()
    html_alphabet = html_alphabet.lower()
    html_result = subprocess.run(
        ['wsl', '~/pdf2htmlEX-0.18.8.rc1-master-20200630-Ubuntu-focal-x86_64.AppImage', '--process-outline', '0', f'/mnt/{pdf_alphabet}/'+pdf_path, '--dest-dir', f'/mnt/{html_alphabet}/'+html_path])
    print(html_result.stdout)
    if (html_result.returncode == 0):
        print(f"[변환 완료] {pdf_path}")
    else:
        print(f"[변환 실패] {pdf_path}")


def convert_to_html(bookmarks):
    html_content = ''

    for item in bookmarks:
        level, title, page, _ = item
        blank = "    "*(level-1)
        html_content += f'{blank}<Level ID="{level}", Page="{page -
                                                             1}">{title}</Level>\n'
    return html_content


def change_html(html_content, html_path):
    with open(html_path, 'r', encoding='utf-8') as file:
        existing_file = file.readlines()
    head_index = None
    for i, line in enumerate(existing_file):
        if '</head>' in line.lower():
            head_index = i
            break

    if head_index is not None:
        updated_file = existing_file[:head_index + 1] + \
            [html_content] + existing_file[head_index + 1:]
    else:
        updated_file = existing_file + [html_content]

    with open(html_path, 'w', encoding='utf-8') as file:
        file.writelines(updated_file)


def save_html(html_content, output_path, pdf_name):
    base_name = os.path.splitext(pdf_name)[0]
    html_path = os.path.join(output_path, base_name + '.html')

    if os.path.exists(html_path):
        change_html(html_content, html_path)
    else:
        with open(html_path, 'w', encoding='utf-8') as file:
            file.write(html_content)
        log_missing_file(base_name, output_path)


def log_missing_file(file_name, output_path):
    log_file_path = os.path.join(output_path, 'log.txt')
    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        log_file.write(file_name + '\n')


def process_files(src_dir, dst_dir, src_alphabet, dst_alphabet):
    for root, _, files in os.walk(src_dir):
        for file in files:
            if file.lower().endswith('.pdf'):
                file_path = os.path.join(root, file)
                pdf2html(file_path, dst_dir, src_alphabet, dst_alphabet)
                relative_path = os.path.relpath(file_path, src_dir)
                new_file_path = os.path.join(dst_dir, relative_path)

                os.makedirs(os.path.dirname(new_file_path), exist_ok=True)

                pdf_name = os.path.basename(file_path)
                save_html(convert_to_html(
                    extract_bookmarks(file_path)), dst_dir, pdf_name)


def main():
    pdf_path = input("pdf 파일이 존재하는 폴더 경로를 입력하세요 : ").strip()
    pdf_alphabet = pdf_path[0]
    output_path = input("저장할 파일 경로를 입력하세요 : ").strip()
    html_alphabet = output_path[0]

    if not os.path.isdir(pdf_path):
        print("입력 폴더의 경로를 다시 한번 확인해주세요")
        return
    if not os.path.exists(output_path):
        os.makedirs(output_path)
    process_files(pdf_path, output_path, pdf_alphabet, html_alphabet)
    print("모든 작업이 정상적으로 완료되었습니다.")


if __name__ == "__main__":
    main()
