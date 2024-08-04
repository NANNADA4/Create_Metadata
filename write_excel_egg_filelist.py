import os
import mmap
import zlib
import bz2
import struct
import openpyxl

# EGG 압축 파일 관련 상수
SIZE_EGG_HEADER = 14

COMPRESS_METHOD_STORE = 0
COMPRESS_METHOD_DEFLATE = 1
COMPRESS_METHOD_BZIP2 = 2


class EggFile:
    def __init__(self, filename):
        self.fp = None
        self.mm = None
        self.data_size = 0
        self.egg_pos = None

        try:
            self.data_size = os.path.getsize(filename)
            self.fp = open(filename, 'rb')
            self.mm = mmap.mmap(self.fp.fileno(), 0, access=mmap.ACCESS_READ)
        except IOError:
            print(f"Error opening file {filename}")
            raise

    def close(self):
        if self.mm:
            self.mm.close()
        if self.fp:
            self.fp.close()

    def read(self, filename):
        ret_data = None

        try:
            fname = self.__FindFirstFileName__()
            while fname:
                if fname == filename:
                    data, method, self.egg_pos = self.__ReadBlockData__()
                    if method == COMPRESS_METHOD_STORE:
                        ret_data = data
                        break
                    elif method == COMPRESS_METHOD_DEFLATE:
                        ret_data = zlib.decompress(data, -15)
                        break
                    elif method == COMPRESS_METHOD_BZIP2:
                        ret_data = bz2.decompress(data)
                        break
                    else:
                        pass
                fname = self.__FindNextFileName__()
        except:
            print("Error reading file from archive")
            raise

        return ret_data

    def namelist(self):
        name_list = []

        try:
            fname = self.__FindFirstFileName__()
            while fname:
                name_list.append(fname)
                fname = self.__FindNextFileName__()
        except:
            print("Error listing files from archive")
            raise

        return name_list

    def __FindFirstFileName__(self):
        self.egg_pos = 0
        fname, self.egg_pos = self.__GetFileName__(self.egg_pos)
        return fname

    def __FindNextFileName__(self):
        fname, self.egg_pos = self.__GetFileName__(self.egg_pos)
        return fname

    def __GetFileName__(self, egg_pos):
        mm = self.mm
        data_size = self.data_size

        try:
            while egg_pos < data_size:
                magic = struct.unpack('<I', mm[egg_pos:egg_pos+4])[0]
                if magic == 0x0A8591AC:  # Filename Header
                    size, fname = self.__EGG_Filename_Header__(mm[egg_pos:])
                    if size == -1:
                        raise SystemError
                    egg_pos += size
                    return fname, egg_pos
                else:
                    egg_pos = self.__DefaultMagicIDProc__(magic, egg_pos)
                    if egg_pos == -1:
                        raise SystemError
        except SystemError:
            print("Error finding filename header")
            raise

        return None, -1

    def __ReadBlockData__(self):
        egg_pos = self.egg_pos
        mm = self.mm
        data_size = self.data_size

        try:
            while egg_pos < data_size:
                magic = struct.unpack('<I', mm[egg_pos:egg_pos+4])[0]
                if magic == 0x02B50C13:  # Block Header
                    size = self.__EGG_Block_Header_Size__(mm[egg_pos:])
                    if size == -1:
                        raise SystemError

                    compress_method_m = mm[egg_pos+4]
                    compress_size = struct.unpack(
                        '<I', mm[egg_pos+10:egg_pos+14])[0]
                    compressed_data = mm[egg_pos+18:egg_pos+18+compress_size]
                    egg_pos += size
                    return compressed_data, compress_method_m, egg_pos
                else:
                    egg_pos = self.__DefaultMagicIDProc__(magic, egg_pos)
                    if egg_pos == -1:
                        raise SystemError
        except SystemError:
            print("Error reading block data")
            raise

        return None, -1, -1

    def __DefaultMagicIDProc__(self, magic, egg_pos):
        mm = self.mm
        data_size = self.data_size

        try:
            if egg_pos < data_size:
                if magic == 0x41474745:  # EGG Header
                    if self.__EGG_Header__(mm) == -1:
                        raise SystemError
                    egg_pos += SIZE_EGG_HEADER
                elif magic == 0x0A8590E3:  # File Header
                    egg_pos += 16
                elif magic == 0x02B50C13:  # Block Header
                    size = self.__EGG_Block_Header_Size__(mm[egg_pos:])
                    if size == -1:
                        raise SystemError
                    egg_pos += size
                elif magic == 0x08D1470F:  # Encrypt Header
                    size = self.__EGG_Encrypt_Header_Size__(mm[egg_pos:])
                    if size == -1:
                        raise SystemError
                    egg_pos += size
                elif magic == 0x2C86950B:  # Windows File Information
                    egg_pos += 16
                elif magic == 0x1EE922E5:  # Posix File Information
                    egg_pos += 27
                elif magic == 0x07463307:  # Dummy Header
                    size = self.__EGG_Dummy_Header_Size__(mm[egg_pos:])
                    if size == -1:
                        raise SystemError
                    egg_pos += size
                elif magic == 0x0A8591AC:  # Filename Header
                    size, fname = self.__EGG_Filename_Header__(mm[egg_pos:])
                    if size == -1:
                        raise SystemError
                    egg_pos += size
                elif magic == 0x04C63672:  # Comment Header
                    raise SystemError  # Not supported
                elif magic == 0x24F5A262:  # Split Compression
                    egg_pos += 15
                elif magic == 0x24E5A060:  # Solid Compression
                    egg_pos += 7
                elif magic == 0x08E28222:  # End of File Header
                    egg_pos += 4
                else:
                    raise SystemError
        except SystemError:
            print("Error processing magic ID")
            return -1

        return egg_pos

    def __EGG_Header__(self, data):
        try:
            magic = struct.unpack('<I', data[0:4])[0]
            if magic != 0x41474745:
                raise SystemError

            version = struct.unpack('<H', data[4:6])[0]
            if version != 0x0100:
                raise SystemError

            header_id = struct.unpack('<I', data[6:10])[0]
            if header_id == 0:
                raise SystemError

            reserved = struct.unpack('<I', data[10:14])[0]
            if reserved != 0:
                raise SystemError

            return 0
        except SystemError:
            print("Error reading EGG header")
            return -1

    def __EGG_Encrypt_Header_Size__(self, data):
        try:
            encrypt_method = data[7]
            if encrypt_method == 0:
                return 24
            elif encrypt_method == 1:
                return 28
            elif encrypt_method == 2:
                return 36
            else:
                raise SystemError
        except SystemError:
            print("Error calculating encrypt header size")
            return -1

    def __EGG_Dummy_Header_Size__(self, data):
        try:
            dummy_size = struct.unpack('<H', data[5:7])[0]
            return 7 + dummy_size
        except:
            print("Error calculating dummy header size")
            return -1

    def __EGG_Filename_Header__(self, data):
        size = -1
        fname = None

        try:
            fname_size = struct.unpack('<H', data[5:7])[0]
            fname = data[7:7+fname_size]
            size = 7 + fname_size
        except:
            print("Error reading filename header")
            pass

        return size, fname.decode('utf-8')

    def __EGG_Block_Header_Size__(self, data):
        try:
            block_size = 18 + 4
            compress_size = struct.unpack('<I', data[10:14])[0]
            size = block_size + compress_size
            return size
        except:
            print("Error calculating block header size")
            return -1


def get_all_files(egg_file):
    file_paths = []

    try:
        namelist = egg_file.namelist()
        for name in namelist:
            # 전체 경로를 저장
            file_paths.append(name)
    except Exception as e:
        print(f"An error occurred: {e}")

    return file_paths


def save_to_excel(file_paths, egg_filepath):
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # 헤더 작성
    sheet.append(["파일 경로", "EGG 파일"])

    egg_filename = os.path.basename(egg_filepath)
    egg_path_no_extension = egg_file_path.replace('.egg', '')

    # 파일 경로와 EGG 파일명 작성 (확장자 포함)
    for path in file_paths:
        full_path = os.path.join(egg_path_no_extension, path)
        sheet.append([full_path, egg_filename])

    # 엑셀 파일 저장
    excel_filename = f"{egg_filename}.xlsx"
    workbook.save(excel_filename)
    print(f"파일 목록이 {excel_filename}에 저장되었습니다.")


# 사용 예시
egg_file_path = '/Users/nannada4/Downloads/test.egg'  # EGG 파일 경로를 여기에 설정하세요
egg_file = EggFile(egg_file_path)

# 모든 파일 경로를 리스트로 가져오기
file_paths = get_all_files(egg_file)

# 파일 경로를 엑셀 파일로 저장 (EGG 파일명에 확장자 포함)
save_to_excel(file_paths, os.path.basename(egg_file_path))

egg_file.close()
