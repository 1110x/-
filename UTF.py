import os

# 변환할 디렉토리 경로
directory = r"C:\CENTER"

# 변환할 확장자 목록
extensions = ('.bas', '.frm', '.ftx')

# 디렉토리 내 모든 파일 확인
for filename in os.listdir(directory):
    if filename.endswith(extensions):  # 지정한 확장자만 변환
        file_path = os.path.join(directory, filename)

        # 파일 읽기
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
            content = file.read()

        # 파일 다시 쓰기 (UTF-8로 저장)
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(content)

print("모든 .bas, .frm, .ftx 파일이 UTF-8로 변환되었습니다.")