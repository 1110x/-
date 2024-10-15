import os
import win32com.client

# 엑셀 파일과 내보낼 경로 설정
excel_file = r'C:\Center\2024-센터청하S3.xlsm'
export_folder = r'C:\Center'

# 폴더가 없으면 생성
if not os.path.exists(export_folder):
    os.makedirs(export_folder)

# 엑셀 애플리케이션 객체 생성
excel_app = win32com.client.Dispatch("Excel.Application")
excel_app.Visible = False

try:
    # 엑셀 파일 열기 (매크로 허용)
    workbook = excel_app.Workbooks.Open(excel_file, ReadOnly=False, Editable=True)

    # VBProject 접근
    vbproject = workbook.VBProject

    # 모듈과 유저폼 내보내기
    for component in vbproject.VBComponents:
        if component.Type in [1, 2, 3]:  # 1: 모듈, 2: 클래스 모듈, 3: 유저폼
            ext = 'bas' if component.Type == 1 else 'frm'
            file_name = os.path.join(export_folder, f"{component.Name}.{ext}")
            component.Export(file_name)
            print(f"Exported: {file_name}")
            
            # 내보낸 파일을 UTF-8로 변환
            with open(file_name, 'r', encoding='ansi') as f:
                content = f.read()
            with open(file_name, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"Converted to UTF-8: {file_name}")

except AttributeError:
    print("엑셀 파일에서 VBProject에 접근할 수 없습니다. 'Trust Access to the VBA project object model' 옵션을 확인하세요.")
finally:
    # 엑셀 파일 닫기
    workbook.Close(SaveChanges=False)
    excel_app.Quit()

print("모든 모듈과 유저폼이 내보내졌고, UTF-8로 변환되었습니다.")
