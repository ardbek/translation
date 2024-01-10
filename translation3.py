import openpyxl
import pyperclip

def translate_text(user_input, excel_file_path):
    workbook = openpyxl.load_workbook(excel_file_path)
    sheet = workbook.active
    translation_dict = {row[0].value: row[1].value for row in sheet.iter_rows(min_row=1, max_col=2) if row[0].value is not None and row[1].value is not None}

    lines = user_input.split('\n')
    translated_lines = []
    for line in lines:
        for key, value in translation_dict.items():
            if key is not None and value is not None:
                line = line.replace(key, value)
        translated_lines.append(line)

    return '\n'.join(translated_lines)


# def get_user_input():
#     print("\n번역할 텍스트를 입력하세요 (입력이 끝나면 Ctrl+D (Unix) 또는 Ctrl+Z 후 Enter (Windows)를 누르세요, 종료하려면 '종료' 입력):")
#     user_input_lines = []
#     try:
#         while True:
#             line = input()
#             if line.strip().lower() == '종료':
#                 return None
#             user_input_lines.append(line)
#     except EOFError:
#         pass
#     return '\n'.join(user_input_lines)

def get_user_input():
    print("\n번역할 텍스트를 입력하세요 (입력을 완료하려면 빈 줄에서 Enter를 누르세요, 프로그램을 종료하려면 '종료'를 입력하세요):")
    user_input_lines = []
    while True:
        line = input()
        if not line:
            break
        if line.strip().lower() == '종료':
            return None
        user_input_lines.append(line)
    return '\n'.join(user_input_lines)



def main():
    excel_file_path = r"C:\Users\USER\Desktop\translation\translation_data.xlsx"

    while True:
        user_input = get_user_input()
        if user_input is None:
            break

        translated_text = translate_text(user_input, excel_file_path)
        pyperclip.copy(translated_text)
        print("번역된 텍스트가 클립보드에 복사되었습니다.")

if __name__ == "__main__":
    main()
