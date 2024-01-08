import re
from fuzzywuzzy import process
import pandas as pd

def load_translation_data(file_path):
    try:
        translation_data = pd.read_excel(file_path, index_col='Key')
        return translation_data
    except FileNotFoundError:
        print(f"파일을 찾을 수 없습니다: {file_path}")
        return None
    except pd.errors.EmptyDataError:
        print(f"빈 파일입니다: {file_path}")
        return None
    except Exception as e:
        print(f"오류 발생: {e}")
        return None

def translate_text(user_input, translation_data):
    translations = []

    # 각 줄마다 번역 수행
    for line in user_input.split('\n'):
        line = line.strip()
        if line:  # 빈 줄은 무시
            match, score = process.extractOne(line, translation_data.index)
            if score >= 80:
                translated_line = translation_data.loc[match, 'Value']
                translations.append(f"번역 결과 ({match} → {line}): {translated_line}")
            else:
                translations.append(f"해당하는 번역을 찾을 수 없습니다: {line}")

    return '\n'.join(translations)

def extract_korean_text(input_text):
    # 한글 문자열 추출을 위한 정규 표현식
    pattern = r'[가-힣]+'
    return re.findall(pattern, input_text)

if __name__ == "__main__":
    excel_file_path = r"C:\Users\USER\Desktop\translation\translation_data.xlsx"
    translation_data = load_translation_data(excel_file_path)

    if translation_data is not None:
        print("여러 줄의 텍스트를 입력하세요 (입력을 마치려면 '종료'를 입력하세요):")
        while True:
            user_input = []
            while True:
                line = input()
                if line == '종료':
                    break
                user_input.append(line)
            user_input = '\n'.join(user_input)

            if user_input.strip() == '0':
                print("프로그램을 종료합니다.")
                break

            for line in user_input.split('\n'):
                korean_texts = extract_korean_text(line)
                for text in korean_texts:
                    translated_text = translate_text(text, translation_data)
                    print(translated_text)