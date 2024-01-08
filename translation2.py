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
    user_input = user_input.strip()

    # 가장 유사한 키 찾기
    match, score = process.extractOne(user_input, translation_data.index)

    # 유사도가 일정 수준 이상인 경우에만 번역
    if score >= 80:
        translated_text = translation_data.loc[match, 'Value']
        return f"번역 결과 ({match} → {user_input}): {translated_text}"
    else:
        return f"해당하는 번역을 찾을 수 없습니다: {user_input}"

if __name__ == "__main__":
    excel_file_path = r"C:\Users\USER\Desktop\새 폴더\translation_data.xlsx"  # 실제 파일 경로에 맞게 수정하세요
    translation_data = load_translation_data(excel_file_path)

    if translation_data is not None:
        while True:
            user_input = input("번역할 한국어 텍스트를 입력하세요 (0을 입력하면 종료): ")

            if user_input == '0':
                print("프로그램을 종료합니다.")
                break

            translated_text = translate_text(user_input, translation_data)
            print(translated_text)
