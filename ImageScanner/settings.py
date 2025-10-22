# settings.py
import os
import sys
import easyocr

# ===== 전역 상태 =====
data = []  # [{'내용': "...", '이미지': "path"}] 누적
output_file_excel = ""
output_file_docx = ""
output_file_pdf = ""
font_size_map = []  # (start, end, size) 튜플 리스트

# ===== 폰트 경로 =====
def get_font_path():
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS  # PyInstaller 등 패키징 시
    else:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, 'NanumGothic.ttf')

font_path = get_font_path()

# ===== EasyOCR 리더 =====
# 한/영 동시 인식
reader = easyocr.Reader(['ko', 'en'])
