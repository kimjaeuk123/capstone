# ocr.py
import os
import numpy as np
from PIL import Image
import settings
from preprocess import resize_image, enhance_image_for_ocr

def group_text_by_lines(text_data, threshold=10):
    lines = []
    current_line = []
    current_y = None

    for result in text_data:
        bbox, text, confidence = result[0], result[1], result[2]
        y = bbox[0][1]
        if current_y is None:
            current_y = y
            current_line.append(text)
        else:
            if abs(y - current_y) < threshold:
                current_line.append(text)
            else:
                lines.append(" ".join(current_line))
                current_line = [text]
                current_y = y
    if current_line:
        lines.append(" ".join(current_line))
    return lines

def extract_text_from_images(path_dir):
    file_list = os.listdir(path_dir)
    for file_name in file_list:
        if file_name.lower() in ["output.xlsx", "output.docx", "output.pdf"]:
            continue
        file_path = os.path.join(path_dir, file_name)
        try:
            img = Image.open(file_path)
            img = resize_image(img)
            img = enhance_image_for_ocr(img)
            img_cv = np.array(img)

            text_data = settings.reader.readtext(img_cv, detail=1)
            lines = group_text_by_lines(text_data)

            settings.data.append({'내용': "\n".join(lines), '이미지': file_path})
        except IOError:
            print(f"파일을 열 수 없습니다: {file_path}")
