# app.py
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

import settings
from ocr import extract_text_from_images
from output_handlers import (
    save_output_files, open_file,
    add_text_to_existing_file, preview_and_edit_pdf
)

# ===== 디렉토리 선택 =====
def select_directory():
    folder_path = filedialog.askdirectory(title="스캔 이미지 폴더 선택")
    return folder_path

# ===== 상위 메뉴 GUI =====
def show_file_selection_gui(path_dir):
    root = tk.Tk()
    root.title("원하는 항목 선택")
    root.geometry("300x400")

    def save_excel_file():
        save_output_files(path_dir, file_type="excel")
        open_file(settings.output_file_excel)

    def save_docx_file():
        save_output_files(path_dir, file_type="docx")
        open_file(settings.output_file_docx)

    def save_pdf_file():
        save_output_files(path_dir, file_type="pdf")
        preview_and_edit_pdf()

    tk.Button(root, text="EXCEL 파일 열기 및 저장", command=save_excel_file).pack(pady=10)
    tk.Button(root, text="DOCX 파일 열기 및 저장", command=save_docx_file).pack(pady=10)
    tk.Button(root, text="PDF 파일 미리보기 및 수정", command=save_pdf_file).pack(pady=10)

    tk.Button(
        root, text="기존 EXCEL 파일에 텍스트 추가",
        command=lambda: add_text_to_existing_file("엑셀", pd.DataFrame(settings.data))
    ).pack(pady=10)

    tk.Button(
        root, text="기존 DOCX 파일에 텍스트 추가",
        command=lambda: add_text_to_existing_file("DOCX", pd.DataFrame(settings.data))
    ).pack(pady=10)

    root.mainloop()

# ===== 엔트리 포인트 =====
if __name__ == "__main__":
    # 1) 폴더 선택
    path_dir = select_directory()
    if path_dir:
        # 2) OCR 진행(이미지→텍스트 추출 및 라인 그룹화)
        extract_text_from_images(path_dir)
        # 3) 상위 메뉴 GUI
        show_file_selection_gui(path_dir)
    else:
        messagebox.showerror("경로 오류", "유효한 폴더 경로를 선택해주세요.")
