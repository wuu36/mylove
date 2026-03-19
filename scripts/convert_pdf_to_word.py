#!/usr/bin/env python3
"""使用 pdf2docx 将 PDF 精确转换为 Word"""

from pdf2docx import Converter

def convert_pdf_to_docx(pdf_path: str, docx_path: str):
    """转换 PDF 到 Word"""
    cv = Converter(pdf_path)
    cv.convert(docx_path)
    cv.close()
    print(f"转换完成: {docx_path}")

if __name__ == "__main__":
    convert_pdf_to_docx(
        "OC-D P241210003.pdf",
        "OC-D P241210003_v8.docx"
    )