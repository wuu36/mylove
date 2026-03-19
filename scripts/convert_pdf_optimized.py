#!/usr/bin/env python3
"""PDF 转 Word 优化转换

解决 pdf2docx 转换的问题：
1. 页面溢出：PDF 是 1 页，Word 变成 2 页
2. 对齐问题：文本对齐与原 PDF 不一致
"""

from pdf2docx import Converter
from docx import Document
from docx.shared import Cm, Pt


def convert_with_optimization(pdf_path: str, docx_path: str):
    """优化转换 PDF 到 Word

    两阶段处理：
    1. 使用 pdf2docx 提取内容和结构（带优化参数）
    2. 使用 python-docx 精细调整格式
    """

    # 第一阶段：使用 pdf2docx 转换（带优化参数）
    print(f"正在转换: {pdf_path} -> {docx_path}")

    cv = Converter(pdf_path)
    cv.convert(
        docx_path,
        page_margin_factor_top=0.3,      # 减少顶部边距因子
        page_margin_factor_bottom=0.3,   # 减少底部边距因子
        max_line_spacing_ratio=1.2,      # 减少行间距
        line_break_free_space_ratio=0.05 # 调整换行阈值
    )
    cv.close()
    print("pdf2docx 转换完成")

    # 第二阶段：后处理优化
    doc = Document(docx_path)

    # 调整页面设置 - A4 纸张，标准边距
    for section in doc.sections:
        section.top_margin = Cm(1.0)
        section.bottom_margin = Cm(1.0)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)

    # 调整段落格式
    for para in doc.paragraphs:
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(3)
        para.paragraph_format.line_spacing = 1.15

    doc.save(docx_path)
    print(f"优化转换完成: {docx_path}")


if __name__ == "__main__":
    convert_with_optimization(
        "OC-D P241210003.pdf",
        "OC-D P241210003_v8_optimized.docx"
    )