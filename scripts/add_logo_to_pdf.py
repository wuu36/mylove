#!/usr/bin/env python3
"""
在PDF上添加Logo图片
用于在不影响Word布局的情况下添加Logo
"""

import fitz  # PyMuPDF
import os


def add_logo_to_pdf(input_pdf: str, output_pdf: str, logo_path: str,
                    x_pt: float = 28.3, y_pt: float = 56.8,
                    width_pt: float = 138, height_pt: float = 30):
    """
    在PDF指定位置添加Logo图片

    Args:
        input_pdf: 输入PDF路径
        output_pdf: 输出PDF路径
        logo_path: Logo图片路径
        x_pt: X位置（点）
        y_pt: Y位置（点）
        width_pt: 图片宽度（点）
        height_pt: 图片高度（点）
    """
    # 打开PDF
    doc = fitz.open(input_pdf)

    # 在每一页添加Logo
    for page_num in range(len(doc)):
        page = doc[page_num]

        # 创建图片矩形
        rect = fitz.Rect(x_pt, y_pt, x_pt + width_pt, y_pt + height_pt)

        # 插入图片
        page.insert_image(rect, filename=logo_path)

    # 保存
    doc.save(output_pdf)
    doc.close()
    print(f"Logo已添加，保存到: {output_pdf}")


if __name__ == "__main__":
    input_pdf = r"C:\000_claude\mylove\OC-D P241210003_v45.pdf"
    output_pdf = r"C:\000_claude\mylove\OC-D P241210003_v45_with_logo.pdf"
    logo_path = r"C:\000_claude\mylove\logo_correct.png"

    add_logo_to_pdf(input_pdf, output_pdf, logo_path)