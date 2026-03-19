#!/usr/bin/env python3
"""
PDF 布局精确分析器
提取每个元素的位置信息用于精确复制
"""

import fitz
from dataclasses import dataclass
from typing import List, Dict, Tuple
import json


@dataclass
class TextBlock:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    font: str
    size: float
    bold: bool
    page_width: float
    page_height: float


def analyze_pdf_layout(pdf_path: str):
    """精确分析 PDF 布局"""
    doc = fitz.open(pdf_path)
    page = doc[0]

    rect = page.rect
    page_width = rect.width
    page_height = rect.height

    print(f"页面尺寸: {page_width:.1f} x {page_height:.1f} pt")
    print(f"         ({page_width/72:.2f} x {page_height/72:.2f} inches)")
    print()

    # 获取所有文本块
    blocks = []

    text_dict = page.get_text("dict")

    for block in text_dict.get("blocks", []):
        if block.get("type") != 0:
            continue

        bbox = block.get("bbox", (0, 0, 0, 0))

        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text", "").strip()
                if not text:
                    continue

                span_bbox = span.get("bbox", bbox)
                flags = span.get("flags", 0)

                block = TextBlock(
                    text=text,
                    x0=span_bbox[0],
                    y0=span_bbox[1],
                    x1=span_bbox[2],
                    y1=span_bbox[3],
                    font=span.get("font", "Arial"),
                    size=span.get("size", 10),
                    bold=bool(flags & 16),
                    page_width=page_width,
                    page_height=page_height
                )
                blocks.append(block)

    doc.close()

    # 按 Y 坐标排序（从上到下）
    blocks.sort(key=lambda b: (b.y0, b.x0))

    # 分析布局
    print("=" * 70)
    print("PDF 布局分析 (按位置排序)")
    print("=" * 70)
    print()
    print(f"{'Y位置':>8} {'X位置':>8} {'宽度':>8} {'字号':>6} {'粗体':>4} 文本")
    print("-" * 70)

    for b in blocks:
        # 计算相对位置（百分比）
        y_pct = b.y0 / page_height * 100
        x_pct = b.x0 / page_width * 100
        width = b.x1 - b.x0

        # 判断对齐
        if abs((b.x0 + b.x1) / 2 - page_width / 2) < page_width * 0.1:
            align = "center"
        elif b.x1 > page_width * 0.85:
            align = "right"
        else:
            align = "left"

        text_preview = b.text[:40] + "..." if len(b.text) > 40 else b.text

        print(f"{b.y0:>8.1f} {b.x0:>8.1f} {width:>8.1f} {b.size:>6.1f} {b.bold!s:>4} {text_preview}")

    return blocks


def identify_regions(blocks: List[TextBlock], page_width: float, page_height: float):
    """识别文档区域"""
    print("\n" + "=" * 70)
    print("区域识别")
    print("=" * 70)

    # 按 Y 坐标分组
    y_groups = {}
    tolerance = 5  # 5pt 容差

    for b in blocks:
        y_key = round(b.y0 / tolerance) * tolerance
        if y_key not in y_groups:
            y_groups[y_key] = []
        y_groups[y_key].append(b)

    # 识别区域
    regions = {
        'header': [],      # 页眉（右上角）
        'title': [],       # 标题
        'customer': [],    # 客户信息
        'batch1': [],      # 批次1
        'batch2': [],      # 批次2
        'footer': []       # 页脚
    }

    for y in sorted(y_groups.keys()):
        group = y_groups[y]
        texts = [b.text for b in group]
        combined = ' '.join(texts)

        # 页眉区域（顶部，右对齐）
        if any(k in combined for k in ['Pulcra', 'Isardamm', 'DEUTSCHLAND', 'Geretsried']):
            regions['header'].extend(group)
        # 标题
        elif 'CERTIFICATE OF ANALYSIS' in combined:
            regions['title'].extend(group)
        # 客户信息
        elif any(k in combined for k in ['Customer', 'Product Name', 'Product Nr', 'Customer Nr']):
            regions['customer'].extend(group)
        # 批次
        elif 'Batch Number' in combined:
            regions['batch1'].extend(group) if 'P241200117' in combined else regions['batch2'].extend(group)
        # 表格数据
        elif any(k in combined for k in ['Specification', 'AUSSEHEN', 'PH;', 'WASSERGEHALT']):
            # 根据位置分配到批次
            if y < page_height * 0.6:
                regions['batch1'].extend(group)
            else:
                regions['batch2'].extend(group)
        # 页脚
        elif any(k in combined for k in ['Released by', 'DIN EN', 'Quality Control', 'electronically']):
            regions['footer'].extend(group)

    for name, items in regions.items():
        if items:
            print(f"\n[{name.upper()}]")
            for b in sorted(items, key=lambda x: (x.y0, x.x0)):
                print(f"  Y={b.y0:.0f} X={b.x0:.0f}: {b.text[:50]}")

    return regions


if __name__ == "__main__":
    import sys
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else r"C:\000_claude\mylove\OC-D P241210003.pdf"

    blocks = analyze_pdf_layout(pdf_path)
    identify_regions(blocks, blocks[0].page_width if blocks else 595, blocks[0].page_height if blocks else 842)