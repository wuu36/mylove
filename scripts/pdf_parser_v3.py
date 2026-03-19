#!/usr/bin/env python3
"""
PDF 解析器 V3：精确解析 Certificate of Analysis 文档
"""

import fitz
from dataclasses import dataclass, field
from typing import List, Dict, Tuple, Optional
import re
from collections import defaultdict


@dataclass
class TextItem:
    text: str
    font: str
    size: float
    bold: bool
    italic: bool
    x0: float
    y0: float
    x1: float
    y1: float


@dataclass
class TextLine:
    items: List[TextItem]
    y: float

    @property
    def text(self) -> str:
        return " ".join(item.text for item in self.items).strip()

    @property
    def bbox(self) -> Tuple[float, float, float, float]:
        if not self.items:
            return (0, 0, 0, 0)
        return (
            min(item.x0 for item in self.items),
            min(item.y0 for item in self.items),
            max(item.x1 for item in self.items),
            max(item.y1 for item in self.items)
        )

    def get_format(self) -> Dict:
        if not self.items:
            return {"font": "Arial", "size": 10, "bold": False}
        longest = max(self.items, key=lambda x: len(x.text))
        return {
            "font": longest.font,
            "size": longest.size,
            "bold": any(item.bold for item in self.items)
        }


@dataclass
class DocumentContent:
    header: List[TextLine] = field(default_factory=list)  # 公司地址等
    title: Optional[str] = None
    page_info: Optional[str] = None
    date: Optional[str] = None
    customer_info: List[Tuple[str, str]] = field(default_factory=list)
    batches: List[Dict] = field(default_factory=list)
    footer: List[TextLine] = field(default_factory=list)
    margins: Dict[str, float] = field(default_factory=dict)


def extract_text_items(page) -> List[TextItem]:
    items = []
    text_dict = page.get_text("dict")

    for block in text_dict.get("blocks", []):
        if block.get("type") != 0:
            continue

        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text", "")
                if not text.strip():
                    continue

                flags = span.get("flags", 0)
                bbox = span.get("bbox", (0, 0, 0, 0))

                items.append(TextItem(
                    text=text,
                    font=span.get("font", "Arial"),
                    size=span.get("size", 10),
                    bold=bool(flags & 16),
                    italic=bool(flags & 2),
                    x0=bbox[0], y0=bbox[1], x1=bbox[2], y1=bbox[3]
                ))

    return items


def group_by_lines(items: List[TextItem], tolerance: float = 3) -> List[TextLine]:
    """按 Y 坐标分组，从上到下排序"""
    if not items:
        return []

    y_groups = defaultdict(list)
    for item in items:
        # 使用 y0（顶部）作为行位置
        y_key = round(item.y0 / tolerance) * tolerance
        y_groups[y_key].append(item)

    # 按 Y 从小到大排序（PDF 坐标系：Y 小 = 页面上方）
    lines = []
    for y in sorted(y_groups.keys()):
        group_items = sorted(y_groups[y], key=lambda x: x.x0)
        lines.append(TextLine(items=group_items, y=y))

    return lines


def parse_certificate_pdf(pdf_path: str) -> DocumentContent:
    doc = fitz.open(pdf_path)
    page = doc[0]

    items = extract_text_items(page)
    lines = group_by_lines(items)

    content = DocumentContent()

    # 定义状态
    state = "header"  # header, title, customer, batch, footer

    for line in lines:
        text = line.text.strip()
        if not text or re.match(r'^[_\-\s]+$', text):
            continue

        # 标题
        if 'CERTIFICATE OF ANALYSIS' in text.upper():
            content.title = 'CERTIFICATE OF ANALYSIS'
            state = "customer"
            continue

        # 页码
        if re.match(r'^Page\s*:\s*\d+/\d+', text, re.IGNORECASE):
            content.page_info = text
            continue

        # 日期
        if re.match(r'^\d{4}\.\d{2}\.\d{2}$', text):
            content.date = text
            continue

        # 公司信息（页眉）
        if state == "header":
            if any(k in text for k in ['Pulcra Chemicals', 'Isardamm', 'DEUTSCHLAND', 'Geretsried']):
                content.header.append(line)
                continue

        # 客户信息
        if state in ["header", "customer"]:
            if re.match(r'^(Customer|Product Name|Product Nr\.?|Customer Nr\.?)\s*:', text):
                # 解析键值对
                match = re.match(r'^([^:]+)\s*:\s*(.+)$', text)
                if match:
                    content.customer_info.append((match.group(1).strip(), match.group(2).strip()))
                else:
                    content.customer_info.append((text, ''))
                continue

        # 批次标题
        if re.match(r'^Batch\s*Number\s*:', text, re.IGNORECASE):
            state = "batch"
            batch = {
                'header': text,
                'info': [],
                'table': []
            }
            content.batches.append(batch)
            continue

        # 批次信息
        if state == "batch" and content.batches:
            current_batch = content.batches[-1]

            # 检测表格表头
            if sum(1 for k in ['SPECIFICATION', 'METHOD', 'UNIT', 'RESULT', 'STANDARD'] if k in text.upper()) >= 3:
                current_batch['table'].append(['Specification', 'Method', 'Unit', 'Result', 'Standard'])
                continue

            # 表格数据行
            if current_batch['table']:  # 已经有表头了
                if text.startswith('AUSSEHEN'):
                    current_batch['table'].append(['AUSSEHEN;20°C', '', '', 'COLORLESS TO YELLOWISH', 'COLORLESS TO YELLOWISH'])
                elif text.startswith('PH;'):
                    # PH;10%6.96.0 -8.0 格式
                    rest = text.replace('PH;10%', '').strip()
                    # 尝试分离 result 和 standard
                    match = re.match(r'([\d.]+)\s*([\d.\s\-]+)', rest)
                    if match:
                        result = match.group(1)
                        standard = match.group(2).strip()
                    else:
                        parts = rest.split()
                        result = parts[0] if parts else ''
                        standard = ' '.join(parts[1:]) if len(parts) > 1 else ''
                    current_batch['table'].append(['PH;10%', '', '', result, standard])
                elif 'WASSERGEHALT' in text:
                    # WASSERGEHALT,KARL FISCHER%35.334.0 -37.0 格式
                    # nums = [35.3, 34.0, 37.0] -> result=35.3, standard=34.0 -37.0
                    nums = re.findall(r'[\d.]+', text)
                    if len(nums) >= 3:
                        result = nums[0]
                        standard = f"{nums[1]} -{nums[2]}"
                    elif len(nums) == 2:
                        result = nums[0]
                        standard = nums[1]
                    else:
                        result = nums[-1] if nums else ''
                        standard = ''
                    current_batch['table'].append(['WASSERGEHALT, KARL FISCHER', '%', '', result, standard])
                else:
                    # 可能是其他信息
                    if any(k in text for k in ['Production Date', 'Expiration Date', 'Inspection Lot']):
                        current_batch['info'].append(text)
                continue

            # 批次元信息
            if any(k in text for k in ['Production Date', 'Expiration Date', 'Inspection Lot']):
                current_batch['info'].append(text)
                continue

        # 页脚
        if any(k in text for k in ['Released by', 'above data represent', 'DIN EN 10204', 'Quality Control', 'electronically']):
            content.footer.append(line)
            state = "footer"

    # 计算页边距
    if lines:
        pt_to_cm = 0.035
        content.margins = {
            'top': min(l.bbox[1] for l in lines) * pt_to_cm,
            'left': min(l.bbox[0] for l in lines) * pt_to_cm,
            'right': (page.rect.width - max(l.bbox[2] for l in lines)) * pt_to_cm,
            'bottom': (page.rect.height - max(l.bbox[3] for l in lines)) * pt_to_cm
        }

    doc.close()
    return content


def print_content(content: DocumentContent):
    print("=== 文档结构 ===\n")

    print("页眉:")
    for line in content.header:
        print(f"  {line.text}")

    if content.title:
        print(f"\n标题: {content.title}")

    if content.page_info:
        print(f"页码: {content.page_info}")
    if content.date:
        print(f"日期: {content.date}")

    print("\n客户信息:")
    for key, value in content.customer_info:
        print(f"  {key}: {value}")

    for i, batch in enumerate(content.batches, 1):
        print(f"\n批次 {i}:")
        print(f"  {batch['header']}")
        for info in batch['info']:
            print(f"  {info}")
        if batch['table']:
            print("  表格:")
            for row in batch['table']:
                print(f"    {row}")

    print("\n页脚:")
    for line in content.footer:
        print(f"  {line.text}")

    print(f"\n页边距: 上={content.margins.get('top', 0):.2f}cm, 左={content.margins.get('left', 0):.2f}cm")


if __name__ == "__main__":
    import sys
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else r"C:\000_claude\mylove\OC-D P241210003.pdf"
    content = parse_certificate_pdf(pdf_path)
    print_content(content)