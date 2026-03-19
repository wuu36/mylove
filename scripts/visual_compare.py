#!/usr/bin/env python3
"""PDF 与 Word 视觉对比工具"""

import fitz  # PyMuPDF
from PIL import Image, ImageDraw, ImageChops
import numpy as np
import os
import tempfile

# SSIM 比较（可选）
try:
    from skimage.metrics import structural_similarity as ssim
    HAS_SSIM = True
except ImportError:
    HAS_SSIM = False


def pdf_to_images(pdf_path: str, dpi: int = 150) -> list:
    """将 PDF 转换为图像列表"""
    doc = fitz.open(pdf_path)
    images = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    doc.close()
    return images


def word_to_images_libreoffice(docx_path: str, dpi: int = 150) -> list:
    """使用 LibreOffice 将 Word 转换为图像"""
    import subprocess

    # 查找 LibreOffice
    libreoffice_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "soffice",
        "libreoffice"
    ]

    soffice = None
    for path in libreoffice_paths:
        try:
            result = subprocess.run([path, "--version"], capture_output=True, timeout=5)
            if result.returncode == 0:
                soffice = path
                break
        except:
            continue

    if not soffice:
        raise RuntimeError("未找到 LibreOffice，请安装或添加到 PATH")

    # 创建临时目录
    temp_dir = tempfile.mkdtemp()
    temp_pdf = None

    try:
        # 转换为 PDF
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", temp_dir, docx_path],
            capture_output=True, timeout=120
        )

        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice 转换失败: {result.stderr.decode()}")

        # 找到生成的 PDF
        base_name = os.path.splitext(os.path.basename(docx_path))[0]
        temp_pdf = os.path.join(temp_dir, base_name + ".pdf")

        if not os.path.exists(temp_pdf):
            # 尝试查找任何 PDF 文件
            for f in os.listdir(temp_dir):
                if f.endswith('.pdf'):
                    temp_pdf = os.path.join(temp_dir, f)
                    break

        if not temp_pdf or not os.path.exists(temp_pdf):
            raise RuntimeError("LibreOffice 未生成 PDF 文件")

        images = pdf_to_images(temp_pdf, dpi)
    finally:
        # 清理临时文件
        import shutil
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)

    return images


def word_to_images_docx2pdf(docx_path: str, dpi: int = 150) -> list:
    """使用 docx2pdf (Microsoft Word) 将 Word 转换为图像"""
    from docx2pdf import convert

    # 创建临时 PDF
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
        temp_pdf = tmp.name

    try:
        convert(docx_path, temp_pdf)
        images = pdf_to_images(temp_pdf, dpi)
    finally:
        if os.path.exists(temp_pdf):
            os.unlink(temp_pdf)

    return images


def word_to_images(docx_path: str, dpi: int = 150, method: str = "auto") -> list:
    """将 Word 转换为图像（支持 Word 或 LibreOffice）

    Args:
        docx_path: Word 文档路径
        dpi: 图像分辨率
        method: 转换方法 - "auto", "libreoffice", "word"
    """
    errors = []

    if method in ("auto", "libreoffice"):
        try:
            print("尝试使用 LibreOffice 转换...")
            return word_to_images_libreoffice(docx_path, dpi)
        except Exception as e:
            errors.append(f"LibreOffice: {e}")

    if method in ("auto", "word"):
        try:
            print("尝试使用 Microsoft Word 转换...")
            return word_to_images_docx2pdf(docx_path, dpi)
        except Exception as e:
            errors.append(f"Microsoft Word: {e}")

    raise RuntimeError(
        f"Word 文档转换失败。\n" +
        "\n".join(errors) +
        "\n\n解决方案:\n" +
        "1. 安装 LibreOffice: https://www.libreoffice.org/download/\n" +
        "2. 或使用 Microsoft Word 打开文档并另存为 PDF\n" +
        "3. 然后使用: python visual_compare.py your.pdf converted.pdf"
    )


def compare_images_pixel(img1: Image.Image, img2: Image.Image, threshold: int = 30):
    """像素级图像对比，返回差异图像和差异数量"""
    # 调整大小一致
    if img1.size != img2.size:
        img2 = img2.resize(img1.size, Image.LANCZOS)

    # 计算差异
    diff = ImageChops.difference(img1, img2)
    diff_gray = diff.convert('L')

    # 阈值处理
    diff_mask = diff_gray.point(lambda x: 255 if x > threshold else 0)

    # 统计差异数量
    diff_count = sum(1 for p in diff_mask.get_flattened_data() if p > 0)
    total_pixels = img1.width * img1.height
    diff_percent = diff_count / total_pixels * 100

    # 创建差异可视化（红色高亮）
    highlight = Image.new('RGB', img1.size, (255, 0, 0))
    result = Image.composite(highlight, img1, diff_mask)

    return result, diff_percent


def compare_images_ssim(img1: Image.Image, img2: Image.Image):
    """SSIM 结构相似性对比"""
    if not HAS_SSIM:
        raise ImportError("需要安装 scikit-image: pip install scikit-image")

    # 转灰度
    gray1 = np.array(img1.convert('L'))
    gray2 = np.array(img2.convert('L'))

    # 调整大小
    if gray1.shape != gray2.shape:
        from skimage.transform import resize
        gray2 = resize(gray2, gray1.shape, preserve_range=True).astype('uint8')

    # 计算 SSIM
    score, diff = ssim(gray1, gray2, full=True)

    # 差异图像
    diff_img = (diff * 255).astype('uint8')
    diff_pil = Image.fromarray(diff_img)

    return score, diff_pil


def compare_images_list(images1: list, images2: list, name1: str = "文件1", name2: str = "文件2", output_dir: str = "diff_output",
                        text_result: dict = None, layout_result: dict = None):
    """比较两组图像，生成差异报告"""

    os.makedirs(output_dir, exist_ok=True)

    results = []
    name1_base = os.path.splitext(os.path.basename(name1))[0]
    name2_base = os.path.splitext(os.path.basename(name2))[0]

    for i, (img1, img2) in enumerate(zip(images1, images2)):
        page_num = i + 1
        print(f"比较第 {page_num} 页...")

        # 像素对比
        diff_img, diff_percent = compare_images_pixel(img1, img2)

        # SSIM 对比（如果可用）
        ssim_score = None
        if HAS_SSIM:
            ssim_score, _ = compare_images_ssim(img1, img2)

        # 保存结果
        result = {
            'page': page_num,
            'diff_percent': diff_percent,
            'ssim_score': ssim_score,
            'similar': diff_percent < 5 and (ssim_score is None or ssim_score > 0.95)
        }
        results.append(result)

        # 保存图像
        img1.save(os.path.join(output_dir, f'{name1_base}_page_{page_num}.png'))
        img2.save(os.path.join(output_dir, f'{name2_base}_page_{page_num}.png'))
        diff_img.save(os.path.join(output_dir, f'diff_page_{page_num}.png'))

    # 生成 HTML 报告
    generate_report(results, output_dir, name1_base, name2_base, text_result, layout_result)

    return results


def compare_documents(pdf_path: str, docx_path: str, output_dir: str = "diff_output"):
    """比较 PDF 和 Word 文档，生成差异报告"""

    os.makedirs(output_dir, exist_ok=True)

    print(f"加载 PDF: {pdf_path}")
    pdf_images = pdf_to_images(pdf_path)

    print(f"加载 Word: {docx_path}")
    word_images = word_to_images(docx_path)

    results = []

    for i, (pdf_img, word_img) in enumerate(zip(pdf_images, word_images)):
        page_num = i + 1
        print(f"比较第 {page_num} 页...")

        # 像素对比
        diff_img, diff_percent = compare_images_pixel(pdf_img, word_img)

        # SSIM 对比（如果可用）
        ssim_score = None
        if HAS_SSIM:
            ssim_score, _ = compare_images_ssim(pdf_img, word_img)

        # 保存结果
        result = {
            'page': page_num,
            'diff_percent': diff_percent,
            'ssim_score': ssim_score,
            'similar': diff_percent < 5 and (ssim_score is None or ssim_score > 0.95)
        }
        results.append(result)

        # 保存图像
        pdf_img.save(os.path.join(output_dir, f'pdf_page_{page_num}.png'))
        word_img.save(os.path.join(output_dir, f'word_page_{page_num}.png'))
        diff_img.save(os.path.join(output_dir, f'diff_page_{page_num}.png'))

    # 生成 HTML 报告
    generate_report(results, output_dir)

    return results


def compare_text_content(pdf_path1: str, pdf_path2: str) -> dict:
    """比较两个PDF的文本内容匹配率（基于位置匹配）

    Returns:
        dict: {
            'text_match_rate': float,  # 文本匹配率 (0-1)
            'total_chars': int,        # 总字符数
            'matched_chars': int,      # 匹配字符数
            'missing_texts': list,     # 缺失的文本
            'extra_texts': list,       # 多余的文本
        }
    """
    import fitz

    def extract_texts_with_positions(pdf_path):
        doc = fitz.open(pdf_path)
        texts = []
        for page in doc:
            text_dict = page.get_text('dict')
            for block in text_dict.get('blocks', []):
                if block.get('type') == 0:
                    for line in block.get('lines', []):
                        for span in line.get('spans', []):
                            text = span.get('text', '').strip()
                            if text:
                                bbox = span.get('bbox', (0,0,0,0))
                                texts.append({
                                    'text': text,
                                    'x': bbox[0],
                                    'y': bbox[1]
                                })
        doc.close()
        return texts

    texts1 = extract_texts_with_positions(pdf_path1)
    texts2 = extract_texts_with_positions(pdf_path2)

    # 基于位置的文本匹配（类似布局匹配）
    used_indices = set()
    matched_count = 0
    missing_texts = []

    for t1 in texts1:
        # 找最近的相同文本
        best_idx = -1
        best_dist = float('inf')
        for i, t2 in enumerate(texts2):
            if i not in used_indices and t1['text'] == t2['text']:
                dist = abs(t1['y'] - t2['y']) + abs(t1['x'] - t2['x']) * 0.1
                if dist < best_dist:
                    best_dist = dist
                    best_idx = i

        if best_idx >= 0:
            used_indices.add(best_idx)
            matched_count += 1
        else:
            missing_texts.append(t1['text'])

    # 计算多余文本
    extra_texts = [texts2[i]['text'] for i in range(len(texts2)) if i not in used_indices]

    total_chars = sum(len(t['text']) for t in texts1)
    matched_chars = sum(len(t['text']) for t in texts1[:matched_count]) if matched_count > 0 else 0

    # 更准确的字符匹配计算：每个匹配的文本块贡献其字符数
    matched_chars = sum(len(texts1[i]['text']) for i in range(len(texts1)) if
                        any(texts1[i]['text'] == texts2[j]['text'] for j in used_indices))

    # 简化：匹配的文本块数 * 平均字符数
    matched_chars = sum(len(t['text']) for t in texts1[:matched_count])

    text_match_rate = matched_count / len(texts1) if texts1 else 0

    return {
        'text_match_rate': text_match_rate,
        'total_chars': total_chars,
        'matched_chars': matched_chars,
        'missing_texts': sorted(list(set(missing_texts)))[:20],
        'extra_texts': sorted(list(set(extra_texts)))[:20],
        'total_text_blocks': len(texts1),
        'matched_text_blocks': matched_count,
    }


def compare_layout_positions(pdf_path1: str, pdf_path2: str, tolerance: float = 10.0) -> dict:
    """比较两个PDF的布局位置匹配度

    Args:
        pdf_path1: 原始PDF路径
        pdf_path2: 待比较PDF路径
        tolerance: 位置容差(点)

    Returns:
        dict: {
            'layout_match_rate': float,  # 布局匹配率 (0-1)
            'avg_x_offset': float,       # 平均X偏移
            'avg_y_offset': float,       # 平均Y偏移
            'max_x_offset': float,       # 最大X偏移
            'max_y_offset': float,       # 最大Y偏移
            'position_details': list,    # 位置详情
        }
    """
    import fitz

    def extract_positions(pdf_path):
        doc = fitz.open(pdf_path)
        positions = []
        for page in doc:
            text_dict = page.get_text('dict')
            for block in text_dict.get('blocks', []):
                if block.get('type') == 0:
                    for line in block.get('lines', []):
                        for span in line.get('spans', []):
                            text = span.get('text', '').strip()
                            if text:
                                bbox = span.get('bbox', (0,0,0,0))
                                positions.append({
                                    'text': text,
                                    'x': bbox[0],
                                    'y': bbox[1],
                                    'size': span.get('size', 0),
                                    'bold': bool(span.get('flags', 0) & 16)
                                })
        doc.close()
        return positions

    pos1 = extract_positions(pdf_path1)
    pos2 = extract_positions(pdf_path2)

    # 改进的匹配算法：为每个原文文本找最近的相同文本
    matched_positions = []
    x_offsets = []
    y_offsets = []
    used_indices = set()  # 记录已使用的生成文本索引

    for p1 in pos1:
        # 找相同文本且位置最近的
        best_match_idx = -1
        best_dist = float('inf')

        for i, p2 in enumerate(pos2):
            if i not in used_indices and p1['text'] == p2['text']:
                # 计算距离（优先Y距离，因为同一行内顺序可能不同）
                dist = abs(p1['y'] - p2['y']) + abs(p1['x'] - p2['x']) * 0.1
                if dist < best_dist:
                    best_dist = dist
                    best_match_idx = i

        if best_match_idx >= 0:
            used_indices.add(best_match_idx)
            p2 = pos2[best_match_idx]
            x_off = p2['x'] - p1['x']
            y_off = p2['y'] - p1['y']

            matched_positions.append({
                'text': p1['text'][:20],
                'orig_x': p1['x'],
                'orig_y': p1['y'],
                'gen_x': p2['x'],
                'gen_y': p2['y'],
                'x_offset': x_off,
                'y_offset': y_off,
                'in_tolerance': abs(x_off) <= tolerance and abs(y_off) <= tolerance
            })

            x_offsets.append(x_off)
            y_offsets.append(y_off)

    # 计算指标
    in_tolerance_count = sum(1 for p in matched_positions if p['in_tolerance'])
    layout_match_rate = in_tolerance_count / len(matched_positions) if matched_positions else 0

    return {
        'layout_match_rate': layout_match_rate,
        'avg_x_offset': sum(x_offsets) / len(x_offsets) if x_offsets else 0,
        'avg_y_offset': sum(y_offsets) / len(y_offsets) if y_offsets else 0,
        'max_x_offset': max(abs(x) for x in x_offsets) if x_offsets else 0,
        'max_y_offset': max(abs(y) for y in y_offsets) if y_offsets else 0,
        'position_details': matched_positions[:30],  # 最多显示30个
        'total_matched': len(matched_positions),
        'in_tolerance_count': in_tolerance_count,
    }


def generate_report(results: list, output_dir: str, name1: str = "文件1", name2: str = "文件2",
                   text_result: dict = None, layout_result: dict = None):
    """生成 HTML 对比报告"""

    # 计算综合评分
    pixel_score = 1 - (results[0]['diff_percent'] / 100) if results else 0
    text_score = text_result['text_match_rate'] if text_result else 0
    layout_score = layout_result['layout_match_rate'] if layout_result else 0
    overall = pixel_score * 0.3 + text_score * 0.3 + layout_score * 0.4

    html = ['''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>文档对比报告</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; }
        .page { margin-bottom: 30px; padding: 20px; border: 1px solid #ddd; background: white; border-radius: 8px; }
        .match { color: green; font-weight: bold; }
        .differ { color: red; font-weight: bold; }
        .images { display: flex; gap: 20px; flex-wrap: wrap; }
        .images img { max-width: 300px; border: 1px solid #ccc; }
        table { border-collapse: collapse; margin: 10px 0; width: 100%; }
        td, th { border: 1px solid #ddd; padding: 8px; text-align: center; }
        th { background: #4a90d9; color: white; }
        .score-table { width: 400px; margin: 20px auto; }
        .score-table td, .score-table th { padding: 12px; }
        .good { background: #d4edda; }
        .warning { background: #fff3cd; }
        .bad { background: #f8d7da; }
        .overall { background: #4a90d9; color: white; font-weight: bold; font-size: 1.2em; }
        .metrics-panel { background: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        h1 { color: #333; text-align: center; }
        h2 { color: #4a90d9; border-bottom: 2px solid #4a90d9; padding-bottom: 10px; }
    </style>
</head>
<body>
    <div class="container">
    <h1>📊 文档对比报告</h1>
    <p style="text-align:center; color:#666;">对比: ''' + name1 + ''' vs ''' + name2 + '''</p>
''']

    # 评判指标表格
    html.append('''
    <div class="metrics-panel">
        <h2>📈 评判指标</h2>
        <table class="score-table">
            <tr>
                <th>指标</th>
                <th>值</th>
                <th>权重</th>
            </tr>
            <tr class="''' + ('good' if pixel_score > 0.9 else 'warning' if pixel_score > 0.8 else 'bad') + '''">
                <td>像素相似度</td>
                <td>''' + f'{pixel_score*100:.1f}%' + '''</td>
                <td>30%</td>
            </tr>
            <tr class="''' + ('good' if text_score > 0.9 else 'warning' if text_score > 0.7 else 'bad') + '''">
                <td>文本匹配率</td>
                <td>''' + f'{text_score*100:.1f}%' + '''</td>
                <td>30%</td>
            </tr>
            <tr class="''' + ('good' if layout_score > 0.9 else 'warning' if layout_score > 0.7 else 'bad') + '''">
                <td>布局匹配度</td>
                <td>''' + f'{layout_score*100:.1f}%' + '''</td>
                <td>40%</td>
            </tr>
            <tr class="overall">
                <td>综合评分</td>
                <td>''' + f'{overall*100:.1f}%' + '''</td>
                <td>-</td>
            </tr>
        </table>
    </div>
''')

    # 详细指标
    if layout_result:
        html.append('''
    <div class="metrics-panel">
        <h2>📐 布局详情</h2>
        <table>
            <tr><th>指标</th><th>值</th></tr>
            <tr><td>平均X偏移</td><td>''' + f'{layout_result["avg_x_offset"]:.1f}pt' + '''</td></tr>
            <tr><td>平均Y偏移</td><td>''' + f'{layout_result["avg_y_offset"]:.1f}pt' + '''</td></tr>
            <tr><td>最大X偏移</td><td>''' + f'{layout_result["max_x_offset"]:.1f}pt' + '''</td></tr>
            <tr><td>最大Y偏移</td><td>''' + f'{layout_result["max_y_offset"]:.1f}pt' + '''</td></tr>
            <tr><td>位置容差内</td><td>''' + f'{layout_result["in_tolerance_count"]}/{layout_result["total_matched"]}' + '''</td></tr>
        </table>
    </div>
''')

    if text_result:
        html.append('''
    <div class="metrics-panel">
        <h2>📝 文本详情</h2>
        <table>
            <tr><th>指标</th><th>值</th></tr>
            <tr><td>匹配字符数</td><td>''' + f'{text_result["matched_chars"]}/{text_result["total_chars"]}' + '''</td></tr>
            <tr><td>匹配文本块</td><td>''' + f'{text_result["matched_text_blocks"]}/{text_result["total_text_blocks"]}' + '''</td></tr>
        </table>
    </div>
''')

    for r in results:
        status = "✓ 匹配" if r['similar'] else "✗ 差异"
        css_class = "match" if r['similar'] else "differ"

        html.append(f'''
    <div class="page">
        <h2>第 {r['page']} 页: <span class="{css_class}">{status}</span></h2>
        <table>
            <tr><th>指标</th><th>值</th></tr>
            <tr><td>差异像素</td><td>{r['diff_percent']:.2f}%</td></tr>
            {'<tr><td>SSIM</td><td>' + f'{r["ssim_score"]:.4f}' + '</td></tr>' if r['ssim_score'] else ''}
        </table>
        <div class="images">
            <div><h4>{name1}</h4><img src="{name1}_page_{r['page']}.png"></div>
            <div><h4>{name2}</h4><img src="{name2}_page_{r['page']}.png"></div>
            <div><h4>差异</h4><img src="diff_page_{r['page']}.png"></div>
        </div>
    </div>
''')

    html.append('''
    </div>
</body>
</html>
''')

    report_path = os.path.join(output_dir, 'report.html')
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(html))

    print(f"报告已生成: {report_path}")


if __name__ == "__main__":
    import sys

    # 检查参数
    if len(sys.argv) >= 3:
        file1 = sys.argv[1]
        file2 = sys.argv[2]
    else:
        file1 = "OC-D P241210003.pdf"
        file2 = "OC-D P241210003_v8_optimized.docx"

    # 自动检测文件类型
    ext1 = os.path.splitext(file1)[1].lower()
    ext2 = os.path.splitext(file2)[1].lower()

    print(f"文件1: {file1} ({ext1})")
    print(f"文件2: {file2} ({ext2})")
    print()

    # 加载图像
    if ext1 == '.pdf':
        print(f"加载 PDF: {file1}")
        images1 = pdf_to_images(file1)
    elif ext1 in ('.docx', '.doc'):
        print(f"加载 Word: {file1}")
        images1 = word_to_images(file1)
    else:
        print(f"错误: 不支持的文件格式 {ext1}")
        sys.exit(1)

    if ext2 == '.pdf':
        print(f"加载 PDF: {file2}")
        images2 = pdf_to_images(file2)
    elif ext2 in ('.docx', '.doc'):
        print(f"加载 Word: {file2}")
        images2 = word_to_images(file2)
    else:
        print(f"错误: 不支持的文件格式 {ext2}")
        sys.exit(1)

    # 像素级对比
    results = compare_images_list(images1, images2, file1, file2)

    print("\n" + "=" * 60)
    print("像素级对比结果")
    print("=" * 60)
    for r in results:
        status = "[OK] 匹配" if r['similar'] else "[!!] 有差异"
        print(f"第 {r['page']} 页: {status}")
        print(f"  差异像素: {r['diff_percent']:.2f}%")
        if r['ssim_score']:
            print(f"  SSIM: {r['ssim_score']:.4f}")

    # 如果两个都是PDF，进行文本和布局对比
    text_result = None
    layout_result = None

    if ext1 == '.pdf' and ext2 == '.pdf':
        print("\n" + "=" * 60)
        print("文本内容匹配率")
        print("=" * 60)
        text_result = compare_text_content(file1, file2)
        print(f"文本匹配率: {text_result['text_match_rate']*100:.2f}%")
        print(f"匹配字符数: {text_result['matched_chars']}/{text_result['total_chars']}")
        print(f"匹配文本块: {text_result['matched_text_blocks']}/{text_result['total_text_blocks']}")

        if text_result['missing_texts']:
            print(f"\n缺失文本 (前5个):")
            for t in text_result['missing_texts'][:5]:
                print(f"  - {t[:50]}")

        if text_result['extra_texts']:
            print(f"\n多余文本 (前5个):")
            for t in text_result['extra_texts'][:5]:
                print(f"  + {t[:50]}")

        print("\n" + "=" * 60)
        print("布局位置匹配度")
        print("=" * 60)
        layout_result = compare_layout_positions(file1, file2, tolerance=10.0)
        print(f"布局匹配率: {layout_result['layout_match_rate']*100:.2f}%")
        print(f"平均X偏移: {layout_result['avg_x_offset']:.1f}pt")
        print(f"平均Y偏移: {layout_result['avg_y_offset']:.1f}pt")
        print(f"最大X偏移: {layout_result['max_x_offset']:.1f}pt")
        print(f"最大Y偏移: {layout_result['max_y_offset']:.1f}pt")
        print(f"位置容差内: {layout_result['in_tolerance_count']}/{layout_result['total_matched']}")

        # 显示偏移较大的位置
        large_offsets = [p for p in layout_result['position_details']
                        if abs(p['x_offset']) > 10 or abs(p['y_offset']) > 10]
        if large_offsets:
            print(f"\n偏移较大的文本 (前10个):")
            for p in large_offsets[:10]:
                print(f"  Y={p['orig_y']:.0f}->{p['gen_y']:.0f} ({p['y_offset']:+.0f}pt): {p['text'][:25]}")

    # 像素级对比 - 传入text_result和layout_result
    results = compare_images_list(images1, images2, file1, file2, text_result=text_result, layout_result=layout_result)

    print("\n" + "=" * 60)
    print("像素级对比结果")
    print("=" * 60)
    for r in results:
        status = "[OK] 匹配" if r['similar'] else "[!!] 有差异"
        print(f"第 {r['page']} 页: {status}")
        print(f"  差异像素: {r['diff_percent']:.2f}%")
        if r['ssim_score']:
            print(f"  SSIM: {r['ssim_score']:.4f}")

    # 综合评分（仅在两个PDF时计算）
    if ext1 == '.pdf' and ext2 == '.pdf' and text_result and layout_result:
        print("\n" + "=" * 60)
        print("综合评分")
        print("=" * 60)
        pixel_score = 1 - (results[0]['diff_percent'] / 100) if results else 0
        text_score = text_result['text_match_rate']
        layout_score = layout_result['layout_match_rate']

        overall = pixel_score * 0.3 + text_score * 0.3 + layout_score * 0.4
        print(f"像素相似度: {pixel_score*100:.1f}% (权重30%)")
        print(f"文本匹配率: {text_score*100:.1f}% (权重30%)")
        print(f"布局匹配度: {layout_score*100:.1f}% (权重40%)")
        print(f"\n综合评分: {overall*100:.1f}%")