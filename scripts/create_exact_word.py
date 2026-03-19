#!/usr/bin/env python3
"""
精确复制 PDF 布局到 Word - v12 版本
根据 PDF 的精确坐标生成完全一致的 Word 文档
主要改进:
1. 所有字体改为 TimesNewRoman
2. 使用Tab定位实现精确的页眉布局
3. 客户信息使用Tab分隔符对齐
4. 表格列宽按PDF X坐标调整
"""

from docx import Document
from docx.shared import Pt, Cm, Inches, Twips, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement
import re


def set_run_font(run, font_name: str, size_pt: float, bold: bool = False, italic: bool = False):
    """设置 run 的字体属性"""
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    # 设置东亚字体
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), font_name)


def remove_table_borders(table):
    """移除整个表格的边框"""
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'nil')
        tblBorders.append(border)
    tblPr.append(tblBorders)


def pt_to_cm(pt: float) -> float:
    """将点转换为厘米"""
    return pt * 2.54 / 72


def create_exact_word_document(output_path: str):
    """
    根据分析的 PDF 布局创建精确匹配的 Word 文档

    PDF坐标参考 (A4: 595x842 pt, 21x29.7 cm):
    - 页眉公司信息: X=184pt (6.5cm)
    - Page/Date: X=482pt (17cm)
    - 标题: Y=184pt, X=187pt (居中)
    - 客户信息标签: X=28pt
    - 冒号: X=136pt
    - 值: X=140pt
    """
    doc = Document()

    # 设置页面为 A4
    section = doc.sections[0]
    section.page_width = Cm(21.0)   # A4 宽度
    section.page_height = Cm(29.7)  # A4 高度

    # 设置页边距
    # PDF左边距约28pt=1cm，我们需要匹配
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(1.0)
    section.left_margin = Cm(1.0)   # 28pt ≈ 1cm
    section.right_margin = Cm(1.0)

    # =====================================================
    # 页眉区域 - 使用Tab定位实现精确布局
    # PDF坐标:
    # 公司信息 X=184pt = 6.5cm (从页面左边)
    # Page/Date X=482pt = 17cm (从页面左边)
    # 但Word中Tab是相对于段落缩进，所以需要减去左边距
    # Tab位置 = X位置 - 左边距 = 184pt - 28pt = 156pt = 5.5cm
    # 右侧Tab位置 = 482pt - 28pt = 454pt = 16cm
    #
    # Y坐标间隔:
    # Pulcra: Y=74
    # Isardamm: Y=86 (间隔12pt)
    # Page: Y=89 (与Isardamm同行)
    # 82538: Y=99 (间隔13pt)
    # DEUTSCHLAND: Y=111 (间隔12pt)
    # Date: Y=131 (间隔20pt)
    # =====================================================

    # 计算Tab位置 (相对于左边距)
    # 公司信息 X=184pt = 6.5cm，减去左边距28pt = 156pt = 5.5cm
    company_tab = Cm(5.5)
    # Page/Date 左对齐位置 X=482pt，减去左边距28pt = 454pt = 16cm
    right_tab = Cm(16.0)  # 这是LEFT tab，不是RIGHT tab！

    # 设置段落默认格式 - 减少行间距
    style = doc.styles['Normal']
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.line_spacing = 1.0

    # 第1行: Pulcra Chemicals GmbH (Y=74)
    # 需要从Y=28pt到Y=74pt，增加46pt的顶部空间
    para = doc.add_paragraph()
    para.paragraph_format.tab_stops.add_tab_stop(company_tab, WD_TAB_ALIGNMENT.LEFT)
    para.paragraph_format.space_before = Pt(46)  # 74pt - 28pt = 46pt
    run = para.add_run('\tPulcra Chemicals GmbH')
    set_run_font(run, 'TimesNewRoman', 11)

    # 第2行: Isardamm 79-83 + Page: 1/1 (Y=86-89)
    # 间隔: 86-74=12pt
    # Page在X=482pt，使用LEFT tab
    para = doc.add_paragraph()
    para.paragraph_format.tab_stops.add_tab_stop(company_tab, WD_TAB_ALIGNMENT.LEFT)
    para.paragraph_format.tab_stops.add_tab_stop(right_tab, WD_TAB_ALIGNMENT.LEFT)  # LEFT tab!
    para.paragraph_format.space_before = Pt(0)
    run = para.add_run('\tIsardamm 79-83\tPage: 1/1')
    set_run_font(run, 'TimesNewRoman', 11)

    # 第3行: 82538 Geretsried (Y=99)
    # 间隔: 99-89=10pt
    para = doc.add_paragraph()
    para.paragraph_format.tab_stops.add_tab_stop(company_tab, WD_TAB_ALIGNMENT.LEFT)
    para.paragraph_format.space_before = Pt(0)
    run = para.add_run('\t82538 Geretsried')
    set_run_font(run, 'TimesNewRoman', 10)

    # 第4行: DEUTSCHLAND (Y=111)
    # 间隔: 111-99=12pt
    para = doc.add_paragraph()
    para.paragraph_format.tab_stops.add_tab_stop(company_tab, WD_TAB_ALIGNMENT.LEFT)
    para.paragraph_format.space_before = Pt(0)
    run = para.add_run('\tDEUTSCHLAND')
    set_run_font(run, 'TimesNewRoman', 10)

    # 第5行: 2025.07.16 (Y=131)
    # 间隔: 131-111=20pt (需要额外空间因为上一行是DEUTSCHLAND在Y=111)
    # 但当前DEUTSCHLAND生成在Y=112，所以需要131-112=19pt额外空间
    # 加上行高约12pt，实际需要19-12=7pt space_before
    # Date在X=482pt，使用LEFT tab
    para = doc.add_paragraph()
    para.paragraph_format.tab_stops.add_tab_stop(right_tab, WD_TAB_ALIGNMENT.LEFT)  # LEFT tab!
    para.paragraph_format.space_before = Pt(8)  # 增加额外空间
    run = para.add_run('\t2025.07.16')
    set_run_font(run, 'TimesNewRoman', 11)

    # 标题在Y=184，日期在Y=131，间隔53pt
    # 标题需要居中
    # 标题是16pt字体，行高约19pt，基线到顶部约13pt
    # 需要从Y=131到Y=184-13=171，间隔40pt
    # 减去日期行高12pt，还需28pt
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(28)

    # =====================================================
    # 标题：CERTIFICATE OF ANALYSIS（居中，加粗，16pt）
    # PDF位置: Y=184, 居中
    # =====================================================
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before = Pt(0)
    title.paragraph_format.space_after = Pt(0)
    run = title.add_run("CERTIFICATE OF ANALYSIS")
    set_run_font(run, 'TimesNewRoman', 16, bold=True)

    # 客户信息在Y=208，标题在Y=184，间隔24pt
    # 标题是16pt字体，基线到顶部约13pt
    # 标题在Y=185，需要到Customer Y=208
    # 间隔 = 208 - 185 - 13 (标题高度) = 10pt
    # 但需要考虑空段落的高度
    # 直接在标题后添加客户信息，不需要空段落
    # =====================================================
    # 客户信息
    # PDF布局: 标签 X=28pt, 冒号 X=136pt (冒号起始位置), 值 X=140pt
    # 值文本带前导空格: " Jiangsu..."
    # 文本顺序: 标签 -> tab -> ": " -> 值
    # =====================================================
    label_tab = Cm(3.8)   # 冒号起始位置 (108pt from margin)

    customer_data = [
        ("Customer", "Jiangsu Mingxin Xuteng Technology C"),
        ("Customer Nr.", "1508866"),
        ("Product Name", "FORYL OC-D(I104)"),
        ("Product Nr.", "20660"),
    ]

    # 直接添加客户信息，调整间距
    for i, (label, value) in enumerate(customer_data):
        para = doc.add_paragraph()
        para.paragraph_format.tab_stops.add_tab_stop(label_tab, WD_TAB_ALIGNMENT.LEFT)
        # 第一个客户信息需要额外间距 (标题Y=185 + 标题高度13pt = 198, 到Y=208还需10pt)
        # 但实际生成Y=214，偏移+6，所以减少space_before
        if i == 0:
            para.paragraph_format.space_before = Pt(5)  # 减少间距
        else:
            para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)

        # 标签加粗
        run1 = para.add_run(label)
        set_run_font(run1, 'TimesNewRoman', 11, bold=True)
        # Tab + 冒号 + 空格 (加粗)
        run2 = para.add_run('\t: ')
        set_run_font(run2, 'TimesNewRoman', 11, bold=True)
        # 值
        run3 = para.add_run(value)
        set_run_font(run3, 'TimesNewRoman', 11)

    # =====================================================
    # 批次信息
    # =====================================================
    add_batch_section(doc, "P241200117", "2024.04.29", "2026.04.29", "40000403286",
                      "6.9", "35.3", first_batch=True)
    add_batch_section(doc, "P241210003", "2024.04.30", "2026.04.30", "90000071881",
                      "6.8", "35.7", first_batch=False)

    # =====================================================
    # Released by
    # Released偏移+39.9pt(位置偏低)，需要减少space_before使其上升
    # =====================================================
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(10)  # 从20pt减少到10pt
    run = para.add_run("Released by: ")
    set_run_font(run, 'TimesNewRoman', 11, bold=True)
    run = para.add_run("SILKE STEIER")
    set_run_font(run, 'TimesNewRoman', 11, bold=True)

    # =====================================================
    # 免责声明（斜体，11pt）
    # 原PDF分成3行，需要保持一致
    # =====================================================
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    run = para.add_run("The above data represent the results of our quality assessment.")
    set_run_font(run, 'TimesNewRoman', 11, italic=True)

    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    run = para.add_run("They do not free the purchaser from his own quality check nor do they confirm that the product has certain properties or")
    set_run_font(run, 'TimesNewRoman', 11, italic=True)

    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    run = para.add_run("is suitable for a specific application.")
    set_run_font(run, 'TimesNewRoman', 11, italic=True)

    # =====================================================
    # 页脚
    # =====================================================
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(130)  # v33最佳值
    run = para.add_run("DIN EN 10204")
    set_run_font(run, 'TimesNewRoman', 8)

    para = doc.add_paragraph()
    run = para.add_run("This certificate is printed out electronically, therefore it has no signature.")
    set_run_font(run, 'TimesNewRoman', 8)

    para = doc.add_paragraph()
    run = para.add_run("Quality Control Department")
    set_run_font(run, 'TimesNewRoman', 8, bold=True)

    doc.save(output_path)
    print(f"文档已保存: {output_path}")


def add_batch_section(doc, batch_no: str, prod_date: str, exp_date: str, lot_no: str,
                      ph_result: str, water_result: str, first_batch: bool = False):
    """添加批次信息部分

    表格列宽按PDF X坐标计算 (pt转cm):
    - Specification: X=28pt (1cm), 宽度=201-28=173pt (6.1cm)
    - Method: X=201pt (7.1cm), 宽度=280-201=79pt (2.8cm)
    - Unit: X=280pt (9.9cm), 宽度=316-280=36pt (1.3cm)
    - Result: X=316pt (11.2cm), 宽度=416-316=100pt (3.5cm)
    - Standard: X=416pt (14.7cm), 宽度至边距
    """
    # Tab位置用于冒号对齐 (与客户信息相同)
    label_tab = Cm(3.8)  # 冒号起始位置

    # Batch Number - 第一个批次需要额外间距
    para = doc.add_paragraph()
    para.paragraph_format.tab_stops.add_tab_stop(label_tab, WD_TAB_ALIGNMENT.LEFT)
    if first_batch:
        para.paragraph_format.space_before = Pt(46)  # v33=48, 减少2pt使Batch Number上移
    else:
        para.paragraph_format.space_before = Pt(18)
    run = para.add_run("Batch Number")
    set_run_font(run, 'TimesNewRoman', 11, bold=True)
    run = para.add_run('\t: ')
    set_run_font(run, 'TimesNewRoman', 11, bold=True)
    run = para.add_run(batch_no)
    set_run_font(run, 'TimesNewRoman', 11)

    # Production Date
    para = doc.add_paragraph()
    para.paragraph_format.tab_stops.add_tab_stop(label_tab, WD_TAB_ALIGNMENT.LEFT)
    run = para.add_run("Production Date")
    set_run_font(run, 'TimesNewRoman', 11, bold=True)
    run = para.add_run('\t: ')
    set_run_font(run, 'TimesNewRoman', 11, bold=True)
    run = para.add_run(prod_date)
    set_run_font(run, 'TimesNewRoman', 11)

    # Expiration Date
    para = doc.add_paragraph()
    para.paragraph_format.tab_stops.add_tab_stop(label_tab, WD_TAB_ALIGNMENT.LEFT)
    run = para.add_run("Expiration Date")
    set_run_font(run, 'TimesNewRoman', 11, bold=True)
    run = para.add_run('\t: ')
    set_run_font(run, 'TimesNewRoman', 11, bold=True)
    run = para.add_run(exp_date)
    set_run_font(run, 'TimesNewRoman', 11)

    # Inspection Lot
    para = doc.add_paragraph()
    para.paragraph_format.tab_stops.add_tab_stop(label_tab, WD_TAB_ALIGNMENT.LEFT)
    run = para.add_run("Inspection Lot")
    set_run_font(run, 'TimesNewRoman', 11, bold=True)
    run = para.add_run('\t: ')
    set_run_font(run, 'TimesNewRoman', 11, bold=True)
    run = para.add_run(lot_no)
    set_run_font(run, 'TimesNewRoman', 11)

    # 表格区域 - 使用段落+Tab替代表格，精确控制行间距
    # PDF位置: Specification Y=356, 分隔线 Y=365, AUSSEHEN Y=373, PH Y=390, WASSERGEHALT Y=407
    # 表头Tab位置: Method X=201, Unit X=280, Result X=316, Standard X=416
    spec_tab = Cm(6.1)   # 201-28=173pt=6.1cm
    unit_tab = Cm(8.9)   # 280-28=252pt=8.9cm
    result_tab = Cm(10.1) # 316-28=288pt=10.1cm
    std_tab = Cm(13.7)   # 416-28=388pt=13.7cm

    # 表头行 - Specification偏移+8pt，需要减少space_before
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(2)  # 从10pt减少到2pt
    para.paragraph_format.tab_stops.add_tab_stop(spec_tab, WD_TAB_ALIGNMENT.LEFT)
    para.paragraph_format.tab_stops.add_tab_stop(unit_tab, WD_TAB_ALIGNMENT.LEFT)
    para.paragraph_format.tab_stops.add_tab_stop(result_tab, WD_TAB_ALIGNMENT.LEFT)
    para.paragraph_format.tab_stops.add_tab_stop(std_tab, WD_TAB_ALIGNMENT.LEFT)
    run = para.add_run("Specification\tMethod\tUnit\tResult\tStandard")
    set_run_font(run, 'TimesNewRoman', 8, bold=True)

    # 分隔线行
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.tab_stops.add_tab_stop(spec_tab, WD_TAB_ALIGNMENT.LEFT)
    para.paragraph_format.tab_stops.add_tab_stop(unit_tab, WD_TAB_ALIGNMENT.LEFT)
    para.paragraph_format.tab_stops.add_tab_stop(result_tab, WD_TAB_ALIGNMENT.LEFT)
    para.paragraph_format.tab_stops.add_tab_stop(std_tab, WD_TAB_ALIGNMENT.LEFT)
    # 使用完整的分隔线字符串（与原PDF匹配）
    run = para.add_run("_______________________________ _______________________________ _______________________________ _______________________________ _")
    set_run_font(run, 'TimesNewRoman', 6, bold=True)

    # 数据行 1: AUSSEHEN (Y=373)
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.tab_stops.add_tab_stop(spec_tab, WD_TAB_ALIGNMENT.LEFT)
    para.paragraph_format.tab_stops.add_tab_stop(unit_tab, WD_TAB_ALIGNMENT.LEFT)
    para.paragraph_format.tab_stops.add_tab_stop(result_tab, WD_TAB_ALIGNMENT.LEFT)
    para.paragraph_format.tab_stops.add_tab_stop(std_tab, WD_TAB_ALIGNMENT.LEFT)
    run = para.add_run(f"AUSSEHEN;20°C\t\t\tCOLORLESS TO YELLOWI\tCOLORLESS TO YELLOWISH")
    set_run_font(run, 'TimesNewRoman', 8)

    # 数据行 2: PH (Y=390，间隔17pt)
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(5)  # 增加行间距
    para.paragraph_format.tab_stops.add_tab_stop(spec_tab, WD_TAB_ALIGNMENT.LEFT)
    para.paragraph_format.tab_stops.add_tab_stop(unit_tab, WD_TAB_ALIGNMENT.LEFT)
    para.paragraph_format.tab_stops.add_tab_stop(result_tab, WD_TAB_ALIGNMENT.LEFT)
    para.paragraph_format.tab_stops.add_tab_stop(std_tab, WD_TAB_ALIGNMENT.LEFT)
    run = para.add_run(f"PH;10%\t\t\t{ph_result}\t6.0 -8.0")
    set_run_font(run, 'TimesNewRoman', 8)

    # 数据行 3: WASSERGEHALT (Y=407，间隔17pt)
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(5)  # 增加行间距
    para.paragraph_format.tab_stops.add_tab_stop(spec_tab, WD_TAB_ALIGNMENT.LEFT)
    para.paragraph_format.tab_stops.add_tab_stop(unit_tab, WD_TAB_ALIGNMENT.LEFT)
    para.paragraph_format.tab_stops.add_tab_stop(result_tab, WD_TAB_ALIGNMENT.LEFT)
    para.paragraph_format.tab_stops.add_tab_stop(std_tab, WD_TAB_ALIGNMENT.LEFT)
    run = para.add_run(f"WASSERGEHALT,KARL FISCHER\t\t%\t{water_result}\t34.0 -37.0")
    set_run_font(run, 'TimesNewRoman', 8)

    # 批次后空行
    doc.add_paragraph()


if __name__ == "__main__":
    output_path = r"C:\000_claude\mylove\OC-D P241210003_v36.docx"
    create_exact_word_document(output_path)