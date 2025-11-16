#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
详细分析论文格式模板文件
"""
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import re

def detailed_analysis(file_path):
    """详细分析docx文件"""
    doc = Document(file_path)

    print("=" * 80)
    print("论文格式模板详细分析报告")
    print("=" * 80)

    # 分析页面设置
    section = doc.sections[0]
    print("\n【页面设置】")
    print(f"页边距:")
    print(f"  上: {section.top_margin.cm:.2f} cm ({section.top_margin.pt:.1f} pt)")
    print(f"  下: {section.bottom_margin.cm:.2f} cm ({section.bottom_margin.pt:.1f} pt)")
    print(f"  左: {section.left_margin.cm:.2f} cm ({section.left_margin.pt:.1f} pt)")
    print(f"  右: {section.right_margin.cm:.2f} cm ({section.right_margin.pt:.1f} pt)")
    print(f"页面尺寸: {section.page_width.cm:.2f} x {section.page_height.cm:.2f} cm")
    print(f"页眉距边界: {section.header_distance.cm:.2f} cm")
    print(f"页脚距边界: {section.footer_distance.cm:.2f} cm")

    # 详细分析每个段落
    print("\n【段落详细分析】")
    print("-" * 80)

    for i, para in enumerate(doc.paragraphs[:50]):  # 分析前50个段落
        if not para.text.strip():
            continue

        print(f"\n段落 {i+1}:")
        print(f"内容: {para.text[:80]}{'...' if len(para.text) > 80 else ''}")
        print(f"样式名称: {para.style.name}")

        # 段落格式
        pf = para.paragraph_format
        print(f"段落格式:")
        if para.alignment is not None:
            align_map = {0: '左对齐', 1: '居中', 2: '右对齐', 3: '两端对齐', 4: '分散对齐'}
            print(f"  对齐方式: {align_map.get(para.alignment, para.alignment)}")

        if pf.line_spacing is not None:
            if isinstance(pf.line_spacing, float):
                print(f"  行距: {pf.line_spacing} 倍")
            else:
                print(f"  行距: {pf.line_spacing}")

        if pf.line_spacing_rule is not None:
            rule_map = {0: '最小值', 1: '固定值', 2: '多倍行距'}
            print(f"  行距规则: {rule_map.get(pf.line_spacing_rule, pf.line_spacing_rule)}")

        if pf.space_before is not None and pf.space_before.pt > 0:
            print(f"  段前间距: {pf.space_before.pt:.1f} pt ({pf.space_before.cm:.2f} cm)")

        if pf.space_after is not None and pf.space_after.pt > 0:
            print(f"  段后间距: {pf.space_after.pt:.1f} pt ({pf.space_after.cm:.2f} cm)")

        if pf.first_line_indent is not None and pf.first_line_indent != 0:
            print(f"  首行缩进: {pf.first_line_indent.cm:.2f} cm ({pf.first_line_indent.pt:.1f} pt)")

        if pf.left_indent is not None and pf.left_indent != 0:
            print(f"  左缩进: {pf.left_indent.cm:.2f} cm")

        if pf.right_indent is not None and pf.right_indent != 0:
            print(f"  右缩进: {pf.right_indent.cm:.2f} cm")

        # 字体格式
        if para.runs:
            print(f"字体格式 (第一个run):")
            run = para.runs[0]

            if run.font.name:
                print(f"  字体名称(ASCII): {run.font.name}")

            # 获取中文字体
            if run._element.rPr is not None:
                rFonts = run._element.rPr.rFonts
                if rFonts is not None:
                    from docx.oxml.ns import qn
                    east_asia = rFonts.get(qn('w:eastAsia'))
                    if east_asia:
                        print(f"  中文字体: {east_asia}")

            if run.font.size:
                print(f"  字号: {run.font.size.pt} pt")

            if run.font.bold is not None:
                print(f"  加粗: {'是' if run.font.bold else '否'}")

            if run.font.italic is not None:
                print(f"  斜体: {'是' if run.font.italic else '否'}")

            if run.font.underline is not None:
                print(f"  下划线: {run.font.underline}")

            if run.font.color.rgb:
                print(f"  颜色: {run.font.color.rgb}")

    # 分析表格
    print(f"\n{'='*80}")
    print(f"【表格分析】")
    print(f"表格总数: {len(doc.tables)}")

    for i, table in enumerate(doc.tables[:5]):  # 分析前5个表格
        print(f"\n表格 {i+1}:")
        print(f"  行数: {len(table.rows)}")
        print(f"  列数: {len(table.columns)}")
        print(f"  表格样式: {table.style.name if table.style else 'None'}")

        # 分析第一个单元格的格式
        if table.rows and table.rows[0].cells:
            cell = table.rows[0].cells[0]
            if cell.paragraphs:
                p = cell.paragraphs[0]
                if p.runs:
                    print(f"  单元格字体格式:")
                    run = p.runs[0]
                    if run.font.name:
                        print(f"    字体: {run.font.name}")
                    if run.font.size:
                        print(f"    字号: {run.font.size.pt} pt")

    # 分析样式定义
    print(f"\n{'='*80}")
    print(f"【文档样式定义】")

    for style in doc.styles:
        if style.type == 1:  # 段落样式
            print(f"\n样式: {style.name}")
            if hasattr(style, 'font'):
                if style.font.name:
                    print(f"  字体: {style.font.name}")
                if style.font.size:
                    print(f"  字号: {style.font.size.pt} pt")

            if hasattr(style, 'paragraph_format'):
                pf = style.paragraph_format
                if pf.alignment is not None:
                    align_map = {0: '左对齐', 1: '居中', 2: '右对齐', 3: '两端对齐'}
                    print(f"  对齐: {align_map.get(pf.alignment, pf.alignment)}")
                if pf.line_spacing is not None:
                    print(f"  行距: {pf.line_spacing}")

    print("\n" + "=" * 80)

if __name__ == "__main__":
    template_file = "2026届论文格式模版（最终版）(1).docx"
    detailed_analysis(template_file)
