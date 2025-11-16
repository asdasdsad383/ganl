#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
完整提取论文格式规范
"""
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

def analyze_all_formats(file_path):
    """分析所有格式"""
    doc = Document(file_path)

    formats = {
        '封面': [],
        '一级标题': [],
        '二级标题': [],
        '三级标题': [],
        '正文': [],
        '图表标题': [],
        '参考文献': [],
        '页眉页脚': []
    }

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text or len(text) < 2:
            continue

        style = para.style.name
        pf = para.paragraph_format

        # 提取格式信息
        fmt = {
            '文本': text[:50],
            '样式': style,
            '对齐': para.alignment,
            '行距': pf.line_spacing,
            '段前': pf.space_before.pt if pf.space_before else 0,
            '段后': pf.space_after.pt if pf.space_after else 0,
            '首行缩进': pf.first_line_indent.cm if pf.first_line_indent else 0,
        }

        if para.runs:
            run = para.runs[0]
            fmt['ASCII字体'] = run.font.name
            fmt['字号'] = run.font.size.pt if run.font.size else None
            fmt['加粗'] = run.font.bold

            # 中文字体
            if run._element.rPr is not None:
                rFonts = run._element.rPr.rFonts
                if rFonts is not None:
                    east_asia = rFonts.get(qn('w:eastAsia'))
                    if east_asia:
                        fmt['中文字体'] = east_asia

        # 分类
        if '本科毕业' in text and len(text) < 20:
            formats['封面'].append(fmt)
        elif style == 'Heading 1' or ('第' in text and '章' in text and len(text) < 10) or text == '绪论':
            formats['一级标题'].append(fmt)
        elif style == 'Heading 2' or (text.startswith(('1.', '2.', '3.', '4.', '5.')) and len(text) < 30):
            formats['二级标题'].append(fmt)
        elif style == 'Heading 3':
            formats['三级标题'].append(fmt)
        elif text.startswith(('图', '表')) and len(text) < 50:
            formats['图表标题'].append(fmt)
        elif '参考文献' in text and len(text) < 10:
            formats['参考文献'].append(fmt)
        elif style == 'Normal' and len(text) > 20 and ('。' in text or '，' in text):
            if len(formats['正文']) < 5:  # 只保留前5个正文样例
                formats['正文'].append(fmt)

    # 打印结果
    print("=" * 80)
    print("论文完整格式规范")
    print("=" * 80)

    for category, items in formats.items():
        if items:
            print(f"\n【{category}】")
            for i, fmt in enumerate(items[:3], 1):  # 每类最多显示3个样例
                print(f"\n样例 {i}: {fmt['文本'][:40]}")
                print(f"  样式名称: {fmt['样式']}")
                if 'ASCII字体' in fmt and fmt['ASCII字体']:
                    print(f"  ASCII字体: {fmt['ASCII字体']}")
                if '中文字体' in fmt:
                    print(f"  中文字体: {fmt['中文字体']}")
                if fmt['字号']:
                    print(f"  字号: {fmt['字号']} pt")
                if fmt['加粗']:
                    print(f"  加粗: 是")

                align_map = {None: '默认', 0: '左对齐', 1: '居中', 2: '右对齐', 3: '两端对齐', 4: '分散对齐'}
                print(f"  对齐方式: {align_map.get(fmt['对齐'], fmt['对齐'])}")

                if fmt['行距']:
                    if isinstance(fmt['行距'], (int, float)) and fmt['行距'] > 10:
                        print(f"  行距: {fmt['行距']} (固定值)")
                    else:
                        print(f"  行距: {fmt['行距']} 倍")

                if fmt['段前'] > 0:
                    print(f"  段前间距: {fmt['段前']:.1f} pt")
                if fmt['段后'] > 0:
                    print(f"  段后间距: {fmt['段后']:.1f} pt")
                if fmt['首行缩进'] != 0:
                    print(f"  首行缩进: {fmt['首行缩进']:.2f} cm")

    # 页面设置
    section = doc.sections[0]
    print(f"\n{'='*80}")
    print("【页面设置】")
    print(f"  纸张: A4 ({section.page_width.cm:.1f} x {section.page_height.cm:.1f} cm)")
    print(f"  上边距: {section.top_margin.cm:.1f} cm")
    print(f"  下边距: {section.bottom_margin.cm:.1f} cm")
    print(f"  左边距: {section.left_margin.cm:.1f} cm")
    print(f"  右边距: {section.right_margin.cm:.1f} cm")
    if section.footer_distance.cm > 0:
        print(f"  页脚距边界: {section.footer_distance.cm:.1f} cm")

    print("\n" + "=" * 80)

if __name__ == "__main__":
    template_file = "2026届论文格式模版（最终版）(1).docx"
    analyze_all_formats(template_file)
