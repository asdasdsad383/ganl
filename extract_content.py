#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
提取文档内容和格式
"""
from docx import Document
from docx.oxml.ns import qn

def extract_content(file_path):
    """提取文档主要内容和格式"""
    doc = Document(file_path)

    print("=" * 80)
    print("论文格式要求提取")
    print("=" * 80)

    # 查找关键部分
    in_body = False
    chapter_found = False

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()

        # 跳过空段落
        if not text:
            continue

        # 查找章节标题
        if ('绪论' in text or '第一章' in text or '第1章' in text or
            text.startswith('第') and '章' in text):
            chapter_found = True
            in_body = True
            print(f"\n{'='*80}")
            print(f"【找到章节】段落 {i+1}")
            print(f"内容: {text}")
            print(f"样式: {para.style.name}")

            # 详细格式
            if para.runs:
                run = para.runs[0]
                print(f"字体格式:")
                if run.font.name:
                    print(f"  ASCII字体: {run.font.name}")
                if run._element.rPr is not None:
                    rFonts = run._element.rPr.rFonts
                    if rFonts is not None:
                        east_asia = rFonts.get(qn('w:eastAsia'))
                        if east_asia:
                            print(f"  中文字体: {east_asia}")
                if run.font.size:
                    print(f"  字号: {run.font.size.pt} pt")
                if run.font.bold is not None:
                    print(f"  加粗: {run.font.bold}")

            pf = para.paragraph_format
            print(f"段落格式:")
            if para.alignment is not None:
                align_map = {0: '左对齐', 1: '居中', 2: '右对齐', 3: '两端对齐'}
                print(f"  对齐: {align_map.get(para.alignment, para.alignment)}")
            if pf.line_spacing is not None:
                print(f"  行距: {pf.line_spacing}")
            if pf.space_before and pf.space_before.pt > 0:
                print(f"  段前: {pf.space_before.pt} pt")
            if pf.space_after and pf.space_after.pt > 0:
                print(f"  段后: {pf.space_after.pt} pt")
            if pf.first_line_indent:
                print(f"  首行缩进: {pf.first_line_indent.cm:.2f} cm")

        # 查找二级标题
        elif in_body and (text.startswith('1.') or text.startswith('2.') or
                         '研究背景' in text or '系统设计' in text):
            print(f"\n【二级标题】段落 {i+1}")
            print(f"内容: {text[:50]}")
            print(f"样式: {para.style.name}")

            if para.runs:
                run = para.runs[0]
                if run.font.name:
                    print(f"  ASCII字体: {run.font.name}")
                if run._element.rPr is not None:
                    rFonts = run._element.rPr.rFonts
                    if rFonts is not None:
                        east_asia = rFonts.get(qn('w:eastAsia'))
                        if east_asia:
                            print(f"  中文字体: {east_asia}")
                if run.font.size:
                    print(f"  字号: {run.font.size.pt} pt")
                if run.font.bold is not None:
                    print(f"  加粗: {run.font.bold}")

            pf = para.paragraph_format
            if para.alignment is not None:
                align_map = {0: '左对齐', 1: '居中', 2: '右对齐', 3: '两端对齐'}
                print(f"  对齐: {align_map.get(para.alignment, para.alignment)}")
            if pf.first_line_indent:
                print(f"  首行缩进: {pf.first_line_indent.cm:.2f} cm")

        # 查找正文段落（足够长的段落）
        elif in_body and len(text) > 50 and not text.startswith('图') and not text.startswith('表'):
            if '。' in text or '，' in text:  # 包含中文标点的正文
                print(f"\n【正文段落】段落 {i+1}")
                print(f"内容: {text[:60]}...")
                print(f"样式: {para.style.name}")

                if para.runs:
                    run = para.runs[0]
                    if run.font.name:
                        print(f"  ASCII字体: {run.font.name}")
                    if run._element.rPr is not None:
                        rFonts = run._element.rPr.rFonts
                        if rFonts is not None:
                            east_asia = rFonts.get(qn('w:eastAsia'))
                            if east_asia:
                                print(f"  中文字体: {east_asia}")
                    if run.font.size:
                        print(f"  字号: {run.font.size.pt} pt")

                pf = para.paragraph_format
                if para.alignment is not None:
                    align_map = {0: '左对齐', 1: '居中', 2: '右对齐', 3: '两端对齐'}
                    print(f"  对齐: {align_map.get(para.alignment, para.alignment)}")
                if pf.line_spacing:
                    print(f"  行距: {pf.line_spacing}")
                if pf.first_line_indent:
                    print(f"  首行缩进: {pf.first_line_indent.cm:.2f} cm")

                # 只分析几个正文段落
                break

        # 查找图表标题
        if text.startswith('图') or text.startswith('表'):
            print(f"\n【图表标题】段落 {i+1}")
            print(f"内容: {text}")
            print(f"样式: {para.style.name}")

            if para.runs:
                run = para.runs[0]
                if run.font.size:
                    print(f"  字号: {run.font.size.pt} pt")

            if para.alignment is not None:
                align_map = {0: '左对齐', 1: '居中', 2: '右对齐', 3: '两端对齐'}
                print(f"  对齐: {align_map.get(para.alignment, para.alignment)}")

    print("\n" + "=" * 80)

if __name__ == "__main__":
    template_file = "2026届论文格式模版（最终版）(1).docx"
    extract_content(template_file)
