#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试格式转换器
"""
import sys
import os

# 添加当前目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from format_converter import PaperFormatConverter

def test_converter():
    """测试格式转换器"""
    print("=" * 60)
    print("测试论文格式转换器")
    print("=" * 60)

    # 创建转换器实例
    converter = PaperFormatConverter()

    # 获取格式摘要
    print("\n【格式要求摘要】")
    summary = converter.get_format_summary()
    for key, value in summary.items():
        if isinstance(value, dict):
            print(f"\n{key}:")
            for k, v in value.items():
                print(f"  {k}: {v}")
        else:
            print(f"{key}: {value}")

    print("\n" + "=" * 60)
    print("✅ 格式转换器测试通过！")
    print("=" * 60)

    # 测试文件转换（如果有模板文件）
    template_file = "../2026届论文格式模版（最终版）(1).docx"
    if os.path.exists(template_file):
        print("\n正在测试格式转换...")
        output_file = "outputs/test_output.docx"
        success, message = converter.convert(template_file, output_file)

        if success:
            print(f"✅ {message}")
        else:
            print(f"❌ {message}")
    else:
        print("\n⚠️  未找到模板文件，跳过转换测试")

if __name__ == "__main__":
    test_converter()
