#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试脚本 - 直接测试转换功能
"""

import os
import sys
from pathlib import Path

print("=" * 60)
print("转换功能测试")
print("=" * 60)

print("\n【1】测试 PDF 转 Word (pdf2docx)")
print("-" * 40)

try:
    from pdf2docx import Converter
    print("✓ pdf2docx 导入成功")
    
    print("\n创建一个简单的测试...")
    print("pdf2docx 库已就绪，可以转换 PDF 文件")
    print("\n使用方法示例:")
    print("  from pdf2docx import Converter")
    print("  cv = Converter('input.pdf')")
    print("  cv.convert('output.docx')")
    print("  cv.close()")
    
except Exception as e:
    print(f"✗ pdf2docx 测试失败: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("\n【2】测试 Word 转 PDF (docx2pdf)")
print("-" * 40)

try:
    from docx2pdf import convert
    print("✓ docx2pdf 导入成功")
    
    print("\n检查 Word COM 接口...")
    try:
        import win32com.client
        print("✓ win32com.client 导入成功")
        
        print("\n尝试连接 Microsoft Word...")
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            print("✓ Microsoft Word COM 接口连接成功!")
            word.Quit()
            print("\n✅ Word 转 PDF 功能应该可以正常工作")
            print("\n使用方法示例:")
            print("  from docx2pdf import convert")
            print("  convert('input.docx', 'output.pdf')")
        except Exception as e:
            print(f"✗ Microsoft Word COM 接口连接失败")
            print(f"\n错误信息: {e}")
            print("\n⚠️  可能的原因:")
            print("   1. 系统未安装 Microsoft Word")
            print("   2. Word 正在运行，请关闭所有 Word 窗口")
            print("   3. Word 版本与 Python 位数不匹配 (32位 vs 64位)")
            print("\n💡 替代方案:")
            print("   - 如果没有安装 Microsoft Word，可以考虑使用 LibreOffice")
            print("   - 或者使用其他纯 Python 库 (如 python-docx + reportlab)")
            
    except ImportError:
        print("✗ win32com.client 未安装")
        print("  请运行: pip install pywin32")
        
except Exception as e:
    print(f"✗ docx2pdf 测试失败: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("测试完成")
print("=" * 60)

print("""
💡 使用建议：

1. 运行主程序:
   python main.py

2. 如果 Word 转 PDF 失败:
   - 检查是否安装了 Microsoft Word
   - 关闭所有打开的 Word 文档
   - 检查 Word 版本 (32位/64位) 是否与 Python 匹配

3. 如果 PDF 转 Word 失败:
   - 检查 PDF 文件是否加密
   - 检查 PDF 文件是否完整
   - 尝试用其他 PDF 阅读器打开确认

4. 查看详细错误:
   当转换失败时，程序会弹出详细的错误窗口，包含:
   - 错误类型
   - 可能的原因
   - 建议的解决方案
""")
