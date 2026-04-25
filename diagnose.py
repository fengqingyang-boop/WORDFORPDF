#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
诊断脚本 - 用于检查转换环境
"""

import sys
import os
from pathlib import Path

print("=" * 60)
print("Word-PDF 转换器诊断工具")
print("=" * 60)

print("\n【1】检查Python环境")
print(f"    Python版本: {sys.version}")
print(f"    Python路径: {sys.executable}")

print("\n【2】检查依赖库")

try:
    import docx2pdf
    print("    ✓ docx2pdf - 已安装 (用于 Word → PDF)")
    print(f"      版本: {getattr(docx2pdf, '__version__', 'unknown')}")
except ImportError as e:
    print("    ✗ docx2pdf - 未安装")
    print(f"      错误: {e}")

try:
    import pdf2docx
    print("    ✓ pdf2docx - 已安装 (用于 PDF → Word)")
    print(f"      版本: {getattr(pdf2docx, '__version__', 'unknown')}")
except ImportError as e:
    print("    ✗ pdf2docx - 未安装")
    print(f"      错误: {e}")

print("\n【3】检查系统环境")
print(f"    操作系统: {sys.platform}")

if sys.platform == 'win32':
    print("\n【4】检查 Microsoft Word COM 接口 (仅 Windows)")
    try:
        import win32com.client
        print("    ✓ pywin32 - 已安装")
        
        try:
            print("    尝试连接 Microsoft Word...")
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            print("    ✓ Microsoft Word COM 接口可用")
            word.Quit()
            print("\n    ✅ Word 转 PDF 功能应该可以正常工作")
        except Exception as e:
            print(f"    ✗ Microsoft Word COM 接口连接失败")
            print(f"      错误: {e}")
            print("\n    ⚠️  Word 转 PDF 功能可能无法使用")
            print("       可能的原因:")
            print("       1. 系统未安装 Microsoft Word")
            print("       2. Word 正在运行，请关闭所有 Word 窗口")
            print("       3. Word 版本不兼容")
            
    except ImportError:
        print("    ✗ pywin32 - 未安装")
        print("\n    ⚠️  Word 转 PDF 功能需要 pywin32")

print("\n【5】检查工作目录")
work_dir = Path.cwd()
print(f"    当前目录: {work_dir}")
print(f"    目录存在: {work_dir.exists()}")
print(f"    可写权限: {os.access(work_dir, os.W_OK)}")

print("\n【6】项目文件检查")
main_file = work_dir / "main.py"
converter_file = work_dir / "converter.py"
requirements_file = work_dir / "requirements.txt"

print(f"    main.py: {'✓ 存在' if main_file.exists() else '✗ 不存在'}")
print(f"    converter.py: {'✓ 存在' if converter_file.exists() else '✗ 不存在'}")
print(f"    requirements.txt: {'✓ 存在' if requirements_file.exists() else '✗ 不存在'}")

print("\n" + "=" * 60)
print("诊断完成")
print("=" * 60)

print("""
使用建议：

1. 如果 Word 转 PDF 失败：
   - 确保系统已安装 Microsoft Word
   - 关闭所有正在运行的 Word 窗口
   - 检查文件是否被其他程序占用

2. 如果 PDF 转 Word 失败：
   - 检查 PDF 文件是否已加密
   - 检查 PDF 文件是否完整
   - 尝试用其他 PDF 工具重新保存

3. 运行主程序：
   python main.py
   
   或双击运行 main.py 文件

4. 查看详细错误：
   当转换失败时，会弹出详细的错误窗口
""")
