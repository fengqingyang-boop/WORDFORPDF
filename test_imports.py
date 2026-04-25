print("测试导入转换器模块...")

try:
    from converter import UniversalConverter, WordToPdfConverter, PdfToWordConverter
    print("✓ converter 模块导入成功")
except ImportError as e:
    print(f"✗ converter 模块导入失败: {e}")
    raise

print("\n测试导入主程序...")
try:
    import main
    print("✓ main 模块导入成功")
except ImportError as e:
    print(f"✗ main 模块导入失败: {e}")
    raise

print("\n所有导入测试通过!")
print("\n项目结构:")
print("  - converter.py: 核心转换模块")
print("  - main.py: GUI界面程序")
print("  - requirements.txt: 依赖文件")
print("\n使用方法:")
print("  1. 双击运行 main.py 启动图形界面")
print("  2. 或在命令行执行: python main.py")
print("  3. 点击 '选择文件' 按钮选择要转换的文件")
print("  4. 点击 '开始转换' 进行转换")
print("  5. 转换完成后可通过 '打开输出文件夹' 查看结果")
print("\n注意事项:")
print("  - Word 转 PDF 需要系统中安装 Microsoft Word")
print("  - 支持的格式: .docx, .doc → .pdf；.pdf → .docx")
