import os
import traceback
from pathlib import Path
from typing import Optional, Callable

try:
    from docx2pdf import convert as docx_to_pdf_convert
    HAS_DOCX2PDF = True
except ImportError:
    HAS_DOCX2PDF = False

try:
    from pdf2docx import Converter as PdfToDocxConverter
    HAS_PDF2DOCX = True
except ImportError:
    HAS_PDF2DOCX = False


class ConversionError(Exception):
    pass


class WordToPdfConverter:
    @staticmethod
    def convert(
        input_path: str,
        output_path: Optional[str] = None,
        progress_callback: Optional[Callable[[int, int], None]] = None
    ) -> str:
        if not HAS_DOCX2PDF:
            raise ConversionError("docx2pdf 库未安装，请运行: pip install docx2pdf")
        
        input_path = Path(input_path).resolve()
        
        if not input_path.exists():
            raise ConversionError(f"输入文件不存在: {input_path}")
        
        if input_path.suffix.lower() not in ('.docx', '.doc'):
            raise ConversionError(f"不支持的文件格式: {input_path.suffix}，仅支持 .docx 和 .doc")
        
        if output_path is None:
            output_path = input_path.with_suffix('.pdf')
        else:
            output_path = Path(output_path).resolve()
        
        output_dir = output_path.parent
        if not output_dir.exists():
            output_dir.mkdir(parents=True)
        
        if progress_callback:
            progress_callback(1, 10)
        
        try:
            print(f"[调试] 开始转换 Word 到 PDF:")
            print(f"[调试]   输入文件: {input_path}")
            print(f"[调试]   输出文件: {output_path}")
            print(f"[调试]   文件大小: {input_path.stat().st_size} 字节")
            
            docx_to_pdf_convert(str(input_path), str(output_path))
            
            if not output_path.exists():
                raise ConversionError(
                    "转换失败: 输出文件未生成\n"
                    "可能的原因:\n"
                    "  1. 系统未安装 Microsoft Word\n"
                    "  2. Microsoft Word 正在运行且占用文件\n"
                    "  3. 文件权限问题\n"
                    f"\n输入文件: {input_path}\n"
                    f"输出文件: {output_path}"
                )
            
            if progress_callback:
                progress_callback(10, 10)
            
            print(f"[调试] 转换成功! 输出文件大小: {output_path.stat().st_size} 字节")
            return str(output_path)
            
        except Exception as e:
            error_details = traceback.format_exc()
            print(f"[调试] 转换错误:\n{error_details}")
            
            error_msg = str(e)
            
            if "COM" in error_msg or "win32" in error_msg:
                raise ConversionError(
                    "Word 转 PDF 失败 (COM 接口错误)\n"
                    "\n可能的原因:\n"
                    "  1. 系统未安装 Microsoft Word\n"
                    "  2. Microsoft Word 版本不兼容\n"
                    "  3. Word 正在运行，请关闭所有 Word 窗口后重试\n"
                    "\n技术详情:\n"
                    f"{error_msg}"
                )
            elif "Permission" in error_msg or "权限" in error_msg:
                raise ConversionError(
                    "文件权限错误\n"
                    f"\n请检查文件是否被其他程序占用:\n"
                    f"  输入: {input_path}\n"
                    f"  输出: {output_path}\n"
                    "\n建议:\n"
                    "  1. 关闭所有打开的 Word 文档\n"
                    "  2. 确保输出目录可写\n"
                    f"\n错误详情: {error_msg}"
                )
            else:
                raise ConversionError(
                    "Word 转 PDF 转换失败\n"
                    f"\n输入文件: {input_path}\n"
                    f"输出文件: {output_path}\n"
                    f"\n错误信息: {error_msg}\n"
                    "\n建议:\n"
                    "  1. 确保已安装 Microsoft Word\n"
                    "  2. 关闭所有 Word 窗口\n"
                    "  3. 检查文件是否损坏"
                )


class PdfToWordConverter:
    @staticmethod
    def convert(
        input_path: str,
        output_path: Optional[str] = None,
        progress_callback: Optional[Callable[[int, int], None]] = None,
        start_page: int = 0,
        end_page: Optional[int] = None
    ) -> str:
        if not HAS_PDF2DOCX:
            raise ConversionError("pdf2docx 库未安装，请运行: pip install pdf2docx")
        
        input_path = Path(input_path).resolve()
        
        if not input_path.exists():
            raise ConversionError(f"输入文件不存在: {input_path}")
        
        if input_path.suffix.lower() != '.pdf':
            raise ConversionError(f"不支持的文件格式: {input_path.suffix}，仅支持 .pdf")
        
        if output_path is None:
            output_path = input_path.with_suffix('.docx')
        else:
            output_path = Path(output_path).resolve()
        
        output_dir = output_path.parent
        if not output_dir.exists():
            output_dir.mkdir(parents=True)
        
        try:
            print(f"[调试] 开始转换 PDF 到 Word:")
            print(f"[调试]   输入文件: {input_path}")
            print(f"[调试]   输出文件: {output_path}")
            print(f"[调试]   文件大小: {input_path.stat().st_size} 字节")
            
            def progress_wrapper(progress):
                if progress_callback:
                    progress_callback(int(progress * 10), 10)
            
            cv = PdfToDocxConverter(str(input_path))
            page_count = len(cv)
            print(f"[调试]   PDF 页数: {page_count}")
            
            if end_page is None:
                end_page = page_count - 1
            
            if start_page < 0:
                start_page = 0
            if end_page >= page_count:
                end_page = page_count - 1
            
            print(f"[调试]   转换范围: 第 {start_page + 1} 页到第 {end_page + 1} 页")
            
            cv.convert(
                str(output_path),
                start=start_page,
                end=end_page + 1,
                callback=progress_wrapper
            )
            cv.close()
            
            if not output_path.exists():
                raise ConversionError(
                    "PDF 转 Word 失败: 输出文件未生成\n"
                    f"\n输入文件: {input_path}\n"
                    f"输出文件: {output_path}\n"
                    "\n可能的原因:\n"
                    "  1. PDF 文件被加密或损坏\n"
                    "  2. PDF 文件包含特殊格式无法解析\n"
                    "  3. 输出目录权限不足"
                )
            
            print(f"[调试] 转换成功! 输出文件大小: {output_path.stat().st_size} 字节")
            return str(output_path)
            
        except Exception as e:
            error_details = traceback.format_exc()
            print(f"[调试] 转换错误:\n{error_details}")
            
            error_msg = str(e)
            
            if "encrypted" in error_msg.lower() or "密码" in error_msg or "加密" in error_msg:
                raise ConversionError(
                    "PDF 文件已加密\n"
                    f"\n文件: {input_path}\n"
                    "\n此 PDF 文件需要密码才能打开，请先解密后再尝试转换。"
                )
            elif "corrupted" in error_msg.lower() or "损坏" in error_msg:
                raise ConversionError(
                    "PDF 文件已损坏\n"
                    f"\n文件: {input_path}\n"
                    "\n文件可能已损坏，请尝试使用其他 PDF 查看器打开确认文件完整性。"
                )
            elif "Permission" in error_msg or "权限" in error_msg:
                raise ConversionError(
                    "文件权限错误\n"
                    f"\n请检查文件是否被其他程序占用:\n"
                    f"  输入: {input_path}\n"
                    f"  输出: {output_path}\n"
                    f"\n错误详情: {error_msg}"
                )
            else:
                raise ConversionError(
                    "PDF 转 Word 转换失败\n"
                    f"\n输入文件: {input_path}\n"
                    f"输出文件: {output_path}\n"
                    f"\n错误信息: {error_msg}\n"
                    "\n建议:\n"
                    "  1. 检查 PDF 文件是否完整\n"
                    "  2. 确保 PDF 未加密\n"
                    "  3. 尝试用其他 PDF 工具重新保存后再转换"
                )


class UniversalConverter:
    @staticmethod
    def convert(
        input_path: str,
        output_path: Optional[str] = None,
        progress_callback: Optional[Callable[[int, int], None]] = None
    ) -> str:
        input_path = Path(input_path).resolve()
        
        if not input_path.exists():
            raise ConversionError(f"文件不存在: {input_path}")
        
        suffix = input_path.suffix.lower()
        
        if suffix in ('.docx', '.doc'):
            return WordToPdfConverter.convert(
                str(input_path),
                output_path,
                progress_callback
            )
        elif suffix == '.pdf':
            return PdfToWordConverter.convert(
                str(input_path),
                output_path,
                progress_callback
            )
        else:
            raise ConversionError(
                f"不支持的文件格式: {suffix}\n"
                f"仅支持: .docx, .doc, .pdf\n"
                f"\n文件: {input_path}"
            )
