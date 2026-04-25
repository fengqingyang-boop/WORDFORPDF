import os
import sys
import time
import traceback
from pathlib import Path
from typing import Optional, Callable

try:
    from pdf2docx import Converter as PdfToDocxConverter
    HAS_PDF2DOCX = True
except ImportError:
    HAS_PDF2DOCX = False

try:
    import win32com.client
    from win32com.client import constants
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False


class ConversionError(Exception):
    pass


class WordToPdfConverter:
    WD_FORMAT_PDF = 17
    
    @staticmethod
    def _get_word_app():
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            return word
        except Exception as e:
            raise ConversionError(
                "无法启动 Microsoft Word\n"
                "\n可能的原因:\n"
                "  1. 系统未安装 Microsoft Word\n"
                "  2. Word 正在运行中，有未保存的文档\n"
                "  3. Word 安装损坏\n"
                f"\n错误信息: {e}"
            )
    
    @staticmethod
    def _close_word_app(word):
        try:
            word.Quit()
        except:
            pass
    
    @staticmethod
    def convert(
        input_path: str,
        output_path: Optional[str] = None,
        progress_callback: Optional[Callable[[int, int], None]] = None
    ) -> str:
        if not HAS_WIN32COM:
            raise ConversionError(
                "win32com 库未安装\n"
                "\n请运行以下命令安装:\n"
                "  pip install pywin32"
            )
        
        input_path = Path(input_path).resolve()
        
        if not input_path.exists():
            raise ConversionError(f"输入文件不存在: {input_path}")
        
        if input_path.suffix.lower() not in ('.docx', '.doc'):
            raise ConversionError(
                f"不支持的文件格式: {input_path.suffix}\n"
                f"仅支持: .docx, .doc"
            )
        
        if output_path is None:
            output_path = input_path.with_suffix('.pdf')
        else:
            output_path = Path(output_path).resolve()
        
        output_dir = output_path.parent
        if not output_dir.exists():
            output_dir.mkdir(parents=True)
        
        if progress_callback:
            progress_callback(1, 10)
        
        print(f"[调试] 开始转换 Word 到 PDF:")
        print(f"[调试]   输入文件: {input_path}")
        print(f"[调试]   输出文件: {output_path}")
        print(f"[调试]   文件大小: {input_path.stat().st_size} 字节")
        print(f"[调试]   文件格式: {input_path.suffix}")
        
        word = None
        doc = None
        
        try:
            if progress_callback:
                progress_callback(2, 10)
            
            print("[调试]   正在启动 Microsoft Word...")
            word = WordToPdfConverter._get_word_app()
            
            if progress_callback:
                progress_callback(3, 10)
            
            print("[调试]   正在打开文档...")
            print(f"[调试]   文档路径: {str(input_path)}")
            
            doc = word.Documents.Open(
                str(input_path),
                ReadOnly=True,
                Visible=False,
                AddToRecentFiles=False
            )
            
            if progress_callback:
                progress_callback(5, 10)
            
            print(f"[调试]   文档页数: {doc.ComputeStatistics(2)} 页")
            print("[调试]   正在转换为 PDF...")
            
            if output_path.exists():
                print(f"[调试]   输出文件已存在，将被覆盖: {output_path}")
                try:
                    output_path.unlink()
                    print(f"[调试]   已删除旧文件")
                except Exception as e:
                    print(f"[调试]   无法删除旧文件: {e}")
            
            doc.SaveAs(
                str(output_path),
                FileFormat=WordToPdfConverter.WD_FORMAT_PDF
            )
            
            if progress_callback:
                progress_callback(8, 10)
            
            print("[调试]   正在关闭文档...")
            
            try:
                doc.Close(SaveChanges=False)
                doc = None
            except Exception as e:
                print(f"[调试]   关闭文档时警告: {e}")
            
            try:
                WordToPdfConverter._close_word_app(word)
                word = None
            except Exception as e:
                print(f"[调试]   关闭 Word 时警告: {e}")
            
            if not output_path.exists():
                raise ConversionError(
                    "PDF 文件未生成\n"
                    f"\n输入文件: {input_path}\n"
                    f"输出文件: {output_path}\n"
                    "\n可能的原因:\n"
                    "  1. 输出目录权限不足\n"
                    "  2. 磁盘空间不足\n"
                    "  3. Word 转换过程中出错"
                )
            
            if progress_callback:
                progress_callback(10, 10)
            
            print(f"[调试] 转换成功!")
            print(f"[调试]   输出文件大小: {output_path.stat().st_size} 字节")
            
            return str(output_path)
            
        except Exception as e:
            error_details = traceback.format_exc()
            print(f"[调试] 转换错误:\n{error_details}")
            
            if doc is not None:
                try:
                    doc.Close(SaveChanges=False)
                except:
                    pass
            
            if word is not None:
                try:
                    WordToPdfConverter._close_word_app(word)
                except:
                    pass
            
            error_msg = str(e)
            
            if "COM" in error_msg or "win32" in error_msg:
                if "拒绝访问" in error_msg or "access" in error_msg.lower():
                    raise ConversionError(
                        "Word COM 接口访问被拒绝\n"
                        "\n可能的原因:\n"
                        "  1. Word 正在运行且有未保存的文档\n"
                        "  2. Word 进程卡住\n"
                        "  3. 权限不足\n"
                        "\n建议:\n"
                        "  1. 保存并关闭所有 Word 文档\n"
                        "  2. 在任务管理器中结束所有 WINWORD.EXE 进程\n"
                        "  3. 重新尝试转换"
                        f"\n\n技术错误: {error_msg}"
                    )
                else:
                    raise ConversionError(
                        "Word COM 接口错误\n"
                        f"\n输入文件: {input_path}\n"
                        "\n可能的原因:\n"
                        "  1. Word 安装不完整\n"
                        "  2. 文档格式不兼容\n"
                        "  3. 文档已损坏\n"
                        f"\n技术错误: {error_msg}"
                    )
            
            elif "Permission" in error_msg or "权限" in error_msg or "拒绝" in error_msg:
                raise ConversionError(
                    "文件权限错误\n"
                    f"\n请检查:\n"
                    f"  1. 输入文件: {input_path}\n"
                    f"  2. 输出文件: {output_path}\n"
                    "\n可能的原因:\n"
                    "  - 文件被其他程序打开\n"
                    "  - 输出目录权限不足\n"
                    "  - 杀毒软件阻止写入\n"
                    f"\n错误信息: {error_msg}"
                )
            
            else:
                raise ConversionError(
                    "Word 转 PDF 转换失败\n"
                    f"\n输入文件: {input_path}\n"
                    f"输出文件: {output_path}\n"
                    f"\n错误信息: {error_msg}\n"
                    "\n建议:\n"
                    "  1. 关闭所有 Word 窗口\n"
                    "  2. 检查文件是否被其他程序占用\n"
                    "  3. 尝试将文件复制到其他目录再转换"
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
            raise ConversionError(
                "pdf2docx 库未安装\n"
                "\n请运行以下命令安装:\n"
                "  pip install pdf2docx"
            )
        
        input_path = Path(input_path).resolve()
        
        if not input_path.exists():
            raise ConversionError(f"输入文件不存在: {input_path}")
        
        if input_path.suffix.lower() != '.pdf':
            raise ConversionError(
                f"不支持的文件格式: {input_path.suffix}\n"
                f"仅支持: .pdf"
            )
        
        if output_path is None:
            output_path = input_path.with_suffix('.docx')
        else:
            output_path = Path(output_path).resolve()
        
        output_dir = output_path.parent
        if not output_dir.exists():
            output_dir.mkdir(parents=True)
        
        print(f"[调试] 开始转换 PDF 到 Word:")
        print(f"[调试]   输入文件: {input_path}")
        print(f"[调试]   输出文件: {output_path}")
        print(f"[调试]   文件大小: {input_path.stat().st_size} 字节")
        
        try:
            def progress_wrapper(progress):
                if progress_callback:
                    progress_callback(int(progress * 10), 10)
            
            cv = PdfToDocxConverter(str(input_path))
            page_count = len(cv)
            print(f"[调试]   PDF 页数: {page_count}")
            
            if page_count == 0:
                raise ConversionError(
                    "PDF 文件没有页面\n"
                    f"\n文件: {input_path}\n"
                    "\n可能的原因:\n"
                    "  - PDF 文件已损坏\n"
                    "  - PDF 文件是空的"
                )
            
            if end_page is None:
                end_page = page_count - 1
            
            if start_page < 0:
                start_page = 0
            if end_page >= page_count:
                end_page = page_count - 1
            
            print(f"[调试]   转换范围: 第 {start_page + 1} 页到第 {end_page + 1} 页")
            
            if output_path.exists():
                print(f"[调试]   输出文件已存在，将被覆盖: {output_path}")
                try:
                    output_path.unlink()
                except Exception as e:
                    print(f"[调试]   无法删除旧文件: {e}")
            
            cv.convert(
                str(output_path),
                start=start_page,
                end=end_page + 1,
                callback=progress_wrapper
            )
            cv.close()
            
            if not output_path.exists():
                raise ConversionError(
                    "Word 文件未生成\n"
                    f"\n输入文件: {input_path}\n"
                    f"输出文件: {output_path}\n"
                    "\n可能的原因:\n"
                    "  1. PDF 文件已加密\n"
                    "  2. PDF 文件格式不兼容\n"
                    "  3. 输出目录权限不足"
                )
            
            print(f"[调试] 转换成功!")
            print(f"[调试]   输出文件大小: {output_path.stat().st_size} 字节")
            
            return str(output_path)
            
        except Exception as e:
            error_details = traceback.format_exc()
            print(f"[调试] 转换错误:\n{error_details}")
            
            error_msg = str(e)
            
            if "encrypted" in error_msg.lower() or "密码" in error_msg or "加密" in error_msg:
                raise ConversionError(
                    "PDF 文件已加密\n"
                    f"\n文件: {input_path}\n"
                    "\n此 PDF 文件需要密码才能打开。\n"
                    "请先解密 PDF 文件后再尝试转换。"
                    f"\n\n技术错误: {error_msg}"
                )
            
            elif "corrupted" in error_msg.lower() or "损坏" in error_msg or "error" in error_msg.lower():
                raise ConversionError(
                    "PDF 文件可能已损坏\n"
                    f"\n文件: {input_path}\n"
                    "\n建议:\n"
                    "  1. 用 PDF 阅读器打开文件确认完整性\n"
                    "  2. 尝试用其他 PDF 工具重新保存\n"
                    "  3. 检查文件是否下载完整"
                    f"\n\n技术错误: {error_msg}"
                )
            
            elif "Permission" in error_msg or "权限" in error_msg or "拒绝" in error_msg:
                raise ConversionError(
                    "文件权限错误\n"
                    f"\n请检查:\n"
                    f"  1. 输入文件: {input_path}\n"
                    f"  2. 输出文件: {output_path}\n"
                    "\n可能的原因:\n"
                    "  - 文件被其他程序打开\n"
                    "  - 输出目录权限不足"
                    f"\n错误信息: {error_msg}"
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
