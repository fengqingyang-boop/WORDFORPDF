import os
from pathlib import Path
from typing import Optional, Callable

from docx2pdf import convert as docx_to_pdf_convert
from pdf2docx import Converter as PdfToDocxConverter


class WordToPdfConverter:
    @staticmethod
    def convert(
        input_path: str,
        output_path: Optional[str] = None,
        progress_callback: Optional[Callable[[int, int], None]] = None
    ) -> str:
        input_path = Path(input_path).resolve()
        
        if not input_path.exists():
            raise FileNotFoundError(f"输入文件不存在: {input_path}")
        
        if not input_path.suffix.lower() in ('.docx', '.doc'):
            raise ValueError(f"不支持的文件格式: {input_path.suffix}，仅支持 .docx 和 .doc")
        
        if output_path is None:
            output_path = input_path.with_suffix('.pdf')
        else:
            output_path = Path(output_path).resolve()
        
        output_dir = output_path.parent
        if not output_dir.exists():
            output_dir.mkdir(parents=True)
        
        if progress_callback:
            progress_callback(1, 10)
        
        docx_to_pdf_convert(str(input_path), str(output_path))
        
        if progress_callback:
            progress_callback(10, 10)
        
        return str(output_path)


class PdfToWordConverter:
    @staticmethod
    def convert(
        input_path: str,
        output_path: Optional[str] = None,
        progress_callback: Optional[Callable[[int, int], None]] = None,
        start_page: int = 0,
        end_page: Optional[int] = None
    ) -> str:
        input_path = Path(input_path).resolve()
        
        if not input_path.exists():
            raise FileNotFoundError(f"输入文件不存在: {input_path}")
        
        if not input_path.suffix.lower() == '.pdf':
            raise ValueError(f"不支持的文件格式: {input_path.suffix}，仅支持 .pdf")
        
        if output_path is None:
            output_path = input_path.with_suffix('.docx')
        else:
            output_path = Path(output_path).resolve()
        
        output_dir = output_path.parent
        if not output_dir.exists():
            output_dir.mkdir(parents=True)
        
        def progress_wrapper(progress):
            if progress_callback:
                progress_callback(int(progress * 10), 10)
        
        cv = PdfToDocxConverter(str(input_path))
        
        if end_page is None:
            end_page = len(cv) - 1
        
        cv.convert(
            str(output_path),
            start=start_page,
            end=end_page + 1,
            callback=progress_wrapper
        )
        cv.close()
        
        return str(output_path)


class UniversalConverter:
    @staticmethod
    def convert(
        input_path: str,
        output_path: Optional[str] = None,
        progress_callback: Optional[Callable[[int, int], None]] = None
    ) -> str:
        input_path = Path(input_path).resolve()
        
        if input_path.suffix.lower() in ('.docx', '.doc'):
            return WordToPdfConverter.convert(
                str(input_path),
                output_path,
                progress_callback
            )
        elif input_path.suffix.lower() == '.pdf':
            return PdfToWordConverter.convert(
                str(input_path),
                output_path,
                progress_callback
            )
        else:
            raise ValueError(
                f"不支持的文件格式: {input_path.suffix}，"
                f"仅支持 .docx, .doc, .pdf"
            )
