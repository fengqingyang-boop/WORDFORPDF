import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Optional

from converter import UniversalConverter


class ConverterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Word-PDF 转换器")
        self.root.geometry("600x450")
        self.root.resizable(False, False)
        
        self.input_file_path: Optional[str] = None
        self.output_file_path: Optional[str] = None
        self.is_converting = False
        
        self._create_widgets()
    
    def _create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        title_label = ttk.Label(
            main_frame,
            text="Word 与 PDF 格式转换器",
            font=("Microsoft YaHei UI", 18, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        input_frame = ttk.LabelFrame(main_frame, text="输入文件", padding="10")
        input_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.input_path_var = tk.StringVar()
        input_entry = ttk.Entry(
            input_frame,
            textvariable=self.input_path_var,
            state="readonly",
            font=("Microsoft YaHei UI", 10)
        )
        input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        select_button = ttk.Button(
            input_frame,
            text="选择文件",
            command=self._select_input_file
        )
        select_button.pack(side=tk.RIGHT)
        
        info_frame = ttk.LabelFrame(main_frame, text="转换信息", padding="10")
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.conversion_type_var = tk.StringVar(value="请先选择文件")
        conversion_type_label = ttk.Label(
            info_frame,
            textvariable=self.conversion_type_var,
            font=("Microsoft YaHei UI", 11)
        )
        conversion_type_label.pack()
        
        output_frame = ttk.LabelFrame(main_frame, text="输出设置", padding="10")
        output_frame.pack(fill=tk.X, pady=(0, 10))
        
        output_label = ttk.Label(output_frame, text="输出路径:")
        output_label.pack(anchor=tk.W)
        
        self.output_path_var = tk.StringVar(value="(将自动生成)")
        output_entry = ttk.Entry(
            output_frame,
            textvariable=self.output_path_var,
            state="readonly",
            font=("Microsoft YaHei UI", 10)
        )
        output_entry.pack(fill=tk.X, pady=(5, 0))
        
        progress_frame = ttk.LabelFrame(main_frame, text="转换进度", padding="10")
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode="determinate",
            length=500
        )
        self.progress_bar.pack(fill=tk.X)
        
        self.status_var = tk.StringVar(value="就绪")
        status_label = ttk.Label(
            progress_frame,
            textvariable=self.status_var,
            font=("Microsoft YaHei UI", 9)
        )
        status_label.pack(pady=(5, 0))
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.convert_button = ttk.Button(
            button_frame,
            text="开始转换",
            command=self._start_conversion,
            state="disabled"
        )
        self.convert_button.pack(side=tk.LEFT, padx=(0, 10))
        
        open_folder_button = ttk.Button(
            button_frame,
            text="打开输出文件夹",
            command=self._open_output_folder
        )
        open_folder_button.pack(side=tk.LEFT)
        
        hint_frame = ttk.LabelFrame(main_frame, text="使用说明", padding="10")
        hint_frame.pack(fill=tk.X, pady=(10, 0))
        
        hint_text = (
            "• 支持的格式: Word (.docx, .doc) 转 PDF，PDF 转 Word (.docx)\n"
            "• Word 转 PDF 需要系统中安装 Microsoft Word\n"
            "• 输出文件将保存在与输入文件相同的目录下"
        )
        hint_label = ttk.Label(
            hint_frame,
            text=hint_text,
            font=("Microsoft YaHei UI", 9),
            justify=tk.LEFT
        )
        hint_label.pack(anchor=tk.W)
    
    def _select_input_file(self):
        filetypes = [
            ("支持的文件", "*.docx *.doc *.pdf"),
            ("Word 文档", "*.docx *.doc"),
            ("PDF 文件", "*.pdf"),
            ("所有文件", "*.*")
        ]
        
        file_path = filedialog.askopenfilename(
            title="选择要转换的文件",
            filetypes=filetypes
        )
        
        if file_path:
            self.input_file_path = file_path
            self.input_path_var.set(file_path)
            
            input_path = Path(file_path)
            
            if input_path.suffix.lower() in ('.docx', '.doc'):
                self.conversion_type_var.set(f"转换方向: Word → PDF")
                output_ext = '.pdf'
            elif input_path.suffix.lower() == '.pdf':
                self.conversion_type_var.set(f"转换方向: PDF → Word")
                output_ext = '.docx'
            else:
                self.conversion_type_var.set("不支持的文件格式")
                self.convert_button.config(state="disabled")
                return
            
            default_output = input_path.with_suffix(output_ext)
            self.output_file_path = str(default_output)
            self.output_path_var.set(str(default_output))
            
            self.convert_button.config(state="normal")
            self.progress_var.set(0)
            self.status_var.set("就绪 - 点击开始转换")
    
    def _progress_callback(self, current: int, total: int):
        progress_percent = (current / total) * 100
        self.root.after(0, lambda: self.progress_var.set(progress_percent))
        self.root.after(0, lambda: self.status_var.set(f"转换中... {current}/{total}"))
    
    def _conversion_thread(self):
        try:
            self.root.after(0, lambda: self.status_var.set("正在初始化转换..."))
            
            output_path = UniversalConverter.convert(
                self.input_file_path,
                self.output_file_path,
                progress_callback=self._progress_callback
            )
            
            self.root.after(0, lambda: self._on_conversion_success(output_path))
            
        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda: self._on_conversion_error(error_msg))
    
    def _on_conversion_success(self, output_path: str):
        self.progress_var.set(100)
        self.status_var.set("转换完成!")
        self.is_converting = False
        self.convert_button.config(state="normal")
        
        messagebox.showinfo(
            "转换成功",
            f"文件已成功转换!\n输出路径: {output_path}"
        )
    
    def _on_conversion_error(self, error_msg: str):
        self.status_var.set(f"转换失败: {error_msg}")
        self.is_converting = False
        self.convert_button.config(state="normal")
        
        messagebox.showerror(
            "转换失败",
            f"转换过程中发生错误:\n{error_msg}"
        )
    
    def _start_conversion(self):
        if self.is_converting:
            return
        
        if not self.input_file_path:
            messagebox.showwarning("提示", "请先选择要转换的文件")
            return
        
        self.is_converting = True
        self.convert_button.config(state="disabled")
        self.progress_var.set(0)
        self.status_var.set("准备转换...")
        
        thread = threading.Thread(target=self._conversion_thread, daemon=True)
        thread.start()
    
    def _open_output_folder(self):
        if self.output_file_path:
            output_path = Path(self.output_file_path)
            if output_path.parent.exists():
                os.startfile(str(output_path.parent))
            else:
                messagebox.showwarning("提示", "输出目录不存在")
        elif self.input_file_path:
            input_path = Path(self.input_file_path)
            os.startfile(str(input_path.parent))
        else:
            messagebox.showwarning("提示", "请先选择文件")


def main():
    root = tk.Tk()
    app = ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
