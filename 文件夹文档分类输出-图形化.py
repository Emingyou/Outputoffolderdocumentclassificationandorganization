import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import time
from datetime import datetime

class FileCopyTool:
    def __init__(self, root):
        self.root = root
        self.root.title("文件格式筛选复制工具")
        self.root.geometry("750x500")
        self.root.minsize(700, 450)
        
        # 确保中文显示正常
        self.style = ttk.Style()
        self.style.configure("TLabel", font=("SimHei", 10))
        self.style.configure("TButton", font=("SimHei", 10))
        self.style.configure("TCheckbutton", font=("SimHei", 10))
        
        # 状态变量
        self.copying = False
        self.stop_requested = False
        self.recent_folders = []
        self.load_recent_folders()
        
        self.create_widgets()
    
    def load_recent_folders(self):
        """加载最近使用的文件夹"""
        try:
            if os.path.exists("recent_folders.txt"):
                with open("recent_folders.txt", "r", encoding="utf-8") as f:
                    self.recent_folders = [line.strip() for line in f.readlines()[:10]]
        except Exception as e:
            print(f"加载最近文件夹失败: {e}")
    
    def save_recent_folder(self, folder):
        """保存最近使用的文件夹"""
        if folder and folder not in self.recent_folders:
            self.recent_folders.insert(0, folder)
            self.recent_folders = self.recent_folders[:10]  # 只保留最近10个
            try:
                with open("recent_folders.txt", "w", encoding="utf-8") as f:
                    f.write("\n".join(self.recent_folders))
            except Exception as e:
                print(f"保存最近文件夹失败: {e}")
    
    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 源文件夹选择
        ttk.Label(main_frame, text="源文件夹:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.source_var = tk.StringVar()
        self.source_entry = ttk.Entry(main_frame, textvariable=self.source_var, width=50)
        self.source_entry.grid(row=0, column=1, sticky=tk.EW, pady=5)
        ttk.Button(main_frame, text="浏览...", command=self.select_source_dir).grid(row=0, column=2, padx=5, pady=5)
        
        # 目标文件夹选择
        ttk.Label(main_frame, text="目标文件夹:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.dest_var = tk.StringVar()
        self.dest_entry = ttk.Entry(main_frame, textvariable=self.dest_var, width=50)
        self.dest_entry.grid(row=1, column=1, sticky=tk.EW, pady=5)
        ttk.Button(main_frame, text="浏览...", command=self.select_dest_dir).grid(row=1, column=2, padx=5, pady=5)
        
        # 最近文件夹选择
        if self.recent_folders:
            ttk.Label(main_frame, text="最近文件夹:").grid(row=2, column=0, sticky=tk.W, pady=5)
            recent_frame = ttk.Frame(main_frame)
            recent_frame.grid(row=2, column=1, columnspan=2, sticky=tk.EW, pady=5)
            
            for i, folder in enumerate(self.recent_folders[:5]):  # 只显示最近5个
                btn = ttk.Button(
                    recent_frame, 
                    text=os.path.basename(folder) if folder else "空",
                    command=lambda f=folder: self.use_recent_folder(f)
                )
                btn.pack(side=tk.LEFT, padx=2)
        
        # 选项设置框架 - 筛选选项部分修改
        options_frame = ttk.LabelFrame(main_frame, text="筛选选项", padding="5")
        options_frame.grid(row=3, column=0, columnspan=3, sticky=tk.EW, pady=5)
        
        # 筛选模式选择
        ttk.Label(options_frame, text="筛选模式:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.filter_mode = tk.StringVar(value="include")
        ttk.Radiobutton(options_frame, text="包含指定格式", variable=self.filter_mode, value="include").grid(row=0, column=1, sticky=tk.W)
        ttk.Radiobutton(options_frame, text="排除指定格式", variable=self.filter_mode, value="exclude").grid(row=0, column=2, sticky=tk.W)
        
        # 文件格式输入（支持多个格式，用逗号分隔）
        ttk.Label(options_frame, text="文件格式:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.extension_var = tk.StringVar()
        self.extension_entry = ttk.Entry(options_frame, textvariable=self.extension_var, width=30)
        self.extension_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(options_frame, text="(例如: txt,docx 或 .txt,.docx，多个格式用逗号分隔)").grid(row=1, column=2, columnspan=2, sticky=tk.W, pady=5)
        
        # 常用格式快捷按钮
        common_exts = [".txt", ".docx", ".xlsx", ".pdf", ".jpg", ".png", ".csv", ".zip"]
        ext_frame = ttk.Frame(options_frame)
        ext_frame.grid(row=2, column=0, columnspan=4, sticky=tk.W, pady=5)
        for ext in common_exts:
            ttk.Button(
                ext_frame, 
                text=ext, 
                command=lambda e=ext: self.add_extension(e)
            ).pack(side=tk.LEFT, padx=2)
        
        # 高级选项
        self.recursive_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            options_frame, 
            text="包含子文件夹", 
            variable=self.recursive_var
        ).grid(row=3, column=0, sticky=tk.W, padx=5, pady=2)
        
        self.overwrite_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            options_frame, 
            text="覆盖已存在的文件", 
            variable=self.overwrite_var
        ).grid(row=3, column=1, sticky=tk.W, padx=5, pady=2)
        
        # 文件名包含文本筛选
        ttk.Label(options_frame, text="文件名包含:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        self.filename_contains = tk.StringVar()
        ttk.Entry(options_frame, textvariable=self.filename_contains, width=30).grid(row=4, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(options_frame, text="(空则不按文件名筛选)").grid(row=4, column=2, sticky=tk.W, pady=5)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=4, column=0, columnspan=3, sticky=tk.EW, pady=5)
        
        # 状态标签
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(main_frame, textvariable=self.status_var).grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=2)
        
        # 操作按钮
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=6, column=0, columnspan=3, pady=10)
        
        self.start_btn = ttk.Button(btn_frame, text="开始复制", command=self.start_copy, width=15)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_btn = ttk.Button(btn_frame, text="停止", command=self.stop_copy, width=15, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        self.clear_btn = ttk.Button(btn_frame, text="清空日志", command=self.clear_log, width=15)
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        
        # 日志区域
        ttk.Label(main_frame, text="操作日志:").grid(row=7, column=0, sticky=tk.NW, pady=5)
        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=7, column=1, columnspan=2, sticky=tk.NSEW, pady=5)
        
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # 配置列权重，使其可以自适应窗口大小
        main_frame.columnconfigure(1, weight=1)
        options_frame.columnconfigure(2, weight=1)
        main_frame.rowconfigure(7, weight=1)
    
    def add_extension(self, ext):
        """添加格式到输入框，支持多个格式"""
        current = self.extension_var.get().strip()
        if current:
            # 检查是否已包含该格式
            exts = [e.strip() for e in current.split(',')]
            if ext not in exts:
                self.extension_var.set(f"{current},{ext}")
        else:
            self.extension_var.set(ext)
    
    def select_source_dir(self):
        """选择源文件夹"""
        dir_path = filedialog.askdirectory(title="选择源文件夹")
        if dir_path:
            self.source_var.set(dir_path)
            self.save_recent_folder(dir_path)
    
    def select_dest_dir(self):
        """选择目标文件夹"""
        dir_path = filedialog.askdirectory(title="选择目标文件夹")
        if dir_path:
            self.dest_var.set(dir_path)
            self.save_recent_folder(dir_path)
    
    def use_recent_folder(self, folder):
        """使用最近的文件夹"""
        # 如果源文件夹为空，优先设置源文件夹
        if not self.source_var.get():
            self.source_var.set(folder)
        else:
            self.dest_var.set(folder)
    
    def log(self, message):
        """添加日志信息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)  # 滚动到最后
    
    def clear_log(self):
        """清空日志"""
        self.log_text.delete(1.0, tk.END)
        self.log("日志已清空")
    
    def stop_copy(self):
        """停止复制操作"""
        if self.copying:
            self.stop_requested = True
            self.status_var.set("正在停止...")
    
    def update_progress(self, value):
        """更新进度条"""
        self.progress_var.set(value)
        self.root.update_idletasks()
    
    def copy_files_by_extension(self, source_dir, dest_dir, extensions_str):
        """复制文件的核心函数（支持多格式筛选和文件名筛选）"""
        # 处理文件格式
        extensions = []
        if extensions_str:
            # 分割并规范化格式
            extensions = [ext.strip().lower() for ext in extensions_str.split(',') if ext.strip()]
            # 确保格式以.开头
            extensions = [ext if ext.startswith('.') else f'.{ext}' for ext in extensions]
        
        # 处理文件名包含文本
        filename_filter = self.filename_contains.get().strip().lower()
        
        # 规范化路径
        source_dir = os.path.abspath(source_dir)
        dest_dir = os.path.abspath(dest_dir)
        
        # 检查源文件夹是否存在
        if not os.path.isdir(source_dir):
            return False, f"错误: 源文件夹 '{source_dir}' 不存在"
        
        # 检查源目录和目标目录是否相同
        if os.path.samefile(source_dir, dest_dir):
            return False, "错误: 源文件夹和目标文件夹不能相同"
        
        # 创建目标文件夹（如果不存在）
        try:
            os.makedirs(dest_dir, exist_ok=True)
        except OSError as e:
            return False, f"错误: 创建目标文件夹失败 - {e}"
        
        # 收集所有要复制的文件
        files_to_copy = []
        if self.recursive_var.get():
            # 递归遍历所有子文件夹
            for root_dir, _, filenames in os.walk(source_dir):
                for filename in filenames:
                    file_path = os.path.join(root_dir, filename)
                    # 检查文件是否符合筛选条件
                    if self.filter_file(filename, extensions, filename_filter):
                        # 计算相对路径，保持目录结构
                        rel_path = os.path.relpath(root_dir, source_dir)
                        files_to_copy.append((file_path, rel_path, filename))
        else:
            # 只处理当前文件夹
            for filename in os.listdir(source_dir):
                file_path = os.path.join(source_dir, filename)
                if os.path.isfile(file_path):
                    if self.filter_file(filename, extensions, filename_filter):
                        files_to_copy.append((file_path, "", filename))
        
        if not files_to_copy:
            return True, "没有找到符合条件的文件"
        
        # 开始复制文件
        copied_count = 0
        skipped_count = 0
        failed_count = 0
        total_files = len(files_to_copy)
        
        for i, (source_path, rel_path, filename) in enumerate(files_to_copy, 1):
            # 检查是否需要停止
            if self.stop_requested:
                return True, f"操作已停止。已复制 {copied_count} 个文件，跳过 {skipped_count} 个，失败 {failed_count} 个"
            
            # 更新进度
            progress = (i / total_files) * 100
            self.root.after(10, self.update_progress, progress)
            self.root.after(10, self.status_var.set, f"正在复制 {i}/{total_files} 文件")
            
            # 构建目标路径
            dest_subdir = os.path.join(dest_dir, rel_path)
            os.makedirs(dest_subdir, exist_ok=True)
            dest_path = os.path.join(dest_subdir, filename)
            
            # 检查文件是否已存在
            if os.path.exists(dest_path) and not self.overwrite_var.get():
                self.root.after(10, self.log, f"跳过: {filename} 已存在")
                skipped_count += 1
                continue
            
            try:
                # 复制文件（保留元数据）
                shutil.copy2(source_path, dest_path)
                self.root.after(10, self.log, f"已复制: {filename}")
                copied_count += 1
            except Exception as e:
                self.root.after(10, self.log, f"复制失败: {filename} - {str(e)}")
                failed_count += 1
        
        return True, (f"操作完成。共处理 {total_files} 个文件\n"
                      f"已复制: {copied_count} 个\n"
                      f"已跳过: {skipped_count} 个\n"
                      f"复制失败: {failed_count} 个")
    
    def filter_file(self, filename, extensions, filename_filter):
        """判断文件是否符合筛选条件"""
        filename_lower = filename.lower()
        
        # 文件名包含文本筛选
        if filename_filter and filename_filter not in filename_lower:
            return False
            
        # 格式筛选
        if not extensions:
            return True  # 没有指定格式，全部匹配
            
        # 检查是否符合包含/排除模式
        if self.filter_mode.get() == "include":
            # 包含模式：文件格式在列表中
            return any(filename_lower.endswith(ext) for ext in extensions)
        else:
            # 排除模式：文件格式不在列表中
            return not any(filename_lower.endswith(ext) for ext in extensions)
    
    def start_copy(self):
        """开始复制文件（在新线程中执行）"""
        source_dir = self.source_var.get().strip()
        dest_dir = self.dest_var.get().strip()
        extensions_str = self.extension_var.get().strip()
        
        # 验证输入
        if not source_dir:
            messagebox.showerror("输入错误", "请选择源文件夹")
            return
        if not dest_dir:
            messagebox.showerror("输入错误", "请选择目标文件夹")
            return
        
        # 禁用开始按钮，启用停止按钮
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.copying = True
        self.stop_requested = False
        self.progress_var.set(0)
        self.log("开始复制文件...")
        
        # 在新线程中执行复制操作，避免界面冻结
        def copy_thread():
            success, result = self.copy_files_by_extension(source_dir, dest_dir, extensions_str)
            
            # 恢复按钮状态
            self.root.after(10, self.start_btn.config, {"state": tk.NORMAL})
            self.root.after(10, self.stop_btn.config, {"state": tk.DISABLED})
            self.root.after(10, lambda: setattr(self, 'copying', False))
            self.root.after(10, self.status_var.set, "就绪")
            self.root.after(10, self.log, result)
            
            # 显示最终结果
            if success:
                self.root.after(10, messagebox.showinfo, "操作完成", result)
            else:
                self.root.after(10, messagebox.showerror, "操作失败", result)
        
        # 启动复制线程
        threading.Thread(target=copy_thread, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = FileCopyTool(root)
    root.mainloop()
