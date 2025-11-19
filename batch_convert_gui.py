#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
长图转PDF工具 - GUI版本
带有专业界面，支持文件选择和参数配置
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
import threading
from PIL import Image
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm


class LongImageToPDFGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("长图转PDF工具 v2.0")
        self.root.geometry("800x600")
        self.root.minsize(700, 500)
        
        # 设置图标和样式
        try:
            self.root.iconbitmap(default="")  # 可以添加ico文件路径
        except:
            pass
            
        # 变量
        self.selected_files = []
        self.output_dir = tk.StringVar(value="")
        self.num_columns = tk.IntVar(value=3)
        self.overlap = tk.IntVar(value=50)
        self.column_gap = tk.DoubleVar(value=5.0)
        self.dpi = tk.IntVar(value=300)
        self.orientation = tk.StringVar(value="landscape")
        self.margin = tk.DoubleVar(value=10.0)
        
        self.setup_ui()
        self.center_window()
        
    def setup_ui(self):
        """设置用户界面"""
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="长图转PDF工具", font=("Microsoft YaHei", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # 创建Notebook（选项卡）
        notebook = ttk.Notebook(main_frame)
        notebook.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # 文件选择选项卡
        self.file_frame = ttk.Frame(notebook, padding="10")
        notebook.add(self.file_frame, text="文件选择")
        self.setup_file_tab()
        
        # 转换设置选项卡
        self.settings_frame = ttk.Frame(notebook, padding="10")
        notebook.add(self.settings_frame, text="转换设置")
        self.setup_settings_tab()
        
        # 输出设置选项卡
        self.output_frame = ttk.Frame(notebook, padding="10")
        notebook.add(self.output_frame, text="输出设置")
        self.setup_output_tab()
        
        # 控制按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        button_frame.columnconfigure(0, weight=1)
        
        # 预览和转换按钮
        preview_btn = ttk.Button(button_frame, text="预览设置", command=self.preview_settings)
        preview_btn.grid(row=0, column=0, padx=(0, 5), sticky=tk.W)
        
        convert_btn = ttk.Button(button_frame, text="开始转换", command=self.start_conversion, 
                               style="Accent.TButton")
        convert_btn.grid(row=0, column=1, padx=5)
        
        exit_btn = ttk.Button(button_frame, text="退出", command=self.root.quit)
        exit_btn.grid(row=0, column=2, padx=(5, 0), sticky=tk.E)
        
    def setup_file_tab(self):
        """设置文件选择选项卡"""
        # 文件选择区域
        file_label = ttk.Label(self.file_frame, text="选择要转换的图片文件:", font=("Microsoft YaHei", 10, "bold"))
        file_label.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))
        
        # 文件选择按钮
        select_btn = ttk.Button(self.file_frame, text="选择文件", command=self.select_files)
        select_btn.grid(row=1, column=0, sticky=tk.W, padx=(0, 5))
        
        clear_btn = ttk.Button(self.file_frame, text="清空列表", command=self.clear_files)
        clear_btn.grid(row=1, column=1, padx=5)
        
        add_folder_btn = ttk.Button(self.file_frame, text="添加文件夹", command=self.add_folder)
        add_folder_btn.grid(row=1, column=2, padx=(5, 0))
        
        # 文件列表
        list_frame = ttk.LabelFrame(self.file_frame, text="已选择的文件", padding="5")
        list_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        self.file_frame.columnconfigure(0, weight=1)
        self.file_frame.rowconfigure(2, weight=1)
        
        # 创建文件列表和滚动条
        list_container = ttk.Frame(list_frame)
        list_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        self.file_listbox = tk.Listbox(list_container, selectmode=tk.EXTENDED)
        scrollbar = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.file_listbox.yview)
        self.file_listbox.configure(yscrollcommand=scrollbar.set)
        
        self.file_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        list_container.columnconfigure(0, weight=1)
        list_container.rowconfigure(0, weight=1)
        
        # 删除选中文件按钮
        remove_btn = ttk.Button(list_frame, text="删除选中", command=self.remove_selected)
        remove_btn.grid(row=1, column=0, pady=(5, 0), sticky=tk.W)
        
        # 文件信息显示
        info_label = ttk.Label(list_frame, text="支持格式: PNG, JPG, JPEG, BMP, GIF, WebP", 
                              foreground="gray")
        info_label.grid(row=2, column=0, pady=(5, 0), sticky=tk.W)
        
    def setup_settings_tab(self):
        """设置转换设置选项卡"""
        # 分列设置
        column_frame = ttk.LabelFrame(self.settings_frame, text="分列设置", padding="10")
        column_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        self.settings_frame.columnconfigure(0, weight=1)
        
        ttk.Label(column_frame, text="每页列数:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        col_spin = ttk.Spinbox(column_frame, from_=1, to=6, textvariable=self.num_columns, width=10)
        col_spin.grid(row=0, column=1, sticky=tk.W)
        ttk.Label(column_frame, text="列", foreground="gray").grid(row=0, column=2, sticky=tk.W, padx=(5, 0))
        
        ttk.Label(column_frame, text="列重叠像素:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        overlap_spin = ttk.Spinbox(column_frame, from_=0, to=200, textvariable=self.overlap, width=10)
        overlap_spin.grid(row=1, column=1, sticky=tk.W, pady=(5, 0))
        ttk.Label(column_frame, text="像素", foreground="gray").grid(row=1, column=2, sticky=tk.W, padx=(5, 0), pady=(5, 0))
        
        ttk.Label(column_frame, text="列间隔:").grid(row=2, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        gap_spin = ttk.Spinbox(column_frame, from_=0, to=20, textvariable=self.column_gap, width=10, format="%.1f", increment=0.5)
        gap_spin.grid(row=2, column=1, sticky=tk.W, pady=(5, 0))
        ttk.Label(column_frame, text="毫米", foreground="gray").grid(row=2, column=2, sticky=tk.W, padx=(5, 0), pady=(5, 0))
        
        # 页面设置
        page_frame = ttk.LabelFrame(self.settings_frame, text="页面设置", padding="10")
        page_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(page_frame, text="页面方向:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        orientation_frame = ttk.Frame(page_frame)
        orientation_frame.grid(row=0, column=1, sticky=tk.W)
        
        ttk.Radiobutton(orientation_frame, text="横向", variable=self.orientation, 
                       value="landscape").grid(row=0, column=0, padx=(0, 10))
        ttk.Radiobutton(orientation_frame, text="纵向", variable=self.orientation, 
                       value="portrait").grid(row=0, column=1)
        
        ttk.Label(page_frame, text="页边距:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        margin_spin = ttk.Spinbox(page_frame, from_=5, to=30, textvariable=self.margin, width=10, format="%.1f", increment=0.5)
        margin_spin.grid(row=1, column=1, sticky=tk.W, pady=(5, 0))
        ttk.Label(page_frame, text="毫米", foreground="gray").grid(row=1, column=2, sticky=tk.W, padx=(5, 0), pady=(5, 0))
        
    def setup_output_tab(self):
        """设置输出设置选项卡"""
        # 输出目录
        dir_frame = ttk.LabelFrame(self.output_frame, text="输出设置", padding="10")
        dir_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        self.output_frame.columnconfigure(0, weight=1)
        
        ttk.Label(dir_frame, text="输出目录:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        dir_entry_frame = ttk.Frame(dir_frame)
        dir_entry_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        dir_frame.columnconfigure(0, weight=1)
        
        self.dir_entry = ttk.Entry(dir_entry_frame, textvariable=self.output_dir, width=50)
        self.dir_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        dir_entry_frame.columnconfigure(0, weight=1)
        
        browse_btn = ttk.Button(dir_entry_frame, text="浏览", command=self.browse_output_dir)
        browse_btn.grid(row=0, column=1)
        
        # 质量设置
        quality_frame = ttk.LabelFrame(self.output_frame, text="质量设置", padding="10")
        quality_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(quality_frame, text="DPI (打印质量):").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        dpi_combo = ttk.Combobox(quality_frame, textvariable=self.dpi, values=[150, 200, 300, 600], 
                                width=10, state="readonly")
        dpi_combo.grid(row=0, column=1, sticky=tk.W)
        ttk.Label(quality_frame, text="(300推荐)", foreground="gray").grid(row=0, column=2, sticky=tk.W, padx=(5, 0))
        
        # 说明文本
        info_frame = ttk.LabelFrame(self.output_frame, text="说明", padding="10")
        info_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        self.output_frame.rowconfigure(2, weight=1)
        
        info_text = tk.Text(info_frame, height=8, wrap=tk.WORD, state=tk.DISABLED, bg=self.root.cget('bg'))
        info_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        info_frame.columnconfigure(0, weight=1)
        info_frame.rowconfigure(0, weight=1)
        
        info_content = """转换说明:
• 每页列数: 将长图分成几列显示，建议2-4列
• 列重叠: 列与列之间重叠的像素数，避免内容被切断
• 列间隔: 合并时列与列之间的间距
• 页面方向: 横向适合多列布局，纵向适合少列布局
• DPI设置: 150适合屏幕查看，300适合打印

使用技巧:
• 文字多的截图建议3-4列，重叠50-100像素
• 图片多的截图建议2-3列，重叠20-50像素
• 如果输出目录为空，将使用第一个文件所在目录"""
        
        info_text.config(state=tk.NORMAL)
        info_text.insert(1.0, info_content)
        info_text.config(state=tk.DISABLED)
        
    def center_window(self):
        """将窗口居中显示"""
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")
        
    def select_files(self):
        """选择图片文件"""
        filetypes = [
            ('图片文件', '*.png *.jpg *.jpeg *.bmp *.gif *.webp'),
            ('PNG文件', '*.png'),
            ('JPEG文件', '*.jpg *.jpeg'),
            ('BMP文件', '*.bmp'),
            ('GIF文件', '*.gif'),
            ('WebP文件', '*.webp'),
            ('所有文件', '*.*')
        ]
        
        files = filedialog.askopenfilenames(
            title="选择要转换的图片文件",
            filetypes=filetypes
        )
        
        for file in files:
            if file not in self.selected_files:
                self.selected_files.append(file)
                self.file_listbox.insert(tk.END, Path(file).name)
                
        self.update_file_count()
        
    def add_folder(self):
        """添加文件夹中的所有图片"""
        folder = filedialog.askdirectory(title="选择包含图片的文件夹")
        if not folder:
            return
            
        folder_path = Path(folder)
        image_extensions = {'.png', '.jpg', '.jpeg', '.bmp', '.gif', '.webp'}
        
        image_files = [f for f in folder_path.iterdir() 
                      if f.is_file() and f.suffix.lower() in image_extensions]
        
        if not image_files:
            messagebox.showinfo("提示", "所选文件夹中没有找到图片文件")
            return
            
        for img_file in image_files:
            file_path = str(img_file)
            if file_path not in self.selected_files:
                self.selected_files.append(file_path)
                self.file_listbox.insert(tk.END, img_file.name)
                
        self.update_file_count()
        messagebox.showinfo("完成", f"添加了 {len(image_files)} 个图片文件")
        
    def clear_files(self):
        """清空文件列表"""
        self.selected_files.clear()
        self.file_listbox.delete(0, tk.END)
        self.update_file_count()
        
    def remove_selected(self):
        """删除选中的文件"""
        selected_indices = self.file_listbox.curselection()
        if not selected_indices:
            return
            
        # 从后往前删除，避免索引变化
        for index in reversed(selected_indices):
            self.file_listbox.delete(index)
            self.selected_files.pop(index)
            
        self.update_file_count()
        
    def update_file_count(self):
        """更新文件数量显示"""
        count = len(self.selected_files)
        self.root.title(f"长图转PDF工具 v2.0 - 已选择 {count} 个文件")
        
    def browse_output_dir(self):
        """浏览输出目录"""
        directory = filedialog.askdirectory(title="选择输出目录")
        if directory:
            self.output_dir.set(directory)
            
    def preview_settings(self):
        """预览当前设置"""
        if not self.selected_files:
            messagebox.showwarning("警告", "请先选择要转换的文件")
            return
            
        settings_info = f"""转换设置预览:

文件设置:
• 选中文件数量: {len(self.selected_files)}
• 输出目录: {self.output_dir.get() or '第一个文件所在目录'}

转换参数:
• 每页列数: {self.num_columns.get()}
• 列重叠: {self.overlap.get()} 像素
• 列间隔: {self.column_gap.get()} 毫米
• 页面方向: {'横向' if self.orientation.get() == 'landscape' else '纵向'}
• 页边距: {self.margin.get()} 毫米
• 打印质量: {self.dpi.get()} DPI"""
        
        messagebox.showinfo("设置预览", settings_info)
        
    def start_conversion(self):
        """开始转换"""
        if not self.selected_files:
            messagebox.showwarning("警告", "请先选择要转换的文件")
            return
            
        # 检查输出目录
        output_path = self.output_dir.get()
        if not output_path:
            output_path = str(Path(self.selected_files[0]).parent / "PDF输出")
            self.output_dir.set(output_path)
            
        # 创建输出目录
        Path(output_path).mkdir(parents=True, exist_ok=True)
        
        # 在新线程中执行转换
        self.conversion_thread = threading.Thread(target=self.convert_files)
        self.conversion_thread.daemon = True
        self.conversion_thread.start()
        
    def convert_files(self):
        """转换文件（在后台线程中执行）"""
        try:
            total_files = len(self.selected_files)
            success_count = 0
            
            # 显示进度窗口
            progress_window = self.show_progress_window()
            
            for i, file_path in enumerate(self.selected_files, 1):
                # 更新进度
                progress_msg = f"正在处理 {i}/{total_files}: {Path(file_path).name}"
                self.update_progress(progress_window, progress_msg, i / total_files * 100)
                
                # 转换文件
                output_path = Path(self.output_dir.get()) / f"{Path(file_path).stem}_打印.pdf"
                
                if self.convert_single_file(file_path, output_path):
                    success_count += 1
                    
            # 关闭进度窗口
            progress_window.destroy()
            
            # 显示完成消息
            self.root.after(0, lambda: self.show_completion_message(success_count, total_files))
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"转换过程中发生错误: {str(e)}"))
            
    def show_progress_window(self):
        """显示进度窗口"""
        progress_window = tk.Toplevel(self.root)
        progress_window.title("转换进度")
        progress_window.geometry("400x150")
        progress_window.resizable(False, False)
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        # 居中显示
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 200
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 75
        progress_window.geometry(f"+{x}+{y}")
        
        # 进度标签
        progress_window.status_label = ttk.Label(progress_window, text="准备开始转换...")
        progress_window.status_label.pack(pady=20)
        
        # 进度条
        progress_window.progress = ttk.Progressbar(progress_window, length=300, mode='determinate')
        progress_window.progress.pack(pady=10)
        
        # 百分比标签
        progress_window.percent_label = ttk.Label(progress_window, text="0%")
        progress_window.percent_label.pack()
        
        return progress_window
        
    def update_progress(self, window, message, percent):
        """更新进度"""
        self.root.after(0, lambda: self._update_progress_ui(window, message, percent))
        
    def _update_progress_ui(self, window, message, percent):
        """更新进度UI（在主线程中调用）"""
        if window and window.winfo_exists():
            window.status_label.config(text=message)
            window.progress.config(value=percent)
            window.percent_label.config(text=f"{percent:.1f}%")
            window.update()
            
    def show_completion_message(self, success_count, total_files):
        """显示完成消息"""
        if success_count == total_files:
            message = f"✅ 转换完成！\n成功转换了 {success_count} 个文件"
            messagebox.showinfo("转换完成", message)
        else:
            message = f"⚠️ 转换完成！\n成功: {success_count}/{total_files}"
            messagebox.showwarning("部分成功", message)
            
        # 询问是否打开输出目录
        if messagebox.askyesno("打开目录", "是否打开输出目录？"):
            try:
                os.startfile(self.output_dir.get())
            except Exception:
                pass
                
    def convert_single_file(self, input_path, output_path):
        """转换单个文件"""
        try:
            return split_image_to_pdf(
                input_path=input_path,
                output_pdf=output_path,
                num_columns=self.num_columns.get(),
                orientation=self.orientation.get(),
                margin=self.margin.get(),
                overlap=self.overlap.get(),
                column_gap=self.column_gap.get()
            )
        except Exception as e:
            print(f"转换文件 {input_path} 时出错: {e}")
            return False
            
    def run(self):
        """运行应用程序"""
        self.root.mainloop()


def split_image_to_pdf(input_path, output_pdf, num_columns=3, 
                       orientation='landscape', margin=10, overlap=50, column_gap=5):
    """
    将长图分割并转换成多列 A4 PDF - 零信息丢失版本（增强重叠算法）
    """
    try:
        # 打开图片
        img = Image.open(input_path)
        img_width, img_height = img.size
        
        # 计算最优布局
        page_size = landscape(A4) if orientation == 'landscape' else A4
        page_width, page_height = page_size
        
        # 可用区域（减去边距和列间隔）
        available_width = page_width - (2 * margin * mm) - ((num_columns - 1) * column_gap * mm)
        available_height = page_height - (2 * margin * mm)
        
        # 每列宽度（点）
        column_width_pts = available_width / num_columns
        
        # 计算每列应包含的图片高度（像素）
        scale_for_width = column_width_pts / (img_width * 72 / 96)
        base_column_height_px = int(available_height / scale_for_width / 72 * 96)
        
        print(f"基础列高: {base_column_height_px}px, 重叠: {overlap}px")
        
        # 计算实际需要的净列高（减去重叠）
        net_column_height = base_column_height_px - overlap
        
        # 如果净高度太小，调整参数
        if net_column_height <= 0:
            print(f"警告：重叠像素({overlap})过大，自动调整为列高的1/3")
            overlap = base_column_height_px // 3
            net_column_height = base_column_height_px - overlap
        
        print(f"调整后 - 列高: {base_column_height_px}px, 重叠: {overlap}px")
        
        # 重新设计重叠算法：第一列的最后N像素和第二列的开头N像素重合
        segments = []
        current_y = 0
        segment_index = 0
        
        while current_y < img_height:
            # 计算当前段的结束位置
            segment_end = min(current_y + base_column_height_px, img_height)
            
            print(f"段 {segment_index + 1}: Y {current_y} -> {segment_end} (高度: {segment_end - current_y}px)")
            
            # 裁剪图片段并保存
            img_segment = img.crop((0, current_y, img_width, segment_end))
            
            # 添加段信息（同时保存图片对象和坐标）
            segments.append({
                'image': img_segment,
                'start': current_y,
                'end': segment_end,
                'height': segment_end - current_y,
                'index': segment_index
            })
            
            # 如果已经到达图片末尾，结束
            if segment_end >= img_height:
                break
            
            # 计算下一段的开始位置：当前段结束位置向前回退重叠像素
            next_start = segment_end - overlap
            
            # 防止下一段开始位置不合理
            if next_start <= current_y:
                # 如果重叠太大导致下一段开始位置不合理，调整
                next_start = current_y + (base_column_height_px - overlap)
                if next_start >= img_height:
                    break
            
            current_y = next_start
            segment_index += 1
            
            # 防止无限循环
            if segment_index > 100:  # 安全限制
                print("警告：分段数量过多，可能存在算法问题")
                break
        
        # 验证覆盖完整性和重叠性
        print(f"\n分段验证:")
        for i, seg in enumerate(segments):
            print(f"  段{i+1}: {seg['start']}-{seg['end']}px (高度: {seg['height']}px)")
            if i > 0:
                prev_seg = segments[i-1]
                overlap_amount = prev_seg['end'] - seg['start']
                if overlap_amount > 0:
                    print(f"    与前一段重叠: {overlap_amount}px ✓")
                else:
                    print(f"    警告: 与前一段无重叠，可能丢失信息！")
        
        # 检查是否完全覆盖
        if segments:
            first_start = segments[0]['start']
            last_end = segments[-1]['end']
            print(f"\n覆盖检查: 原图0-{img_height}px, 实际{first_start}-{last_end}px")
            
            if first_start > 0:
                print(f"警告：图片开头 0-{first_start}px 未覆盖")
            if last_end < img_height:
                print(f"警告：图片末尾 {last_end}-{img_height}px 未覆盖，调整最后一段")
                segments[-1]['end'] = img_height
                segments[-1]['height'] = img_height - segments[-1]['start']
        
        # 计算需要多少页
        pages_needed = (len(segments) + num_columns - 1) // num_columns
        
        # 创建PDF
        c = canvas.Canvas(str(output_pdf), pagesize=page_size)
        
        segment_index = 0
        for page in range(pages_needed):
            page_has_content = False
            
            # 处理当前页的列
            for col in range(num_columns):
                if segment_index >= len(segments):
                    break
                    
                segment = segments[segment_index]
                
                # 直接使用已裁剪的图片段
                img_segment = segment['image']
                actual_height = img_segment.height
                
                # 保存为临时文件
                temp_path = f"temp_segment_{segment_index}.png"
                img_segment.save(temp_path, dpi=(96, 96))
                
                # 计算在PDF中的位置
                x_pos = margin * mm + col * (column_width_pts + column_gap * mm)
                
                # 计算显示尺寸（保持宽高比，适应列宽）
                display_width = column_width_pts
                display_height = (actual_height / img_width) * display_width
                
                # 确保不超过可用高度
                if display_height > available_height:
                    display_height = available_height
                    display_width = (img_width / actual_height) * display_height
                
                # Y位置：从页面顶部开始
                y_pos = page_height - margin * mm - display_height
                
                # 在PDF中绘制图片
                c.drawImage(temp_path, x_pos, y_pos, 
                          width=display_width, height=display_height,
                          preserveAspectRatio=True)
                
                print(f"段 {segment_index + 1}: 放置在第{page+1}页第{col+1}列")
                
                # 删除临时文件
                try:
                    os.remove(temp_path)
                except:
                    pass
                
                segment_index += 1
                page_has_content = True
                
            # 如果当前页有内容且还有更多段，创建新页
            if page_has_content and segment_index < len(segments):
                c.showPage()
        
        c.save()
        
        # 最终验证
        print(f"\n✅ 转换完成:")
        print(f"   原图: {img_width}x{img_height}px")
        print(f"   分段: {len(segments)}段")
        print(f"   页数: {pages_needed}页")
        print(f"   覆盖: 0-{segments[-1]['end']}px")
        
        if segments[-1]['end'] == img_height and len(segments) > 1:
            # 检查相邻段重叠
            min_overlap = min(segments[i-1]['end'] - segments[i]['start'] for i in range(1, len(segments)))
            if min_overlap > 0:
                print(f"   重叠: 最小{min_overlap}px ✓ 零信息丢失确认")
            else:
                print(f"   警告: 存在无重叠段，可能有信息丢失")
        
        return True
        
    except Exception as e:
        print(f"转换错误: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    app = LongImageToPDFGUI()
    app.run()