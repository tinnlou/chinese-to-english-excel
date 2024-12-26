# -*- coding: utf-8 -*-
from deep_translator import GoogleTranslator
import pandas as pd
import time
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.font import Font
import threading
from queue import Queue

class ExcelTranslator:
    """Excel文件中英文翻译工具"""
    
    def __init__(self):
        """初始化翻译器和配置"""
        self.translator = GoogleTranslator(source='zh-CN', target='en')
        self.max_retries = 5  # 增加重试次数
        self.min_delay = 1    # 最小延迟
        self.max_delay = 5    # 最大延迟
        self.current_delay = self.min_delay  # 当前延迟
        self.batch_size = 5   # 批量处理大小
        self.cancel_flag = False  # 添加取消标志
        
    def translate_text(self, text):
        """翻译单条文本"""
        # 添加取消检查
        if self.cancel_flag:
            print("检测到取消标志，停止翻译")
            return text
            
        if not isinstance(text, str) or not text.strip():
            return text
            
        print(f"\n[翻译] 开始: {text[:50]}..." if len(text) > 50 else f"\n[翻译] 开始: {text}")
        
        delay = self.min_delay
        for i in range(self.max_retries):
            # 次重试前检查取消标志
            if self.cancel_flag:
                print("[翻译] 检测到取消标志，停止翻译")
                return text
                
            try:
                result = self.translator.translate(text=text)
                print(f"[翻译] 成功: {result[:50]}..." if len(result) > 50 else f"[翻译] 成功: {result}")
                return result
                
            except Exception as e:
                print(f"[翻译] 出错 (尝试 {i+1}/{self.max_retries})")
                print(f"[翻译] 错误类型: {type(e).__name__}")
                if self.cancel_flag:
                    print("[翻译] 检测到取消标志，停止重试")
                    return text
                time.sleep(delay)
                delay = min(self.max_delay, delay * 2)

    def translate_batch(self, texts):
        """批量翻译文本"""
        if not texts:
            return []
            
        try:
            text_list = texts.tolist() if hasattr(texts, 'tolist') else list(texts)
            results = []
            
            print(f"\n[批量] 开始处理 {len(text_list)} 个文本")
            for idx, text in enumerate(text_list):
                # 检查取消标志
                if self.cancel_flag:
                    print("[批量] 检测到取消标志，停止批量翻译")
                    return results
                    
                print(f"[批量] 处理第 {idx+1}/{len(text_list)} 个文本")
                result = self.translate_text(text)
                results.append(result)
                
            return results
            
        except Exception as e:
            print(f"[批量] 出错: {str(e)}")
            return [self.translate_text(text) for text in text_list]

    def process_excel(self, input_file, output_file, progress_callback=None):
        """处理Excel文件"""
        try:
            print("\n=== 开始处理Excel文件 ===")
            
            # 添加时间跟踪变量
            start_time = time.time()
            last_update_time = start_time
            last_processed_values = 0
            processing_speed = 0  # 每秒处理的条目数
            
            # 检查文件大小
            file_size = os.path.getsize(input_file)
            if progress_callback:
                progress_callback(0, f"文件大小: {file_size/1024/1024:.1f}MB")
            
            if file_size > 10 * 1024 * 1024:  # 大于10MB
                if not messagebox.askyesno("警告", 
                    f"文件大小为 {file_size/1024/1024:.1f}MB，处理可能需要较长时间。\n是否继续？"):
                    return False
                    
            print(f"开始读取Excel文件: {input_file}")
            excel_file = pd.ExcelFile(input_file)
            
            # 使用 with 语句来确保正确关闭文件
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                total_sheets = len(excel_file.sheet_names)
                print(f"共发现 {total_sheets} 个工作表")
                
                # 计算总的唯一值数量
                total_all_values = 0
                processed_values = 0
                sheet_values_map = {}
                
                # 预扫描所有工作表的唯一值数量
                for sheet_name in excel_file.sheet_names:
                    df = pd.read_excel(input_file, sheet_name=sheet_name)
                    sheet_values = 0
                    for column in df.columns:
                        unique_count = len(df[column].dropna().unique())
                        sheet_values += unique_count
                    sheet_values_map[sheet_name] = sheet_values
                    total_all_values += sheet_values
                
                # 处理每个工作表
                for sheet_idx, sheet_name in enumerate(excel_file.sheet_names):
                    if self.cancel_flag:
                        print("\n[处理] 检测到取消标志，停止处理工作表")
                        return False
                        
                    print(f"\n[DEBUG] 开始处理工作表 {sheet_idx + 1}/{total_sheets}: {sheet_name}")
                    
                    df = pd.read_excel(input_file, sheet_name=sheet_name)
                    print(f"[DEBUG] 成功读取工作表，列数: {len(df.columns)}")
                    
                    total_cells = len(df.columns) * len(df)
                    cells_processed = 0
                    
                    print(f"工作表大小: {len(df)} 行 x {len(df.columns)} 列")
                    
                    if progress_callback:
                        sheet_progress = (sheet_idx / total_sheets) * 100
                        progress_callback(sheet_progress, f"正在处理工作表: {sheet_name}")
                    
                    for column_idx, column in enumerate(df.columns):
                        if self.cancel_flag:
                            print("\n[处理] 检测到取消标志，停止处理列")
                            return False
                            
                        print(f"\n[DEBUG] 开始处理第 {column_idx + 1}/{len(df.columns)} 列: {column}")
                        
                        if progress_callback:
                            column_progress = (cells_processed / total_cells) * 100
                            total_progress = (sheet_progress + column_progress) / 2
                            status_line1 = f"正在翻译: {sheet_name} - {column}"
                            progress_callback(total_progress, status_line1)
                        
                        # 尝试批量翻译
                        values = df[column].dropna().unique()
                        total_unique = len(values)
                        print(f"该列共有 {total_unique} 个唯一值需要翻译")
                        
                        if total_unique > 0:
                            # 分批处理
                            translations = {}
                            for i in range(0, total_unique, self.batch_size):
                                batch = values[i:i+self.batch_size]
                                batch_list = batch.tolist() if hasattr(batch, 'tolist') else list(batch)
                                
                                # 更新处理进度信息
                                if progress_callback:
                                    batch_end = min(i + self.batch_size, total_unique)
                                    # 第二行显示批次进度
                                    status_line2 = f"处理进度: 第{i+1}-{batch_end}条(共{total_unique}条唯一值) | 已完成{int((i/total_unique)*100)}%"
                                    
                                    # 更新处理速度和预计剩余时间
                                    current_time = time.time()
                                    time_elapsed = current_time - last_update_time
                                    if time_elapsed >= 1:  # 每秒更新一次速度
                                        values_processed = processed_values - last_processed_values
                                        processing_speed = values_processed / time_elapsed
                                        last_update_time = current_time
                                        last_processed_values = processed_values
                                    
                                    # 计算预计剩余时间
                                    if processing_speed > 0:
                                        remaining_values = total_all_values - processed_values
                                        estimated_seconds = remaining_values / processing_speed
                                        if estimated_seconds < 60:
                                            time_str = f"{int(estimated_seconds)}秒"
                                        elif estimated_seconds < 3600:
                                            time_str = f"{int(estimated_seconds/60)}分钟"
                                        else:
                                            hours = int(estimated_seconds/3600)
                                            minutes = int((estimated_seconds % 3600)/60)
                                            time_str = f"{hours}小时{minutes}分钟"
                                    else:
                                        time_str = "计算中..."
                                    
                                    # 第三行显示总体进度和预计剩余时间
                                    processed_values += len(batch)
                                    total_progress = int((processed_values / total_all_values) * 100)
                                    status_line3 = f"总体进度: {processed_values}/{total_all_values}条唯一值 | 完成{total_progress}% | 预计剩余: {time_str}"
                                    progress_callback(total_progress, f"{status_line1}\n{status_line2}\n{status_line3}")
                                
                                print(f"\n处理第 {i+1}-{batch_end} 个值")
                                batch_results = self.translate_batch(batch_list)
                                translations.update(dict(zip(batch_list, batch_results)))
                            
                            # 使用映射进行翻译
                            en_column = f"{column}_EN"
                            df[en_column] = df[column].map(lambda x: translations.get(x, x))
                        else:
                            print("该列没有需要翻译的有效值")
                            en_column = f"{column}_EN"
                            df[en_column] = df[column]
                        
                        # 更新处理进度
                        cells_processed += len(df)
                        print(f"[DEBUG] 列 {column} 处理完成")
                        print(f"[DEBUG] DataFrame 列数: {len(df.columns)}")
                        print(f"[DEBUG] 当前列: {list(df.columns)}")
                        
                        # 重新排序列
                        try:
                            cols = list(df.columns)
                            idx = cols.index(column)
                            print(f"[DEBUG] 当前列索引: {idx}")
                            print(f"[DEBUG] 英文列名: {en_column}")
                            cols.remove(en_column)
                            cols.insert(idx + 1, en_column)
                            print(f"[DEBUG] 重排后的列: {cols}")
                            df = df[cols]
                        except Exception as e:
                            print(f"[DEBUG] 列重排序出错: {str(e)}")
                        
                    print(f"\n[DEBUG] 准备保存工作表: {sheet_name}")
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"[DEBUG] 工作表 {sheet_name} 保存完成")
                
                print("\n保存文件...")
                # 不需要显式调用 save() 方法，with 语句会自动处理
                
            print(f"文件已保存至: {output_file}")
            return True
            
        except Exception as e:
            if progress_callback:
                progress_callback(0, f"处理出错: {str(e)}")
            return False

class TranslatorGUI:
    """翻译工具GUI界面 - Apple风格"""
    
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Excel翻译工具")
        self.window.geometry("600x380")  
        self.window.overrideredirect(False)
        self.window.attributes('-alpha', 0.98)
        
        # 配置Apple风格
        self.setup_style()
        
        # 添加文件路径变量
        self.file_path = tk.StringVar()
        
        self.translator = ExcelTranslator()
        self.setup_ui()
        self.bind_hover_effects()
        self.bind_shortcuts()
        
        # 保存最后使用的目录
        self.last_dir = os.path.expanduser("~")
        
        # 添加取消标志
        self.cancel_translation = False
        
        # 添加消息队列用于线程间通信
        self.message_queue = Queue()
        # 添加翻译状态标志
        self.is_translating = False
        
    def setup_style(self):
        """配置Apple风格主题"""
        # 设置窗口背景色为浅灰色
        self.window.configure(bg='#F5F5F7')  # Apple 经典的浅灰色背景
        
        # 创建自定义字体
        self.default_font = Font(family='Microsoft YaHei', size=9)
        self.title_font = Font(family='Microsoft YaHei', size=16, weight='bold')  # 增大标题字号
        self.subtitle_font = Font(family='Microsoft YaHei', size=9)  # 副标题字体
        
        # 配置进度条样式
        self.style = ttk.Style()
        self.style.layout('Mac.TProgressbar', 
            [('Horizontal.Progressbar.trough',
                {'children': [('Horizontal.Progressbar.pbar',
                              {'side': 'left', 'sticky': 'ns'})],
                 'sticky': 'nswe'})])
        
        # 进度条使用更现代的样式
        self.style.configure('Mac.TProgressbar',
            troughcolor='#F5F5F7',    # 更浅的轨道颜色
            background='#0A84FF',     # Apple 蓝色
            thickness=8,              # 增加高度
            borderwidth=0,
            relief='flat'
        )
        
    def setup_ui(self):
        """设置GUI界面"""
        # 主容器
        main_frame = tk.Frame(
            self.window,
            bg='#FFFFFF',  # 纯白背景
        )
        main_frame.pack(expand=True, fill='both', padx=30, pady=25)
        
        # 标题区域
        title_frame = tk.Frame(main_frame, bg='white')
        title_frame.pack(fill='x', pady=(0, 30))
        
        title_label = tk.Label(
            title_frame,
            text="Excel 文件翻译工具",
            font=self.title_font,
            bg='white',
            fg='#1D1D1F'  # Apple 的深灰色
        )
        title_label.pack()
        
        subtitle_label = tk.Label(
            title_frame,
            text="将 Excel 文件中的中文内容翻译成英文",
            font=self.subtitle_font,
            bg='white',
            fg='#86868B'  # Apple 的次要文本颜色
        )
        subtitle_label.pack(pady=(5, 0))
        
        # 文件选择框
        self.file_frame = tk.Frame(
            main_frame,
            bg='#FFFFFF',
            bd=1,
            relief='solid'
        )
        self.file_frame.pack(fill='x', padx=2, pady=(0, 15))
        
        # 文件输入框
        self.file_entry = tk.Entry(
            self.file_frame,
            textvariable=self.file_path,
            font=self.default_font,
            relief='flat',
            bg='#FFFFFF',
            fg='#1D1D1F',
            width=50
        )
        self.file_entry.pack(side=tk.LEFT, padx=(16, 5), pady=12)
        
        # 选择文件按钮
        self.browse_button = tk.Button(
            self.file_frame,
            text="选择文件",
            font=self.default_font,
            bg='#0A84FF',
            fg='white',
            relief='flat',
            padx=20,
            pady=6,
            cursor='hand2',
            command=self.browse_file
        )
        self.browse_button.pack(side=tk.RIGHT, padx=(5, 10), pady=4)
        
        # 进度条框架
        self.progress_frame = tk.Frame(main_frame, bg='white')
        self.progress_frame.pack(fill='x', pady=15)
        
        self.progress_bar = ttk.Progressbar(
            self.progress_frame,
            style='Mac.TProgressbar',
            orient=tk.HORIZONTAL,
            mode='determinate'
        )
        self.progress_bar.pack(fill='x', padx=2)
        
        # 状态标签
        self.status_label = tk.Label(
            main_frame,
            text="请选择要翻译的Excel文件",
            font=self.default_font,
            bg='white',
            fg='#86868B',  # Apple 的次要文本颜色
            justify=tk.CENTER,  # 文本居中对齐
            wraplength=450     # 适当调整换行宽度
        )
        self.status_label.pack(pady=10)
        
        # 按钮容器增加上边距
        button_frame = tk.Frame(main_frame, bg='white')
        button_frame.pack(pady=(25, 20))
        
        # 主按钮 - 添加固定宽度
        self.start_button = tk.Button(
            button_frame,
            text="开始翻译",
            font=self.default_font,
            bg='#0A84FF',
            fg='white',
            relief='flat',
            width=12,      # 添加固定宽度
            height=1,      # 添加固定高度
            cursor='hand2',
            command=self.start_translation
        )
        self.start_button.pack(side=tk.LEFT, padx=8)
        
        # 取消按钮 - 添加固定宽度
        self.cancel_button = tk.Button(
            button_frame,
            text="取消",
            font=self.default_font,
            bg='#FF453A',
            fg='white',
            relief='flat',
            width=12,      # 添加固定宽度
            height=1,      # 添加固定高度
            cursor='hand2',
            command=self.cancel_translation_task,
            state=tk.DISABLED
        )
        self.cancel_button.pack(side=tk.LEFT, padx=8)
        
    def bind_hover_effects(self):
        """绑定按钮悬停效果"""
        def on_enter(e):
            if e.widget == self.browse_button or e.widget == self.start_button:
                e.widget['bg'] = '#0051D5'
            elif e.widget == self.cancel_button:
                e.widget['bg'] = '#E01E1E'
        
        def on_leave(e):
            if e.widget == self.browse_button or e.widget == self.start_button:
                e.widget['bg'] = '#0066FF'
            elif e.widget == self.cancel_button:
                e.widget['bg'] = '#FF3B30'
        
        for button in (self.browse_button, self.start_button, self.cancel_button):
            button.bind('<Enter>', on_enter)
            button.bind('<Leave>', on_leave)
        
    def browse_file(self):
        """打开文件选择对话框"""
        file_path = filedialog.askopenfilename(
            initialdir=self.last_dir,
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.last_dir = os.path.dirname(file_path)
            self.file_path.set(file_path)
            self.status_label.config(
                text='已选择文件，点击"开始翻译"按钮开始处理 (Ctrl+Enter)',
                fg='#1d1d1f'
            )
            
    def update_ui(self):
        """更新UI的循环"""
        if not self.is_translating:
            return
            
        while not self.message_queue.empty():
            msg_type, data = self.message_queue.get()
            if msg_type == 'progress':
                progress, status = data
                self.progress_bar['value'] = progress
                self.status_label.config(text=status)
            elif msg_type == 'log':
                print(data)  # 在控制台显示日志
                
        # 每100ms检查一次消息队列
        self.window.after(100, self.update_ui)
        
    def queue_message(self, msg_type, data):
        """消息添加到队列"""
        self.message_queue.put((msg_type, data))
        
    def translation_callback(self, progress, status):
        """翻译进度回调"""
        self.queue_message('progress', (progress, status))
        
    def start_translation(self):
        """开始翻译流程"""
        input_file = self.file_path.get()
        if not input_file:
            messagebox.showerror("错误", "请先选择Excel文件!")
            return
            
        # 重置状态
        self.cancel_translation = False
        self.translator.cancel_flag = False  # 重置翻译器的取消标志
        self.is_translating = True
        
        # 更新按钮状态
        self.start_button.config(state=tk.DISABLED)
        self.browse_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)
        
        # 生成输出文件路径
        file_dir = os.path.dirname(input_file)
        file_name = os.path.basename(input_file)
        name, ext = os.path.splitext(file_name)
        output_file = os.path.join(file_dir, f"{name}_translated{ext}")
        
        # 启动UI更新循环
        self.update_ui()
        
        # 在新线程中执行翻译
        def translation_thread():
            try:
                success = self.translator.process_excel(
                    input_file, 
                    output_file,
                    self.translation_callback
                )
                
                # 主线程中更新UI
                self.window.after(0, self.translation_completed, success)
                
            except Exception as e:
                self.queue_message('log', f"翻译出错: {str(e)}")
                self.window.after(0, self.translation_completed, False)
            
        thread = threading.Thread(target=translation_thread)
        thread.daemon = True  # 设置为守护线程
        thread.start()
        
    def translation_completed(self, success):
        """翻译完成后的处理"""
        self.is_translating = False
        
        # 恢复按钮状态
        self.start_button.config(state=tk.NORMAL)
        self.browse_button.config(state=tk.NORMAL)
        self.cancel_button.config(state=tk.DISABLED)
        
        if success:
            self.status_label.config(
                text="翻译完成!",
                fg='#34C759'  # Apple绿色
            )
            messagebox.showinfo(
                "完成",
                f"翻译完成!\n\n文件已保存至:\n{self.file_path.get()[:-5]}_translated.xlsx"
            )
        elif self.cancel_translation:
            self.status_label.config(
                text="翻译已取消",
                fg='#FF3B30'  # Apple红色
            )
        
    def cancel_translation_task(self):
        """取消翻译任务"""
        if messagebox.askyesno("确认", "确定要取消翻译吗？"):
            print("\n=== 开始取消翻译 ===")
            print("1. 设置GUI取消标志")
            self.cancel_translation = True
            
            print("2. 设置翻译器取消标志")
            self.translator.cancel_flag = True
            
            print("3. 更新状态显示")
            self.status_label.config(text="正在取消...")
            print("=== 取消流程初始化完成 ===\n")
        
    def bind_shortcuts(self):
        """绑定键盘快捷键"""
        def start_translation_hotkey(event):
            # 只在按钮可用时响应快捷键
            if self.start_button['state'] != tk.DISABLED:
                self.start_translation()
        
        # 绑定 Ctrl+Enter 快捷键
        self.window.bind('<Control-Return>', start_translation_hotkey)
        
    def run(self):
        """运行GUI程序"""
        # 设置窗口图标
        try:
            self.window.iconbitmap('app.ico')
        except:
            print("Warning: 图标文件未找到")
        
        # 设置窗口最小尺寸
        self.window.minsize(600, 400)
        # 居中显示
        self.window.eval('tk::PlaceWindow . center')
        # 运行
        self.window.mainloop()

def main():
    """主函数"""
    app = TranslatorGUI()
    app.run()

if __name__ == "__main__":
    main() 