import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from config import SOFTWARE_VERSION
from core import visio_to_word_copy_paste, visio_to_word_export_png, kill_visio_processes, kill_word_processes

class VisioConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Visio转Word工具 {SOFTWARE_VERSION}")

        # 初始化变量
        self.selected_dir = tk.StringVar()
        self.all_select_var = tk.BooleanVar(value=True)
        self.files_data = {}
        self.conversion_method = tk.StringVar(value="copy_paste")
        self.separate_files_var = tk.BooleanVar(value=False)
        self.word_processor = tk.StringVar(value="Word")  # 新增软件选择变量

        # 创建界面组件
        self.create_widgets()

        # 配置样式
        style = ttk.Style()
        style.configure("TButton", padding=6)
        style.configure("TEntry", padding=6)

    def create_widgets(self):
        # 目录选择区域
        dir_frame = ttk.Frame(self.root, padding=10)
        dir_frame.pack(fill=tk.X)

        ttk.Label(dir_frame, text="目标目录:").pack(side=tk.LEFT)
        ttk.Entry(dir_frame, textvariable=self.selected_dir, width=50).pack(
            side=tk.LEFT, padx=5
        )
        ttk.Button(dir_frame, text="选择目录", command=self.select_directory).pack(
            side=tk.LEFT
        )
        ttk.Checkbutton(
            dir_frame,
            text="全选",
            variable=self.all_select_var,
            command=self.toggle_select_all,
        ).pack(side=tk.LEFT, padx=5)

        # 转换方式选择区域
        method_frame = ttk.Frame(self.root, padding=(10, 5))
        method_frame.pack(fill=tk.X)

        ttk.Label(method_frame, text="转换方式:").pack(side=tk.LEFT)
        ttk.Radiobutton(
            method_frame,
            text="直接复制粘贴",
            variable=self.conversion_method,
            value="copy_paste",
        ).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(
            method_frame,
            text="导出PNG图片",
            variable=self.conversion_method,
            value="export_png",
        ).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(
            method_frame, text="单独转换每个文件", variable=self.separate_files_var
        ).pack(side=tk.LEFT, padx=5)

        # 文件列表区域
        list_frame = ttk.Frame(self.root, padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(
            list_frame, columns=("selected", "filename", "order"), show="headings"
        )
        self.tree.heading("selected", text="选择", anchor=tk.CENTER)
        self.tree.heading("filename", text="文件名", anchor=tk.W)
        self.tree.heading("order", text="排序号", anchor=tk.W)
        self.tree.column("selected", width=50, anchor=tk.CENTER)
        self.tree.column("filename", width=350)
        self.tree.column("order", width=100)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(
            list_frame, orient=tk.VERTICAL, command=self.tree.yview
        )
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Button-1>", self.on_treeview_click)

        # 控制按钮区域
        ctrl_frame = ttk.Frame(self.root, padding=10)
        ctrl_frame.pack(fill=tk.X)

        self.status_label = ttk.Label(ctrl_frame, text="就绪")
        self.status_label.pack(side=tk.LEFT, padx=5)

        # 软件选择组件
        software_frame = ttk.Frame(ctrl_frame)
        software_frame.pack(side=tk.RIGHT, padx=5)
        ttk.Label(software_frame, text="选择软件:").pack(side=tk.LEFT)
        ttk.Radiobutton(
            software_frame, text="Word", variable=self.word_processor, value="Word"
        ).pack(side=tk.LEFT)
        ttk.Radiobutton(
            software_frame, text="WPS", variable=self.word_processor, value="WPS"
        ).pack(side=tk.LEFT)

        ttk.Button(ctrl_frame, text="开始转换", command=self.start_conversion).pack(
            side=tk.RIGHT, padx=5
        )

    def on_treeview_click(self, event):
        """处理复选框点击事件"""
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            item = self.tree.identify_row(event.y)

            if column == "#1":
                current_values = list(self.tree.item(item, "values"))
                new_selected = "☐" if current_values[0] == "☑" else "☑"
                current_values[0] = new_selected
                self.tree.item(item, values=tuple(current_values))

                filename = current_values[1]
                self.files_data[filename]["selected"] = new_selected == "☑"
                self.update_all_select_status()

    def toggle_select_all(self):
        """全选/反选所有文件"""
        select_all = self.all_select_var.get()
        for child in self.tree.get_children():
            current_values = list(self.tree.item(child, "values"))
            current_values[0] = "☑" if select_all else "☐"
            self.tree.item(child, values=tuple(current_values))

            filename = current_values[1]
            self.files_data[filename]["selected"] = select_all

    def update_all_select_status(self):
        """自动更新全选复选框状态"""
        all_selected = all(
            self.files_data[filename]["selected"] for filename in self.files_data
        )
        self.all_select_var.set(all_selected)

    def on_double_click(self, event):
        """双击事件处理函数，用于编辑排序号"""
        item = self.tree.selection()[0]
        column = self.ttree.identify_column(event.x)
        if column == "#3":
            x, y, width, height = self.tree.bbox(item, column)
            entry = tk.Entry(self.tree, width=10)
            entry.place(x=x, y=y, width=width, height=height)

            entry.insert(0, self.tree.item(item, "values")[2])
            entry.focus_set()

            entry.bind("<FocusOut>", lambda e: self.on_focus_out(entry, item, column))
            entry.bind("<Return>", lambda e: self.on_focus_out(entry, item, column))

    def on_focus_out(self, entry, item, column):
        """焦点离开事件处理函数，用于保存编辑内容"""
        new_value = entry.get()
        try:
            new_value = int(new_value)
        except ValueError:
            new_value = 0

        current_values = list(self.tree.item(item, "values"))
        current_values[2] = new_value
        self.tree.item(item, values=tuple(current_values))

        filename = current_values[1]
        self.files_data[filename]["order"] = new_value

        entry.destroy()

    def select_directory(self):
        """选择目录并加载文件列表"""
        dir_path = filedialog.askdirectory()
        if dir_path:
            # 规范化路径
            dir_path = os.path.abspath(os.path.normpath(dir_path))
            self.selected_dir.set(dir_path)
            self.load_files(dir_path)

    def load_files(self, dir_path):
        """加载目录中的Visio文件到列表"""
        self.tree.delete(*self.tree.get_children())
        self.files_data.clear()

        # 确保路径是绝对路径
        dir_path = os.path.abspath(os.path.normpath(dir_path))
        
        files = [
            f for f in os.listdir(dir_path) 
            if f.lower().endswith((".vsd", ".vsdx")) and os.path.isfile(os.path.join(dir_path, f))
        ]
        files.sort()

        for index, filename in enumerate(files):
            selected_icon = "☑" if self.all_select_var.get() else "☐"
            self.files_data[filename] = {
                "selected": self.all_select_var.get(),
                "order": index + 1,
            }
            self.tree.insert("", tk.END, values=(selected_icon, filename, index + 1))

    def start_conversion(self):
        """启动转换流程"""
        if not self.selected_dir.get():
            messagebox.showerror("错误", "请先选择目录！")
            return

        try:
            kill_visio_processes()
            kill_word_processes(self.word_processor.get())
        except Exception as e:
            messagebox.showerror("错误", f"终止进程时出错: {e}")
            return

        for child in self.tree.get_children():
            filename = self.tree.item(child)["values"][1]
            try:
                self.files_data[filename]["order"] = int(
                    self.tree.item(child)["values"][2]
                )
            except ValueError:
                self.files_data[filename]["order"] = 0

        self.status_label.config(text="正在初始化转换...")

        thread = threading.Thread(
            target=self.process_files,
            args=(
                self.selected_dir.get(),
                self.conversion_method.get(),
                self.separate_files_var.get(),
                self.word_processor.get(),
            ),
        )
        thread.start()

    def process_files(self, visio_dir, method, separate_files, word_processor):
        """处理文件的主逻辑"""
        try:
            # 确保路径是绝对路径且规范化
            visio_dir = os.path.abspath(os.path.normpath(visio_dir))
            
            sorted_files = sorted(
                [
                    (f, self.files_data[f])
                    for f in self.files_data
                    if self.files_data[f]["selected"]
                ],
                key=lambda x: x[1]["order"],
                reverse=False,
            )
            file_list = [f[0] for f in sorted_files]

            if not file_list:
                self.root.after(
                    0, lambda: messagebox.showwarning("警告", "没有选择任何文件！")
                )
                return

            def handle_progress(current_file, current, total):
                self.root.after(
                    0,
                    lambda: self.status_label.config(
                        text=f"正在处理：{current_file} ({current}/{total})"
                    ),
                )

            if method == "copy_paste":
                visio_to_word_copy_paste(
                    visio_dir,
                    file_list,
                    handle_progress,
                    separate_files,
                    word_processor,
                )
            else:
                visio_to_word_export_png(
                    visio_dir,
                    file_list,
                    handle_progress,
                    separate_files,
                    word_processor,
                )

            output_path = os.path.abspath(os.path.join(
                visio_dir, "Converted_Files" if separate_files else "output.docx"
            ))
            self.root.after(
                0,
                lambda: [
                    messagebox.showinfo(
                        "完成", f"文件已转换完成！\n保存路径：{output_path}"
                    ),
                    self.status_label.config(text="转换完成"),
                ],
            )
        except Exception as e:
            # DEBUG
            raise e
            error_msg = str(e)
            self.root.after(
                0,
                lambda msg=error_msg: [
                    messagebox.showerror("错误", f"转换失败: {msg}"),
                    self.status_label.config(text=f"错误：{msg}"),
                ],
            )

def center_window(root, width, height):
    """窗口居中显示"""
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2 - 60
    root.geometry(f"{width}x{height}+{x}+{y}")

if __name__ == "__main__":
    root = tk.Tk()
    app = VisioConverterApp(root)
    center_window(root, 650, 450)
    root.mainloop()