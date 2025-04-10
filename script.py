import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import pythoncom
import win32com.client
import subprocess

SOFTWARE_VERSION = "v2.1"
WORD_APP_VISIBLE = False

class VisioConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Visio转Word工具 {SOFTWARE_VERSION}")  # 更新版本号

        # 初始化变量
        self.selected_dir = tk.StringVar()
        self.all_select_var = tk.BooleanVar(value=True)
        self.files_data = {}  # 存储文件名和对应的排序值及选择状态
        self.conversion_method = tk.StringVar(
            value="copy_paste"
        )  # 默认使用复制粘贴方式
        self.separate_files_var = tk.BooleanVar(
            value=False
        )  # 新增：是否单独转换每个文件

        # 创建界面组件
        self.create_widgets()

        # 配置样式
        style = ttk.Style()
        style.configure("TButton", padding=6)
        style.configure("TEntry", padding=6)

    def kill_visio_processes(self):
        """
        杀死Visio进程，为了防止用户提前打开要使用的VISIO文件，导致报错VISIO无法重复打开
        """
        try:
            
            subprocess.run(["taskkill", "/F", "/IM", "visio.exe"], check=True)
            print("所有Visio进程已终止。")
        except subprocess.CalledProcessError as e:
            print(f"终止Visio进程时出错: {e}")

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

        # 添加全选复选框
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

        # 新增单独转换复选框
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

        # 添加滚动条
        scrollbar = ttk.Scrollbar(
            list_frame, orient=tk.VERTICAL, command=self.tree.yview
        )
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 绑定事件
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Button-1>", self.on_treeview_click)

        # 控制按钮区域
        ctrl_frame = ttk.Frame(self.root, padding=10)
        ctrl_frame.pack(fill=tk.X)

        # 添加状态标签
        self.status_label = ttk.Label(ctrl_frame, text="就绪")
        self.status_label.pack(side=tk.LEFT, padx=5)

        ttk.Button(ctrl_frame, text="开始转换", command=self.start_conversion).pack(
            side=tk.RIGHT, padx=5
        )

    def on_treeview_click(self, event):
        """处理复选框点击事件"""
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            item = self.tree.identify_row(event.y)

            # 检查是否点击复选框列（第一列）
            if column == "#1":
                current_values = list(self.tree.item(item, "values"))
                new_selected = "☐" if current_values[0] == "☑" else "☑"
                current_values[0] = new_selected
                self.tree.item(item, values=tuple(current_values))

                # 更新数据存储
                filename = current_values[1]
                self.files_data[filename]["selected"] = new_selected == "☑"

                # 更新全选复选框状态
                self.update_all_select_status()

    def toggle_select_all(self):
        """全选/反选所有文件"""
        select_all = self.all_select_var.get()
        for child in self.tree.get_children():
            current_values = list(self.tree.item(child, "values"))
            current_values[0] = "☑" if select_all else "☐"
            self.tree.item(child, values=tuple(current_values))

            # 更新数据存储
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
        column = self.tree.identify_column(event.x)
        if column == "#3":  # 检查是否点击了排序号列
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

        # 更新数据存储
        filename = current_values[1]
        self.files_data[filename]["order"] = new_value

        entry.destroy()

    def select_directory(self):
        """选择目录并加载文件列表"""
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.selected_dir.set(dir_path)
            self.load_files(dir_path)

    def load_files(self, dir_path):
        """加载目录中的Visio文件到列表"""
        self.tree.delete(*self.tree.get_children())
        self.files_data.clear()

        # 获取文件列表并排序
        files = [
            f for f in os.listdir(dir_path) if f.lower().endswith((".vsd", ".vsdx"))
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

        # 在转换开始前杀死所有Visio进程

        try:
            self.kill_visio_processes()
        except Exception as e:
            messagebox.showerror("错误", f"终止Visio进程时出错: {e}")
            return

        # 获取用户输入的排序值和选择状态
        for child in self.tree.get_children():
            filename = self.tree.item(child)["values"][1]
            try:
                self.files_data[filename]["order"] = int(
                    self.tree.item(child)["values"][2]
                )
            except ValueError:
                self.files_data[filename]["order"] = 0

        # 更新状态为初始化
        self.status_label.config(text="正在初始化转换...")

        # 创建处理线程
        thread = threading.Thread(
            target=self.process_files,
            args=(
                self.selected_dir.get(),
                self.conversion_method.get(),
                self.separate_files_var.get(),  # 传递单独转换选项
            ),
        )
        thread.start()

    def process_files(self, visio_dir, method, separate_files):
        """处理文件的主逻辑"""
        try:
            # 根据排序值排序文件（固定升序），只处理选中的文件
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

            # 定义处理进度的回调函数
            def handle_progress(current_file, current, total):
                self.root.after(
                    0,
                    lambda: self.status_label.config(
                        text=f"正在处理：{current_file} ({current}/{total})"
                    ),
                )

            # 根据选择的转换方式调用不同的转换函数
            if method == "copy_paste":
                visio_to_word_copy_paste(
                    visio_dir, file_list, handle_progress, separate_files
                )
            else:
                visio_to_word_export_png(
                    visio_dir, file_list, handle_progress, separate_files
                )

            # 构造输出路径提示
            if separate_files:
                output_path = os.path.join(visio_dir, "Converted_Files")
            else:
                output_path = os.path.join(visio_dir, "output.docx")

            # 处理完成后提示
            self.root.after(
                0,
                lambda: [
                    messagebox.showinfo(
                        "完成",
                        f"文件已转换完成！\n保存路径：{output_path}",
                    ),
                    self.status_label.config(text="转换完成"),
                ],
            )
        except Exception as e:
            # debug
            # raise e
            error_msg = str(e)
            self.root.after(
                0,
                lambda msg=error_msg: [
                    messagebox.showerror("错误", f"转换失败: {msg}"),
                    self.status_label.config(text=f"错误：{msg}"),
                ],
            )


def visio_to_word_copy_paste(
    visio_dir, file_list, update_progress=None, separate_files=False
):
    """使用复制粘贴方式转换Visio到Word，支持单独转换"""
    pythoncom.CoInitialize()
    try:
        visio_app = win32com.client.Dispatch("Visio.Application")
        visio_app.Visible = False
        output_dir = None

        # 创建输出目录（如果需要单独转换）
        if separate_files:
            output_dir = os.path.join(visio_dir, "Converted_Files")
            os.makedirs(output_dir, exist_ok=True)
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = WORD_APP_VISIBLE
            
        else:
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = WORD_APP_VISIBLE
            word_doc = word_app.Documents.Add()
            word_app.Selection.EndKey(6)  # 初始光标定位到文档末尾
            

        total_files = len(file_list)
        for idx, filename in enumerate(file_list):
            if update_progress:
                update_progress(filename, idx + 1, total_files)

            visio_file_path = os.path.join(visio_dir, filename)
            visio_doc = visio_app.Documents.Open(visio_file_path)

            # 如果需要单独转换，为每个文件创建新文档
            if separate_files:
                current_doc = word_app.Documents.Add()
                current_selection = word_app.Selection
                current_selection.EndKey(6)
            else:
                current_doc = word_doc
                current_selection = word_app.Selection

            # 处理页面
            total_pages = visio_doc.Pages.Count
            for i, page in enumerate(visio_doc.Pages):
                visio_window = visio_app.ActiveWindow
                visio_window.Page = page
                visio_window.SelectAll()
                visio_window.Selection.Copy()

                # 粘贴到Word
                range_end = current_doc.Content
                range_end.Collapse(0)
                range_end.Paste()

                # 添加分页符（最后一页不添加）
                if i < total_pages - 1:
                    range_end.InsertBreak(7)

            visio_doc.Close()

            # 单独转换时保存并关闭当前文档
            if separate_files:
                output_name = os.path.splitext(filename)[0] + ".docx"
                output_path = os.path.join(output_dir, output_name)
                current_doc.SaveAs(output_path)
                current_doc.Close()

        # 统一保存并退出（非单独转换时）
        if not separate_files:
            output_word_path = os.path.join(visio_dir, "output.docx")
            word_doc.SaveAs(output_word_path)
        word_app.Quit()

    finally:
        pythoncom.CoUninitialize()


def visio_to_word_export_png(
    visio_dir, file_list, update_progress=None, separate_files=False
):
    """使用导出PNG方式转换Visio到Word，支持单独转换"""
    pythoncom.CoInitialize()
    try:
        visio_app = win32com.client.Dispatch("Visio.Application")
        visio_app.Visible = False
        output_dir = None

        if separate_files:
            output_dir = os.path.join(visio_dir, "Converted_Files")
            os.makedirs(output_dir, exist_ok=True)
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = WORD_APP_VISIBLE
        else:
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = WORD_APP_VISIBLE
            word_doc = word_app.Documents.Add()
            word_app.Selection.EndKey(6)

        total_files = len(file_list)
        for idx, filename in enumerate(file_list):
            if update_progress:
                update_progress(filename, idx + 1, total_files)

            visio_file_path = os.path.join(visio_dir, filename)
            visio_doc = visio_app.Documents.Open(visio_file_path)

            if separate_files:
                current_doc = word_app.Documents.Add()
                current_selection = word_app.Selection
                current_selection.EndKey(6)
            else:
                current_doc = word_doc
                current_selection = word_app.Selection

            total_pages = visio_doc.Pages.Count
            for i, page in enumerate(visio_doc.Pages):
                temp_image_path = os.path.join(
                    visio_dir, f"temp_{filename}_{i + 1}.png"
                )
                page.Export(temp_image_path)

                # 插入图片到Word
                range_end = current_doc.Content
                range_end.Collapse(0)
                range_end.InlineShapes.AddPicture(temp_image_path)

                if i < total_pages - 1:
                    range_end.InsertBreak(7)

                os.remove(temp_image_path)

            visio_doc.Close()

            if separate_files:
                output_name = os.path.splitext(filename)[0] + ".docx"
                output_path = os.path.join(output_dir, output_name)
                current_doc.SaveAs(output_path)
                current_doc.Close()

        if not separate_files:
            output_word_path = os.path.join(visio_dir, "output.docx")
            word_doc.SaveAs(output_word_path)
        word_app.Quit()

    finally:
        pythoncom.CoUninitialize()

def center_window(root, width, height):
    # 获取屏幕宽度和高度
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # 计算窗口左上角的位置
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2

    # 设置窗口大小和位置
    root.geometry(f"{width}x{height}+{x}+{y}")




if __name__ == "__main__":
    # 在启动应用程序之前终止所有Visio进程

    root = tk.Tk()
    app = VisioConverterApp(root)

    # 设置窗口大小
    window_width = 650
    window_height = 450

    # 居中窗口
    center_window(root, window_width, window_height)

    root.mainloop()
