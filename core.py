import os
import pythoncom
import win32com.client
import subprocess
from config import WORD_APP_VISIBLE




def kill_visio_processes():
    """
    终止所有正在运行的 Visio 进程以防止文件被占用。

    使用Windows的taskkill命令强制终止所有visio.exe进程。
    如果终止失败会捕获异常并打印错误信息。
    """
    try:
        subprocess.run(["taskkill", "/F", "/IM", "visio.exe"], check=True)
        print("所有Visio进程已终止。")
    except subprocess.CalledProcessError as e:
        print(f"终止Visio进程时出错: {e}")


def kill_word_processes(word_processor="Word"):
    """
    根据指定的应用程序类型终止 Word 或 WPS 进程。

    参数:
        word_processor (str): 指定要终止的办公软件类型，可选"Word"或"WPS"

    根据参数决定终止winword.exe(Word)或wps.exe(WPS)进程。
    使用非严格模式(taskkill的check=False)，即使进程不存在也不会报错。
    """
    targets = ["winword.exe"] if word_processor == "Word" else ["wps.exe"]
    try:
        for proc in targets:
            subprocess.run(["taskkill", "/F", "/IM", proc], check=False)  # 非严格模式
        print(f"已终止{word_processor}相关进程")
    except Exception as e:
        print(f"进程终止异常: {e}")


def create_office_app(app_type):
    """
    创建指定类型的办公软件应用程序实例 (Word 或 WPS)。

    参数:
        app_type (str): 应用程序类型，"Word"或"WPS"

    返回:
        win32com.client.Dispatch对象: 创建的办公应用程序实例

    对于Word直接创建实例，对于WPS会尝试不同版本(Kwps.Application和Wps.Application)。
    如果都无法创建则抛出异常。
    """
    if app_type == "Word":
        return win32com.client.Dispatch("Word.Application")

    # 尝试不同WPS版本
    for prog_id in ["Kwps.Application", "Wps.Application"]:
        try:
            return win32com.client.Dispatch(prog_id)
        except:
            continue
    raise Exception("无法启动WPS，请确认安装正确版本")


def visio_to_word_copy_paste(
    visio_dir,
    file_list,
    update_progress=None,
    separate_files=False,
    word_processor="Word",
):
    """
    使用复制粘贴方式将Visio文件内容转换到Word/WPS文档中。

    参数:
        visio_dir (str): Visio文件所在目录路径
        file_list (list): 要转换的Visio文件名列表
        update_progress (function, 可选): 进度回调函数，格式为func(文件名, 当前序号, 总数)
        separate_files (bool): 是否每个Visio文件生成单独的Word文档
        word_processor (str): 目标办公软件类型，"Word"或"WPS"

    流程:
    1. 初始化COM环境
    2. 启动Visio和Word/WPS应用程序
    3. 根据separate_files决定创建单个或多个Word文档
    4. 遍历每个Visio文件，复制所有页面内容到Word
    5. 保存生成的Word文档并退出应用程序

    注意:
    - 使用前确保没有Visio和Word/WPS进程运行(可调用kill_*_processes)
    - 会创建临时Word应用程序实例，操作完成后自动退出
    """
    pythoncom.CoInitialize()
    try:
        visio_app = win32com.client.Dispatch("Visio.Application")
        visio_app.Visible = False
        output_dir = None

        office_app = create_office_app(word_processor)
        office_app.Visible = WORD_APP_VISIBLE

        if not separate_files:
            doc = office_app.Documents.Add()
            office_app.Selection.EndKey(6)

        total_files = len(file_list)
        for idx, filename in enumerate(file_list):
            if update_progress:
                update_progress(filename, idx + 1, total_files)

            visio_file_path = os.path.normpath(os.path.join(visio_dir, filename))
            visio_doc = visio_app.Documents.Open(visio_file_path)

            if separate_files:
                current_doc = office_app.Documents.Add()
                current_selection = office_app.Selection
                current_selection.EndKey(6)
            else:
                current_doc = doc
                current_selection = office_app.Selection

            total_pages = visio_doc.Pages.Count
            for i, page in enumerate(visio_doc.Pages):
                visio_window = visio_app.ActiveWindow
                visio_window.Page = page
                visio_window.SelectAll()
                visio_window.Selection.Copy()

                range_end = current_doc.Content
                range_end.Collapse(0)
                range_end.Paste()

                if i < total_pages - 1:
                    range_end.InsertBreak(7)

            visio_doc.Close()

            if separate_files:
                if separate_files:
                    output_name = os.path.splitext(filename)[0] + ".docx"
                    output_path = os.path.join(
                        visio_dir, "Converted_Files", output_name
                    )
                    os.makedirs(os.path.dirname(output_path), exist_ok=True)
                    current_doc.SaveAs(output_path)
                    current_doc.Close()

        if not separate_files:
            output_word_path = os.path.join(visio_dir, "output.docx")
            doc.SaveAs(output_word_path)
        office_app.Quit()

    finally:
        pythoncom.CoUninitialize()


def visio_to_word_export_png(
    visio_dir,
    file_list,
    update_progress=None,
    separate_files=False,
    word_processor="Word",
):
    """
    使用导出PNG图片方式将Visio文件内容转换到Word/WPS文档中。

    参数:
        visio_dir (str): Visio文件所在目录路径
        file_list (list): 要转换的Visio文件名列表
        update_progress (function, 可选): 进度回调函数，格式为func(文件名, 当前序号, 总数)
        separate_files (bool): 是否每个Visio文件生成单独的Word文档
        word_processor (str): 目标办公软件类型，"Word"或"WPS"

    流程:
    1. 初始化COM环境
    2. 启动Visio和Word/WPS应用程序
    3. 根据separate_files决定创建单个或多个Word文档
    4. 遍历每个Visio文件，将每页导出为PNG图片并插入Word
    5. 保存生成的Word文档并退出应用程序

    注意:
    - 会创建临时PNG图片文件，操作完成后自动删除
    - 使用前确保没有Visio和Word/WPS进程运行(可调用kill_*_processes)
    - 会创建临时Word应用程序实例，操作完成后自动退出
    """
    pythoncom.CoInitialize()
    try:
        visio_app = win32com.client.Dispatch("Visio.Application")
        visio_app.Visible = False

        office_app = create_office_app(word_processor)
        office_app.Visible = WORD_APP_VISIBLE

        if not separate_files:
            doc = office_app.Documents.Add()
            office_app.Selection.EndKey(6)

        total_files = len(file_list)
        for idx, filename in enumerate(file_list):
            if update_progress:
                update_progress(filename, idx + 1, total_files)

            visio_file_path = os.path.normpath(os.path.join(visio_dir, filename))
            visio_doc = visio_app.Documents.Open(visio_file_path)

            if separate_files:
                current_doc = office_app.Documents.Add()
                current_selection = office_app.Selection
                current_selection.EndKey(6)
            else:
                current_doc = doc
                current_selection = office_app.Selection

            total_pages = visio_doc.Pages.Count
            for i, page in enumerate(visio_doc.Pages):
                temp_image_path = os.path.join(
                    visio_dir, f"temp_{filename}_{i + 1}.png"
                )
                page.Export(temp_image_path)

                range_end = current_doc.Content
                range_end.Collapse(0)
                range_end.InlineShapes.AddPicture(temp_image_path)

                if i < total_pages - 1:
                    range_end.InsertBreak(7)

                os.remove(temp_image_path)

            visio_doc.Close()

            if separate_files:
                output_name = os.path.splitext(filename)[0] + ".docx"
                output_path = os.path.join(visio_dir, "Converted_Files", output_name)
                os.makedirs(os.path.dirname(output_path), exist_ok=True)
                current_doc.SaveAs(output_path)
                current_doc.Close()

        if not separate_files:
            output_word_path = os.path.join(visio_dir, "output.docx")
            doc.SaveAs(output_word_path)
        office_app.Quit()

    finally:
        pythoncom.CoUninitialize()

def visio_to_images(
    visio_dir,
    file_list,
    update_progress=None,
    image_format="PNG",
    word_processor="Word",
):
    """
    将Visio文件导出为图片到Converted_Files目录下

    参数:
        visio_dir (str): Visio文件所在目录路径
        file_list (list): 要转换的Visio文件名列表
        update_progress (function, 可选): 进度回调函数，格式为func(文件名, 当前序号, 总数)
        image_format (str): 导出的图片格式，支持"PNG"、"JPG"、"GIF"等Visio支持的格式
        word_processor (str): 保留参数，保持接口一致性，实际不使用

    返回:
        list: 生成的图片文件路径列表

    流程:
    1. 初始化COM环境
    2. 启动Visio应用程序
    3. 为每个Visio文件创建单独的子目录
    4. 将每页导出为指定格式的图片
    5. 返回所有生成的图片路径

    注意:
    - 图片会保存在visio_dir/Converted_Files/原文件名/目录下
    - 图片命名为"Page_1.png"、"Page_2.png"等形式
    - 使用前确保没有Visio进程运行(可调用kill_visio_processes)
    """
    pythoncom.CoInitialize()
    generated_files = []
    try:
        visio_app = win32com.client.Dispatch("Visio.Application")
        visio_app.Visible = False

        total_files = len(file_list)
        for idx, filename in enumerate(file_list):
            if update_progress:
                update_progress(filename, idx + 1, total_files)

            # 创建文件输出目录: visio_dir/Converted_Files/原文件名/
            output_dir = os.path.join(
                visio_dir, "Converted_Files", os.path.splitext(filename)[0]
            )
            os.makedirs(output_dir, exist_ok=True)

            # 打开Visio文件
            visio_file_path = os.path.normpath(os.path.join(visio_dir, filename))
            visio_doc = visio_app.Documents.Open(visio_file_path)

            # 导出每一页
            for i, page in enumerate(visio_doc.Pages):
                page_number = i + 1
                image_name = f"Page_{page_number}.{image_format.lower()}"
                image_path = os.path.join(output_dir, image_name)

                # 导出图片 (使用完整的导出方法确保质量)
                page.Export(image_path)
                generated_files.append(image_path)

            visio_doc.Close()

        return generated_files

    except Exception as e:
        print(f"导出图片时出错: {e}")
        return []
    finally:
        try:
            visio_app.Quit()
        except:
            pass
        pythoncom.CoUninitialize()

def get_visio_files(visio_dir, extensions=None, func=None):
    """
    获取指定目录下所有Visio文件

    参数:
        visio_dir (str): 要扫描的目录路径
        extensions (list, 可选): 要匹配的文件扩展名列表，默认包含 .vsdx/.vsd

    返回:
        list: 匹配到的Visio文件名列表(相对路径)
    """
    if extensions is None:
        extensions = [".vsdx", ".vsd"]

    try:
        all_files = [
            f
            for f in os.listdir(visio_dir)
            if os.path.isfile(os.path.join(visio_dir, f))
        ]

        file_list = [
            f for f in all_files if os.path.splitext(f.lower())[1] in extensions
        ]

        return sorted(file_list)  # 按名称排序保证处理顺序一致

    except Exception as e:
        print(f"扫描目录失败: {e}")
        return []

def run_visio_task(visio_dir, func, *args, **kwargs):
    """
    自动获取 file_list 并执行指定的 Visio 处理任务函数。

    参数:
        visio_dir (str): Visio 文件所在目录
        func (callable): 要执行的处理函数（如 visio_to_word_export_png）
        *args, **kwargs: 会透传给 func 的额外参数

    返回:
        func 返回值，或 None（如果无文件或出错）
    """
    kill_visio_processes()
    kill_word_processes("Word")
    file_list = get_visio_files(visio_dir)
    if not file_list:
        print(f"在目录 {visio_dir} 中未找到任何 Visio 文件。")
        return None

    return func(visio_dir, file_list, *args, **kwargs)



if __name__ == "__main__":
 
    # visio_dir = r"D:\Lenovo\Desktop\进迭时空SOP-v2\财务SOP\应付管理"
    # run_visio_task(visio_dir, visio_to_word_export_png, separate_files=True)

    # visio_dir = r"D:\Lenovo\Desktop\进迭时空SOP-v2\财务SOP\应收管理"
    # run_visio_task(visio_dir, visio_to_word_export_png, separate_files=True)

    # visio_dir = r"D:\Lenovo\Desktop\进迭时空SOP-v2\财务SOP\资产管理"
    # run_visio_task(visio_dir, visio_to_word_export_png, separate_files=True)

    # visio_dir = r"D:\Lenovo\Desktop\进迭时空SOP-v2\财务SOP\总账管理"
    # run_visio_task(visio_dir, visio_to_word_export_png, separate_files=True)

    # visio_dir = r"D:\Lenovo\Desktop\进迭时空SOP-v2\制造SOP\采购管理"
    # run_visio_task(visio_dir, visio_to_word_export_png, separate_files=True)

    # visio_dir = r"D:\Lenovo\Desktop\进迭时空SOP-v2\制造SOP\委外仓库管理"
    # run_visio_task(visio_dir, visio_to_word_export_png, separate_files=True)

    # visio_dir = r"D:\Lenovo\Desktop\进迭时空SOP-v2\制造SOP\物料BOM"
    # run_visio_task(visio_dir, visio_to_word_export_png, separate_files=True)

    # visio_dir = r"D:\Lenovo\Desktop\进迭时空SOP-v2\制造SOP\物料BOM-IC"
    # run_visio_task(visio_dir, visio_to_word_export_png, separate_files=True)

    visio_dir = r"D:\Lenovo\Desktop\进迭时空SOP-v2\制造SOP\销售管理"
    run_visio_task(visio_dir, visio_to_word_export_png, separate_files=True)