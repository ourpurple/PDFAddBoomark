import os
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import PyPDF2


def merge_pdfs_in_folder(folder_path, output_pdf_path):
    """
    合并文件夹中的所有PDF文件，文件名以“封面”开头的文件放在最前面。
    :param folder_path: 文件夹路径
    :param output_pdf_path: 合并后的PDF文件路径
    """
    # 获取文件夹中的所有PDF文件
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]

    # 如果没有PDF文件，则跳过合并
    if not pdf_files:
        log(f"警告：文件夹 '{folder_path}' 中没有 PDF 文件，跳过合并。")
        return False  # 返回 False 表示跳过合并

    pdf_writer = PyPDF2.PdfWriter()

    # 将文件名以“封面”开头的文件放在最前面
    cover_files = [f for f in pdf_files if f.startswith('封面')]
    other_files = [f for f in pdf_files if not f.startswith('封面')]
    sorted_files = cover_files + sorted(other_files)

    # 合并PDF文件
    for filename in sorted_files:
        pdf_path = os.path.join(folder_path, filename)
        log(f"正在合并文件: {filename}")
        pdf_reader = PyPDF2.PdfReader(pdf_path)
        for page in pdf_reader.pages:
            pdf_writer.add_page(page)

    # 保存合并后的PDF文件
    with open(output_pdf_path, 'wb') as output_pdf:
        pdf_writer.write(output_pdf)
    log(f"已保存合并后的PDF文件到: {output_pdf_path}")
    return True  # 返回 True 表示合并成功


def add_bookmarks(pdf_path, bookmarks):
    """
    将书签添加到PDF文件中。
    :param pdf_path: PDF文件路径
    :param bookmarks: 书签列表，每个书签是一个元组 (页码, 书签名称)
    """
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        pdf_writer = PyPDF2.PdfWriter()

        # 添加所有页面
        for page in pdf_reader.pages:
            pdf_writer.add_page(page)

        # 添加书签
        for bm_page_num, bookmark_name in bookmarks:
            if bm_page_num <= len(pdf_reader.pages):  # 检查页码是否有效
                pdf_writer.add_outline_item(bookmark_name, bm_page_num - 1)  # PDF页码从0开始
                log(f"正在处理文件 {os.path.basename(pdf_path)} 的第 {bm_page_num} 页，书签: {bookmark_name}")
            else:
                log(f"警告：页码 {bm_page_num} 超出范围，已跳过书签: {bookmark_name}")

        # 保存添加书签后的PDF文件
        with open(pdf_path, 'wb') as output_pdf:
            pdf_writer.write(output_pdf)
        log(f"已保存添加书签后的PDF文件到: {pdf_path}")


def process_folders(input_folder, output_folder):
    """
    处理输入文件夹中的所有子文件夹，合并PDF并添加书签。
    :param input_folder: 输入文件夹路径
    :param output_folder: 输出文件夹路径
    """
    # 检查输入文件夹是否存在
    if not os.path.exists(input_folder):
        log(f"错误：输入文件夹 '{input_folder}' 不存在！")
        return

    # 创建输出文件夹（如果不存在）
    os.makedirs(output_folder, exist_ok=True)

    # 获取表格中的书签信息
    bookmarks = []
    for child in tree.get_children():
        page_num = tree.item(child, "values")[0]
        bookmark_name = tree.item(child, "values")[1]
        try:
            page_num = int(page_num)
            bookmarks.append((page_num, bookmark_name))
        except ValueError:
            log(f"警告：页码 '{page_num}' 不是有效的整数，已跳过该行。")
            continue

    # 遍历输入文件夹中的所有子文件夹
    for root, dirs, files in os.walk(input_folder):
        for dir_name in dirs:
            subfolder_path = os.path.join(root, dir_name)
            log(f"正在处理子文件夹: {subfolder_path}")

            # 合并子文件夹中的所有PDF文件
            merged_pdf_name = f"{dir_name}_合并后的pdf.pdf"
            merged_pdf_path = os.path.join(output_folder, merged_pdf_name)

            # 如果合并成功，则添加书签
            if merge_pdfs_in_folder(subfolder_path, merged_pdf_path):
                add_bookmarks(merged_pdf_path, bookmarks)


def log(message):
    """
    将日志信息输出到日志窗口。
    :param message: 日志信息
    """
    log_text.insert(tk.END, message + "\n")
    log_text.see(tk.END)  # 自动滚动到最新日志
    root.update()  # 更新 GUI


def select_input_folder():
    """
    选择输入文件夹。
    """
    folder = filedialog.askdirectory()
    if folder:
        input_folder_var.set(folder)


def select_output_folder():
    """
    选择输出文件夹。
    """
    folder = filedialog.askdirectory()
    if folder:
        output_folder_var.set(folder)


def add_bookmark():
    """
    添加书签到表格中。
    """
    page_num = page_num_entry.get()
    bookmark_name = bookmark_name_entry.get()
    if page_num and bookmark_name:
        tree.insert("", "end", values=(page_num, bookmark_name))
        page_num_entry.delete(0, tk.END)
        bookmark_name_entry.delete(0, tk.END)
    else:
        messagebox.showwarning("警告", "页码和书签名不能为空！")


def remove_bookmark():
    """
    从表格中删除选中的书签。
    """
    selected_item = tree.selection()
    if selected_item:
        tree.delete(selected_item)
    else:
        messagebox.showwarning("警告", "请先选择一个书签！")


def start_processing():
    """
    开始处理。
    """
    input_folder = input_folder_var.get()
    output_folder = output_folder_var.get()

    if not input_folder:
        log("错误：请选择输入文件夹！")
        return

    log("开始处理...")
    process_folders(input_folder, output_folder)
    log("处理完成！")


# 创建主窗口
root = tk.Tk()
root.title("PDF 书签工具 By WanderInDoor V2025-0107")
root.geometry("900x700")  # 设置窗口大小为 900x600

# 设置主题样式
style = ttk.Style()
style.theme_use("clam")  # 使用现代主题

# 获取当前用户的桌面路径
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
output_folder_default = os.path.join(desktop_path, "PDF_Out")

# 如果默认输出文件夹不存在，则创建
os.makedirs(output_folder_default, exist_ok=True)

# 输入文件夹选择
input_folder_var = tk.StringVar()
input_frame = ttk.Frame(root)
input_frame.pack(pady=10, padx=10, fill=tk.X)

ttk.Label(input_frame, text="输入文件夹:").pack(side=tk.LEFT, padx=5)
ttk.Entry(input_frame, textvariable=input_folder_var, width=80).pack(side=tk.LEFT, padx=5)
ttk.Button(input_frame, text="选择", command=select_input_folder).pack(side=tk.LEFT, padx=5)

# 输出文件夹选择
output_folder_var = tk.StringVar(value=output_folder_default)  # 设置默认输出文件夹
output_frame = ttk.Frame(root)
output_frame.pack(pady=10, padx=10, fill=tk.X)

ttk.Label(output_frame, text="输出文件夹:").pack(side=tk.LEFT, padx=5)
ttk.Entry(output_frame, textvariable=output_folder_var, width=80).pack(side=tk.LEFT, padx=5)
ttk.Button(output_frame, text="选择", command=select_output_folder).pack(side=tk.LEFT, padx=5)

# 使用说明
instructions_frame = ttk.Frame(root)
instructions_frame.pack(pady=10, padx=10, fill=tk.X)

instructions_text = """
使用说明：
1、本工具会遍历并合并所选择的文件夹及子文件夹中的PDF文件，合并时以“封面”开头的PDF会合并至新PDF文件的开头。
2、合并后的PDF文件以父文件夹名命名，并添加后缀“合并后的pdf”。
3、会在合并后的PDF文件的对应页码添加自定义书签。
"""
ttk.Label(instructions_frame, text=instructions_text, justify=tk.LEFT).pack(side=tk.LEFT, padx=5)

# 表格和操作按钮
table_frame = ttk.Frame(root)
table_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

# 书签表格
columns = ("页码", "书签名")
tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=7)  # 设置表格高度为9行
tree.heading("页码", text="页码", anchor=tk.CENTER)
tree.heading("书签名", text="书签名", anchor=tk.CENTER)
tree.column("页码", anchor=tk.CENTER)
tree.column("书签名", anchor=tk.CENTER)
tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# 默认书签
default_bookmarks = [
    (1, "封面"),
    (2, "扉页"),
    (3, "版权"),
    (4, "目录"),
    (6, "正文")
]
for page_num, bookmark_name in default_bookmarks:
    tree.insert("", "end", values=(page_num, bookmark_name))

# 添加书签和删除书签的输入框和按钮
input_frame = ttk.Frame(table_frame)
input_frame.pack(side=tk.LEFT, padx=10, fill=tk.Y)

ttk.Label(input_frame, text="页码:").pack(pady=5)
page_num_entry = ttk.Entry(input_frame)
page_num_entry.pack(pady=5)

ttk.Label(input_frame, text="书签名:").pack(pady=5)
bookmark_name_entry = ttk.Entry(input_frame)
bookmark_name_entry.pack(pady=5)

ttk.Button(input_frame, text="添加书签", command=add_bookmark).pack(pady=5)
ttk.Button(input_frame, text="删除书签", command=remove_bookmark).pack(pady=5)

# 开始按钮
ttk.Button(root, text="开始处理", command=start_processing).pack(pady=10)

# 日志窗口
log_frame = ttk.Frame(root)
log_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

log_text = scrolledtext.ScrolledText(log_frame, width=80, height=20)
log_text.pack(fill=tk.BOTH, expand=True)

# 运行主循环
root.mainloop()