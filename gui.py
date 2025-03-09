import os
from order import ordertrans
import tkinter as tk
import logging
from tkinter import scrolledtext

# 配置日志记录器
logger = logging.getLogger()
logger.setLevel(logging.INFO)

# 配置日志
class TextHandler(logging.Handler):
    def __init__(self, text):
        logging.Handler.__init__(self)
        self.text = text

    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text.configure(state='normal')
            self.text.insert(tk.END, msg + '\n')
            self.text.configure(state='disabled')
            # 滚动到最新日志
            self.text.yview(tk.END)
        # 确保在主线程中更新 GUI
        self.text.after(0, append)

# 创建主窗口
root = tk.Tk()
root.title("订单转换")
root.geometry("800x600")

# 创建一个滚动文本框用于显示日志
log_text = scrolledtext.ScrolledText(root, state='disabled', height=20)
log_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

# 创建日志处理器并添加到记录器
text_handler = TextHandler(log_text)
text_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logger.addHandler(text_handler)

# 创建一个标签
label = tk.Label(root, text="欢迎使用订单转换程序!")
label.pack(pady=10)

# 创建一个按钮
def on_button_click():
    label.config(text="转换中......")
    # 定义 input 目录路径
    input_dir = 'input'

    # 初始化一个空列表用于存储找到的 Excel 文件路径
    excel_files = []

    # 遍历 input 目录及其子目录
    for base, dirs, files in os.walk(input_dir):
        for file in files:
            # 检查文件扩展名是否为 .xlsx 或 .xls
            if file.endswith(('.xlsx', '.xls')):
                # 构建文件的完整路径
                file_path = os.path.join(base, file)
                # 将符合条件的文件路径添加到列表中
                excel_files.append(file_path)

    # 打印找到的 Excel 文件路径
    for file_path in excel_files:
        ordertrans(file_path)

    label.config(text="转换完成，请到output目录下查看。")

button = tk.Button(root, text="转换input目录下的订单", command=on_button_click)
button.pack(pady=10)

# 进入主事件循环
root.mainloop()