import os
import tkinter as tk
import logging
from tkinter import scrolledtext
from main import main

version = 0.1
logger = logging.getLogger()


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
root.title("订单转换v%s" % version)
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
    main()
    label.config(text="转换完成，请查看output目录。")


button = tk.Button(root, text="转换input目录下的订单", command=on_button_click)
button.pack(pady=10)

# 进入主事件循环
root.mainloop()
