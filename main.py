import os
from order import ordertrans
import logging

# 配置日志记录器
logger = logging.getLogger()
logger.setLevel(logging.INFO)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def main():
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


if __name__ == '__main__':
    main()
