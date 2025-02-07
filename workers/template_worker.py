import sys
import os

# 计算项目根目录路径（假设 workers 目录在项目根目录下）
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)  # 添加到 Python 路径

# 然后导入其他模块
from modules.split import copy, run


# ... 原有代码 ...

def main():
    copy()
    run('./temp/template.xlsx')


if __name__ == "__main__":
    main()
