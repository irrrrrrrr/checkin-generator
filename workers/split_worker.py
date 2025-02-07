import sys
import os
from pathlib import Path

project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

from modules.split import process_excel_file


def main():
    desktop_path = Path(sys.argv[1])
    process_excel_file(desktop_path)


if __name__ == "__main__":
    main()
