import sys
import os
from pathlib import Path

project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

from modules import funcs


def main():
    try:
        in_hrs, in_min, out_hrs, out_min, num, rev, desktop_path = sys.argv[1:8]
        in_hrs = int(in_hrs)
        in_min = int(in_min)
        out_hrs = int(out_hrs)
        out_min = int(out_min)
        num = int(num)
        rev = int(rev)

        time_data = funcs.generate(in_hrs, in_min, out_hrs, out_min, num, rev)
        funcs.write(Path(desktop_path), time_data)
    except Exception as e:
        print(f"Error in generate_worker: {e}")


if __name__ == "__main__":
    main()
