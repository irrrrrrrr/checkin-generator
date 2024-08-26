import sys

sys.path.append('./modules')
import yaml
import os
from pathlib import Path
from PyQt6 import QtWidgets, QtCore
from gui.gui import Ui_MainWindow
from modules import split, funcs

CONFIG_FILE = 'config.yaml'
click_time = 1000
desktop_path = Path(os.path.expanduser("~")) / "Desktop"
print(f"Desktop path: {desktop_path}")


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        print("UI setup complete")
        self.load_config()
        print("Config loaded")
        self.thread_manager = QtCore.QThreadPool()
        self.click_timer = QtCore.QTimer()
        self.click_timer.timeout.connect(self.reset_click)
        self.double_clicked = False
        self.ui.generate.clicked.connect(self.handle_generate_safely)
        print("Generate_signals connected")
        self.ui.open_template.clicked.connect(self.handle_template_safely)
        print("Template_signals connected")
        self.ui.split_off_time.clicked.connect(self.handle_split_safely)
        print("Spliter_signals connected")
        self.ui.remove_temp.clicked.connect(self.handle_remove_temp)
        print("Remover_signals connected")

    @QtCore.pyqtSlot()
    def reset_click(self):
        self.double_clicked = False

    def load_config(self):
        try:
            with open(CONFIG_FILE, 'r') as file:
                config = yaml.safe_load(file)
                if config:
                    self.ui.in_time.setTime(QtCore.QTime.fromString(config.get('in_time', '00:00'), 'HH:mm'))
                    self.ui.out_time.setTime(QtCore.QTime.fromString(config.get('out_time', '00:00'), 'HH:mm'))
                    self.ui.num.setValue(config.get('num', 0))
                    self.ui.rev.setValue(config.get('rev', 0))
        except FileNotFoundError:
            print("Config file not found")

    def save_config(self):
        config = {
            'in_time': self.ui.in_time.time().toString('HH:mm'),
            'out_time': self.ui.out_time.time().toString('HH:mm'),
            'num': self.ui.num.value(),
            'rev': self.ui.rev.value()
        }
        with open(CONFIG_FILE, 'w') as file:
            yaml.safe_dump(config, file)
        print("Config saved")

    @QtCore.pyqtSlot()
    def handle_generate(self):
        try:
            in_time_value = self.ui.in_time.time()
            out_time_value = self.ui.out_time.time()
            num = int(self.ui.num.value())
            rev = int(self.ui.rev.value())
            # 将时间值转换为字符串以便显示
            in_time_str = in_time_value.toString('HH:mm')
            out_time_str = out_time_value.toString('HH:mm')
            in_hrs, in_min = in_time_str.split(':', 1)
            out_hrs, out_min = out_time_str.split(':', 1)
            in_hrs, in_min, out_hrs, out_min = int(in_hrs), int(in_min), int(out_hrs), int(out_min)
            time = funcs.generate(in_hrs, in_min, out_hrs, out_min, num, rev)
            funcs.write(desktop_path, time)
            print(f"上班时间: {in_time_str}")
            print(f"下班时间: {out_time_str}")
            print(f"人数: {num}")
            print(f"预留值: {rev}")
            self.save_config()
        except Exception as exp:
            print(f"Error in handle_generate: {exp}")

    @QtCore.pyqtSlot()
    def handle_split(self):
        try:
            print(f"Checking if temp_path exists: {os.path.exists('./temp/off_time.xlsx')}")
            split.process_excel_file(desktop_path)
            QtWidgets.QApplication.processEvents()
        except Exception as e:
            print(f"Error in process_excel_file: {e}")
            import traceback
            traceback.print_exc()

        try:
            if os.path.exists(f'{desktop_path}\\off_time.xlsx'):
                split.run(f'{desktop_path}\\off_time.xlsx')
            else:
                print(f"File not found: {desktop_path}\\off_time.xlsx")
        except Exception as e:
            print(f"Error in run_off_time: {e}")
            import traceback
            traceback.print_exc()

    @QtCore.pyqtSlot()
    def handle_template(self):
        try:
            split.copy()
            split.run('./temp/template.xlsx')
        except Exception as e:
            print(f"Error in handle_template: {e}")

    @QtCore.pyqtSlot()
    def handle_remove_temp(self):
        try:
            split.remove_temp()
        except Exception as e:
            print(f"Error in remove_temp: {e}")

    @QtCore.pyqtSlot()
    def handle_split_safely(self):
        if not self.double_clicked:
            self.double_clicked = True
            self.click_timer.start(click_time)  # 设置双击间隔时间
            self.thread_manager.start(self.handle_split)
            print('Button clicked once')
        else:
            print('Button was double clicked')

    @QtCore.pyqtSlot()
    def handle_generate_safely(self):
        if not self.double_clicked:
            self.double_clicked = True
            self.click_timer.start(click_time)  # 设置双击间隔时间
            self.thread_manager.start(self.handle_generate)
            print('Button clicked once')
        else:
            print('Button was double clicked')

    @QtCore.pyqtSlot()
    def handle_template_safely(self):
        if not self.double_clicked:
            self.double_clicked = True
            self.click_timer.start(click_time)  # 设置双击间隔时间
            self.thread_manager.start(self.handle_template)
            print('Button clicked once')
        else:
            print('Button was double clicked')


if __name__ == '__main__':
    try:
        app = QtWidgets.QApplication(sys.argv)
        mainWindow = MainWindow()
        mainWindow.show()
        sys.exit(app.exec())
    except Exception as e:
        print(f"Error in main: {e}")
