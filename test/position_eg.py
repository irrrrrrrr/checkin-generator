import sys
from PyQt6.QtWidgets import QApplication, QMainWindow
from PyQt6.QtCore import QRect
import yaml

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # 读取窗口位置和大小
        self.load_window_settings()

    def closeEvent(self, event):
        # 保存窗口位置和大小
        self.save_window_settings()
        event.accept()

    def load_window_settings(self):
        try:
            with open('window_settings.yaml', 'r') as file:
                settings = yaml.safe_load(file)
                rect = settings['geometry']
                self.setGeometry(QRect(rect['x'], rect['y'], rect['width'], rect['height']))
        except FileNotFoundError:
            # 如果文件不存在，则使用默认窗口大小和位置
            self.setGeometry(100, 100, 800, 600)

    def save_window_settings(self):
        settings = {
            'geometry': {
                'x': self.geometry().x(),
                'y': self.geometry().y(),
                'width': self.geometry().width(),
                'height': self.geometry().height(),
            }
        }
        with open('window_settings.yaml', 'w') as file:
            yaml.dump(settings, file)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
