from PyQt6.QtWidgets import QApplication, QPushButton, QWidget, QVBoxLayout
from PyQt6.QtCore import QTimer


class MyWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.button = QPushButton('Click me')
        self.button.clicked.connect(self.on_button_clicked)
        self.click_timer = QTimer()
        self.click_timer.timeout.connect(self.reset_click)
        self.double_clicked = False

        layout = QVBoxLayout()
        layout.addWidget(self.button)
        self.setLayout(layout)

    def on_button_clicked(self):
        if not self.double_clicked:
            self.double_clicked = True
            self.click_timer.start(250)  # 设置双击间隔时间
            # 处理单击事件
            print('Button clicked once')
        else:
            print('Button was double clicked')

    def reset_click(self):
        self.double_clicked = False


app = QApplication([])
widget = MyWidget()
widget.show()
app.exec()