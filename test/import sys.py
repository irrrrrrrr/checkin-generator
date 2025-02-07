import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog
 
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("File Dialog Example")
        self.show()
 
    def open_file(self):
        filename, _ = QFileDialog.getOpenFileName(self, "Open File", "./", "All Files (*)")
        if filename:
            print(f"Selected File: {filename}")
 
def main():
    app = QApplication(sys.argv)
    mainWin = MainWindow()
    mainWin.open_file()
    sys.exit(app.exec())
 
if __name__ == '__main__':
    main()