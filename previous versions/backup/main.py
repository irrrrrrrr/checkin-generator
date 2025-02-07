import sys
import os
import yaml
import winreg
from pathlib import Path
from PyQt6 import QtWidgets, QtCore
from gui.gui import Ui_MainWindow

# 配置常量
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.yaml")
USER_KEY = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders')
DESKTOP_PATH = winreg.QueryValueEx(USER_KEY, "Desktop")[0]
CLICK_DELAY = 1000  # 双击间隔时间 (ms)
PROCESS_CLOSE_DELAY = 1000  # 任务完成后延迟关闭进程的时间 (ms)


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # 初始化状态
        self.processes = {}  # 存储活跃进程 {task_type: QProcess}
        self.double_clicked = False
        self.click_timer = QtCore.QTimer()
        self.click_timer.setSingleShot(True)
        self.click_timer.timeout.connect(self._reset_click_state)

        # 初始化配置
        self._load_config()

        # 连接信号
        self.ui.generate.clicked.connect(lambda: self._handle_task("generate"))
        self.ui.split_off_time.clicked.connect(lambda: self._handle_task("split"))
        self.ui.open_template.clicked.connect(lambda: self._handle_task("template"))
        self.ui.remove_temp.clicked.connect(self._handle_remove_temp)

    def _handle_task(self, task_type):
        """通用任务处理入口"""
        if self.double_clicked:
            return

        self.double_clicked = True
        self.click_timer.start(CLICK_DELAY)

        # 构造参数
        args = []
        if task_type == "generate":
            in_time = self.ui.in_time.time().toString("HH:mm").split(":")
            out_time = self.ui.out_time.time().toString("HH:mm").split(":")
            args = [
                in_time[0], in_time[1],  # 上班时间
                out_time[0], out_time[1],  # 下班时间
                str(self.ui.num.value()),  # 人数
                str(self.ui.rev.value()),  # 预留值
                DESKTOP_PATH  # 桌面路径
            ]
        elif task_type == "split":
            args = [DESKTOP_PATH]

        self._start_worker_process(task_type, args)

    def _start_worker_process(self, task_type, args):
        """启动工作进程"""
        if task_type in self.processes:
            self.ui.statusbar.showMessage("该任务已在运行中...", 3000)
            return

        process = QtCore.QProcess()
        process.setProgram("python")
        process.setArguments([f"./workers/{task_type}_worker.py"] + args)

        # 连接信号
        process.readyReadStandardOutput.connect(
            lambda: self._log_output(process, task_type)
        )
        process.readyReadStandardError.connect(
            lambda: self._log_error(process, task_type)
        )
        process.finished.connect(
            lambda exit_code, _: self._on_process_finished(task_type, exit_code)
        )
        process.errorOccurred.connect(
            lambda: self._on_process_error(task_type)
        )

        process.start()
        self.processes[task_type] = process
        self.ui.statusbar.showMessage(f"{task_type} 任务已启动...", 2000)

    def _on_process_finished(self, task_type, exit_code):
        """进程完成处理"""
        process = self.processes.get(task_type)
        if not process:
            return

        # 显示完成状态
        msg = "成功" if exit_code == 0 else f"异常退出 (代码: {exit_code})"
        self.ui.statusbar.showMessage(f"{task_type} 任务完成 - {msg}", 5000)

        # 延迟关闭进程
        QtCore.QTimer.singleShot(PROCESS_CLOSE_DELAY, lambda: self._close_process(task_type))
        self.ui.statusbar.showMessage(f"{task_type} 任务关闭 - {msg}", 5000)

    def _close_process(self, task_type):
        """关闭指定进程"""
        process = self.processes.pop(task_type, None)
        if process:
            process.terminate()
            process.waitForFinished()  # 确保进程终止
            process.deleteLater()  # 释放资源
            print(f"[{task_type}] 进程已关闭")

    def _on_process_error(self, task_type):
        """进程错误处理"""
        process = self.processes.get(task_type)
        if not process:
            return

        error_msg = process.errorString()
        self.ui.statusbar.showMessage(f"{task_type} 任务错误: {error_msg}", 8000)
        print(f"[进程错误] {task_type}: {error_msg}")
        process.terminate()

    def _log_output(self, process, task_type):
        """捕获标准输出"""
        try:
            raw_data = process.readAllStandardOutput().data()
            text = raw_data.decode("gbk").strip()
            if text:
                print(f"[{task_type}] {text}")
        except UnicodeDecodeError:
            print(f"[{task_type}] 输出解码失败")

    def _log_error(self, process, task_type):
        """捕获错误输出"""
        try:
            raw_data = process.readAllStandardError().data()
            text = raw_data.decode("gbk").strip()
            if text:
                print(f"[{task_type} ERROR] {text}")
                self.ui.statusbar.showMessage(f"错误: {text}", 8000)
        except UnicodeDecodeError:
            print(f"[{task_type}] 错误信息解码失败")

    def _reset_click_state(self):
        """重置双击状态"""
        self.double_clicked = False

    def _handle_remove_temp(self):
        """直接处理临时文件删除"""
        try:
            temp_file = os.path.join(os.path.dirname(__file__), "temp", "template.xlsx")
            if os.path.exists(temp_file):
                os.remove(temp_file)
                self.ui.statusbar.showMessage("临时文件已清除", 3000)
            else:
                self.ui.statusbar.showMessage("没有找到临时文件", 3000)
        except Exception as e:
            self.ui.statusbar.showMessage(f"清除失败: {str(e)}", 5000)

    # region 配置文件操作
    def _load_config(self):
        """加载配置文件"""
        try:
            with open(CONFIG_FILE, "r") as f:
                config = yaml.safe_load(f) or {}
                # 加载时间设置
                self.ui.in_time.setTime(
                    QtCore.QTime.fromString(config.get("in_time", "08:00"), "HH:mm")
                )
                self.ui.out_time.setTime(
                    QtCore.QTime.fromString(config.get("out_time", "17:30"), "HH:mm")
                )
                self.ui.num.setValue(config.get("num", 1))
                self.ui.rev.setValue(config.get("rev", 4))

                # 修复窗口位置和大小加载
                geometry = config.get("geometry", {})
                x = geometry.get("x", 100)
                y = geometry.get("y", 100)
                width = geometry.get("width", 800)
                height = geometry.get("height", 600)
                self.setGeometry(x, y, width, height)

        except FileNotFoundError:
            pass

    def closeEvent(self, event):
        """窗口关闭事件处理"""
        # 终止所有进程
        for name, process in self.processes.items():
            if process.state() == QtCore.QProcess.ProcessState.Running:
                process.terminate()

        # 保存配置
        config = {
            "in_time": self.ui.in_time.time().toString("HH:mm"),
            "out_time": self.ui.out_time.time().toString("HH:mm"),
            "num": self.ui.num.value(),
            "rev": self.ui.rev.value(),
            "geometry": {
                "x": self.x(),
                "y": self.y(),
                "width": self.width(),
                "height": self.height()
            }
        }
        with open(CONFIG_FILE, "w") as f:
            yaml.safe_dump(config, f)

        event.accept()
    # endregion


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())