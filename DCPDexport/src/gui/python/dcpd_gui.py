import sys
import os
import threading
from pathlib import Path

# ensure parent of the DCPDexport package exists on sys.path so it can be imported
# the GUI script lives under DCPDexport/src/gui/python; adding the parent directory
# (workspace root) lets Python locate the top-level DCPDexport package.
package_parent = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", ".."))
if package_parent not in sys.path:
    sys.path.insert(0, package_parent)

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QTextEdit, QFileDialog, QMessageBox,
    QRadioButton, QButtonGroup, QProgressBar
)
from PySide6.QtCore import Qt, Signal, QObject, QEvent
from PySide6.QtGui import QTextCursor

# import processing functions from shared core
from DCPDexport.src.core import write_output


class OutputRedirector(QObject):
    """Redirect stdout/stderr to a QTextEdit."""
    output_written = Signal(str)

    def write(self, text):
        if text:
            self.output_written.emit(text)

    def flush(self):
        pass


class ProcessCompleteEvent(QEvent):
    EVENT_TYPE = QEvent.Type(QEvent.registerEventType())

    def __init__(self, success: bool):
        super().__init__(self.EVENT_TYPE)
        self.success = success


class DataProcessorGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DCPD 数据导出工具")
        self.setMinimumSize(800, 600)
        self.processing = False

        # redirect output
        self.redirector = OutputRedirector()
        self.redirector.output_written.connect(self.append_output)
        sys.stdout = self.redirector
        sys.stderr = self.redirector

        self.setup_ui()

    def setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(10)

        # file selector
        file_layout = QHBoxLayout()
        file_label = QLabel("数据文件：")
        file_layout.addWidget(file_label)
        self.file_input = QLineEdit()
        file_layout.addWidget(self.file_input, 1)
        browse = QPushButton("浏览...")
        browse.clicked.connect(self.browse_file)
        file_layout.addWidget(browse)
        layout.addLayout(file_layout)

        # time unit
        unit_layout = QHBoxLayout()
        unit_label = QLabel("时间单位：")
        unit_layout.addWidget(unit_label)
        self.sec_radio = QRadioButton("秒 (sec)")
        self.sec_radio.setChecked(True)
        unit_layout.addWidget(self.sec_radio)
        self.hr_radio = QRadioButton("小时 (hr)")
        unit_layout.addWidget(self.hr_radio)
        layout.addLayout(unit_layout)

        # process button
        self.process_btn = QPushButton("开始处理")
        self.process_btn.clicked.connect(self.start_processing)
        layout.addWidget(self.process_btn, alignment=Qt.AlignCenter)

        # progress bar
        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.hide()
        layout.addWidget(self.progress)

        # output text
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log, 1)

    def browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择数据文件", "", "文本文件 (*.txt *.csv);;所有文件 (*.*)"
        )
        if path:
            self.file_input.setText(path)
            print(f"[info] 已选择文件: {path}")

    def start_processing(self):
        if self.processing:
            QMessageBox.warning(self, "警告", "正在处理中，请稍候")
            return
        file_path = self.file_input.text().strip()
        if not file_path:
            QMessageBox.warning(self, "警告", "请先选择数据文件")
            return
        if not os.path.isfile(file_path):
            QMessageBox.critical(self, "错误", "文件不存在")
            return
        unit = "sec" if self.sec_radio.isChecked() else "hr"
        self.processing = True
        self.process_btn.setEnabled(False)
        self.progress.show()
        self.progress.setRange(0, 0)
        thread = threading.Thread(target=self._worker, args=(file_path, unit))
        thread.daemon = True
        thread.start()

    def _worker(self, file_path: str, unit: str):
        try:
            write_output(Path(file_path), unit, None, ask_overwrite=True)
            success = True
        except Exception as e:
            print(f"[warn] 处理失败: {e}")
            success = False
        QApplication.instance().postEvent(self, ProcessCompleteEvent(success))

    def processing_complete(self, success: bool):
        self.progress.hide()
        self.processing = False
        self.process_btn.setEnabled(True)
        if success:
            QMessageBox.information(self, "完成", "处理完成")

    def customEvent(self, event):
        if isinstance(event, ProcessCompleteEvent):
            self.processing_complete(event.success)

    def append_output(self, text: str):
        # ensure cursor at end before inserting
        self.log.moveCursor(QTextCursor.End)
        self.log.insertPlainText(text)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = DataProcessorGUI()
    win.show()
    sys.exit(app.exec())
