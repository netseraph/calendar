"""主程序入口"""

from __future__ import annotations
import sys
from datetime import datetime

from PySide6.QtWidgets import (
    QAbstractButton,
    QApplication,
    QDialog,
    QDialogButtonBox,
    QVBoxLayout,
)


def get_time_str() -> str:
    """获取时间"""
    return datetime.now().isoformat(sep=" ")


def my_accepted():
    """接受"""
    print(f"{get_time_str()} accepted\n")


def my_rejected():
    """拒绝"""
    print(f"{get_time_str()} rejected\n")


def my_help_requested():
    """帮助"""
    print(f"{get_time_str()} help requested\n")


class MyDialogButtonWindow(QDialog):
    """对话框窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Dialog Buttons")

        self.dialog_button_flags = QDialogButtonBox.StandardButton.Ok
        # 设置按钮的标志
        for button in QDialogButtonBox.StandardButton:
            if button == QDialogButtonBox.StandardButton.NoButton:
                continue
            self.dialog_button_flags |= button
        # 这里需要显式的生成按钮实例
        self.dialog_button = QDialogButtonBox(self.dialog_button_flags)

        # 点击按钮时设置对应的槽函数，在槽函数中打印按钮名称
        self.dialog_button.clicked.connect(self.clicked)
        self.dialog_button.accepted.connect(my_accepted)
        self.dialog_button.rejected.connect(my_rejected)
        self.dialog_button.helpRequested.connect(my_help_requested)

        # 设置按钮的布局
        self.v_layout = QVBoxLayout()
        self.v_layout.addWidget(self.dialog_button)
        self.setLayout(self.v_layout)

        # 当 dialog 对话框退出时其返回的状态信息
        self.finished.connect(self.my_finished)

    def clicked(self, button: QAbstractButton):
        """按钮事件"""
        # 获取按钮名称，这里一定要调用 standardButton(xxx) 方法，而不是 StandardButton(xxx)
        print(
            f"{get_time_str()} clicked: "
            f"{self.dialog_button.standardButton(button)} "
            f"{self.dialog_button.buttonRole(button)}"
        )

    def my_finished(self, result: int):
        """关闭事件"""
        if result == QDialog.DialogCode.Accepted:
            print(
                f"{get_time_str()} QDialog Window Finished: {result}(QDialog.DialogCode.Accepted)"
            )
        elif result == QDialog.DialogCode.Rejected:
            print(
                f"{get_time_str()} QDialog Window Finished: {result}(QDialog.DialogCode.Rejected)"
            )
        else:
            print(
                f"{get_time_str()} QDialog Window Finished: {result}(Unknown QDialogCode)"
            )
        self.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ins = MyDialogButtonWindow()
    ins.setWindowTitle("对话框程序")
    ins.exec()  # 模态对话框开启独立的事件循环，此时不需要单独开启 app.exec() 事件循环
    sys.exit(0)
