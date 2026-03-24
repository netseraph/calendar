"""字体测试"""

from typing import List

from PySide6.QtGui import QFont, QFontDatabase
from PySide6.QtWidgets import (
    QApplication,
    QLabel,
    QLayout,
    QMainWindow,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)


class MyMainWindow(QMainWindow):
    """窗口类"""

    def __init__(self):
        super().__init__()

        self.setWindowTitle("支持的字体类型")
        self.font_families = QFontDatabase.families()
        fonts_v_layout = self.init_fonts_layout(self.font_families)
        fonts_area = self.init_scroll_area(fonts_v_layout)
        self.setCentralWidget(fonts_area)

    def init_scroll_area(self, layout: QLayout) -> QScrollArea:
        """设置滚动条"""
        colors_container = QWidget()
        colors_container.setLayout(layout)
        colors_scroll_area = QScrollArea(self)
        colors_scroll_area.setWidgetResizable(True)
        colors_scroll_area.setWidget(colors_container)
        return colors_scroll_area

    def init_fonts_layout(self, fonts: List[str]) -> QVBoxLayout:
        """字体布局"""
        # count = 0
        v_colors_layout = QVBoxLayout()
        # h_colors_layout = QHBoxLayout()
        for c in fonts:
            print(type(c), c, type(c.capitalize()), c.capitalize())
            tmp_label = QLabel(c.capitalize(), parent=self)
            tmp_label.setFont(QFont(c, 10))
            v_colors_layout.addWidget(tmp_label)
            # count += 1
            # if count % 5 == 0:
            #    v_colors_layout.addLayout(h_colors_layout)
            #    h_colors_layout = QHBoxLayout()
            #    count = 0
        # if count > 0:
        #    v_colors_layout.addLayout(h_colors_layout)
        return v_colors_layout


if __name__ == "__main__":
    app = QApplication()
    ins = MyMainWindow()
    ins.show()
    app.exec()
