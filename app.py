from __future__ import annotations

import marshal
import sys
from pathlib import Path
from types import CodeType

# These imports are intentionally explicit: the recovered core imports some of
# them dynamically, while PyInstaller needs to see them during analysis.
import docx  # noqa: F401
import fitz  # noqa: F401
import pythoncom  # noqa: F401
import win32com.client  # noqa: F401
from PIL import Image, ImageColor, ImageDraw, ImageFont  # noqa: F401
from PyQt5.QtCore import QRect, Qt  # noqa: F401
from PyQt5.QtGui import QColor, QFont, QIcon, QPainter, QPen, QPixmap
from PyQt5.QtWidgets import QApplication, QLabel, QVBoxLayout, QWidget


APP_NAME = "PDF 工具箱"


def resource_path(relative: str) -> Path:
    """Resolve files next to the source or inside a PyInstaller bundle."""
    root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return root / relative


def load_core() -> dict[str, object]:
    """Load the preserved processing core from the original packaged app."""
    payload = resource_path("app.marshalled").read_bytes()
    code = marshal.loads(payload)
    if not isinstance(code, CodeType):
        raise RuntimeError("PDF 工具核心文件无效")

    namespace: dict[str, object] = {
        "__name__": "pdf_toolbox_core",
        "__file__": str(resource_path("app.py")),
    }
    exec(code, namespace)
    return namespace


class WhiteboardHeader(QWidget):
    """A plain white header with no character artwork or branded imagery."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("whiteboardHeader")
        self.setFixedHeight(112)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 18)
        layout.addStretch()

        title = QLabel(APP_NAME)
        title.setObjectName("heroTitle")
        subtitle = QLabel("合并 · 页面处理 · 转换 · 水印")
        subtitle.setObjectName("heroSubtitle")
        layout.addWidget(title)
        layout.addWidget(subtitle)

    def paintEvent(self, event) -> None:  # noqa: N802 - Qt API name
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.fillRect(self.rect(), QColor("#ffffff"))

        # A restrained document mark keeps the header useful without adding a
        # new theme illustration.
        right = self.width() - 40
        top = 23
        painter.setPen(QPen(QColor("#d7dce4"), 2))
        painter.setBrush(QColor("#fafbfc"))
        painter.drawRoundedRect(QRect(right - 54, top, 44, 58), 5, 5)
        painter.setPen(QPen(QColor("#b8c0cc"), 2))
        for offset in (15, 25, 35):
            painter.drawLine(right - 45, top + offset, right - 18, top + offset)

        painter.setPen(QPen(QColor("#e5e7eb"), 1))
        painter.drawLine(0, self.height() - 1, self.width(), self.height() - 1)
        painter.end()
        super().paintEvent(event)


WHITEBOARD_STYLE = """
QMainWindow, QWidget {
    background: #f6f7f9;
    color: #202124;
}
#sidebar {
    background: #ffffff;
    border-right: 1px solid #e2e5e9;
}
#brand {
    color: #1f2937;
    font-size: 20px;
    font-weight: 700;
    padding: 8px;
    background: transparent;
}
#navButton {
    color: #4b5563;
    text-align: left;
    border: 0;
    border-radius: 8px;
    padding: 13px 14px;
    background: transparent;
}
#navButton:hover {
    background: #f1f3f5;
    color: #111827;
}
#navButton:checked {
    background: #e8f0fe;
    color: #1d4ed8;
    font-weight: 600;
}
#sideTip {
    color: #8a919c;
    padding: 10px;
    line-height: 1.5;
    background: transparent;
}
#whiteboardHeader {
    background: #ffffff;
}
#heroTitle {
    color: #1f2937;
    font-size: 26px;
    font-weight: 700;
    background: transparent;
}
#heroSubtitle {
    color: #6b7280;
    font-size: 13px;
    background: transparent;
}
#pageTitle {
    font-size: 23px;
    font-weight: 700;
    color: #1f2937;
}
#pageSubtitle {
    color: #747b86;
    padding-bottom: 4px;
}
QGroupBox {
    border: 1px solid #dfe3e8;
    border-radius: 10px;
    margin-top: 12px;
    padding: 14px 10px 10px;
    font-weight: 600;
    background: #ffffff;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 6px;
    color: #374151;
}
#pageScroll, #pageScroll > QWidget > QWidget {
    background: transparent;
}
QListWidget, QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {
    background: #ffffff;
    border: 1px solid #d6dbe1;
    border-radius: 7px;
    padding: 5px 7px;
    selection-background-color: #2563eb;
}
QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {
    min-height: 24px;
}
QListWidget {
    padding: 6px;
}
QListWidget::item {
    padding: 7px;
    border-radius: 5px;
}
QListWidget::item:selected {
    background: #dbeafe;
    color: #1e3a8a;
}
QPushButton {
    background: #ffffff;
    border: 1px solid #d4d9df;
    border-radius: 7px;
    padding: 7px 14px;
    min-height: 22px;
}
QPushButton:hover {
    border-color: #7aa2e3;
    background: #f3f7fd;
}
QPushButton:pressed {
    background: #e7eef9;
}
#primaryButton {
    background: #2563eb;
    color: #ffffff;
    border-color: #2563eb;
    font-weight: 600;
}
#primaryButton:hover {
    background: #1d4ed8;
}
QProgressBar {
    border: 0;
    background: #e5e7eb;
    border-radius: 4px;
    height: 8px;
}
QProgressBar::chunk {
    background: #2563eb;
    border-radius: 4px;
}
"""


def apply_whiteboard_theme(core: dict[str, object]) -> None:
    """Replace only the original presentation layer, preserving all tools."""
    core["APP_NAME"] = APP_NAME
    core["ThemeHeader"] = WhiteboardHeader

    main_window = core["MainWindow"]
    original_build_ui = main_window._build_ui

    def build_ui(self) -> None:
        original_build_ui(self)
        brand = self.findChild(QLabel, "brand")
        if brand is not None:
            brand.setText(APP_NAME)

    def apply_style(self) -> None:
        app = QApplication.instance()
        if app is not None:
            app.setFont(QFont("Microsoft YaHei UI", 10))
        self.setStyleSheet(WHITEBOARD_STYLE)

    main_window._build_ui = build_ui
    main_window._apply_style = apply_style


def main() -> int:
    core = load_core()
    apply_whiteboard_theme(core)
    return int(core["main"]())


if __name__ == "__main__":
    raise SystemExit(main())
