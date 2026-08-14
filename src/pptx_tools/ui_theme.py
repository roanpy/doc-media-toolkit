"""Shared visual contract for repeated desktop controls."""

from PySide6.QtCore import QEvent, QObject
from PySide6.QtGui import QFont, QFontDatabase
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QProxyStyle,
    QPushButton,
    QStyle,
    QWidget,
)

from pptx_tools.manager_i18n import tr

UI_FONT_CANDIDATES = (
    "PingFang SC",
    "Hiragino Sans GB",
    "SF Pro Text",
    "Helvetica Neue",
    "Microsoft YaHei",
    "Segoe UI",
    "Noto Sans CJK SC",
    "Noto Sans SC",
    "Arial",
)


class DialogCenteringFilter(QObject):
    def eventFilter(self, watched, event):  # noqa: N802, ANN001
        if isinstance(watched, QDialog) and event.type() == QEvent.Type.Show:
            apply_shared_dialog_contract(watched)
            self._center(watched)
        return super().eventFilter(watched, event)

    @staticmethod
    def _center(dialog: QDialog) -> None:
        if not dialog.isVisible() or dialog.isMaximized():
            return
        parent = dialog.parentWidget()
        anchor = parent.window() if parent is not None else None
        screen = dialog.screen() or QApplication.primaryScreen()
        if anchor is None and screen is None:
            return
        area = (
            anchor.frameGeometry() if anchor is not None else screen.availableGeometry()
        )
        dialog.move(
            area.x() + (area.width() - dialog.width()) // 2,
            area.y() + (area.height() - dialog.height()) // 2,
        )


class DelayedTooltipStyle(QProxyStyle):
    """Keep incidental pointer movement from opening help immediately."""

    def styleHint(self, hint, option=None, widget=None, returnData=None):  # noqa: N802, ANN001
        if hint == QStyle.StyleHint.SH_ToolTip_WakeUpDelay:
            return 1000
        return super().styleHint(hint, option, widget, returnData)


def configure_tooltip_behavior(app: QApplication) -> None:
    if app.property("_ppt_tools_tooltips_configured"):
        return
    app.setStyle(DelayedTooltipStyle(app.style()))
    center_filter = DialogCenteringFilter(app)
    app.installEventFilter(center_filter)
    app._ppt_tools_dialog_center_filter = center_filter
    app.setProperty("_ppt_tools_tooltips_configured", True)


def configure_ui_font(app: QApplication) -> None:
    """Apply the shared desktop font once for every tool surface."""

    configure_tooltip_behavior(app)
    if app.property("_ppt_tools_ui_font_configured"):
        return
    families = set(QFontDatabase.families())
    for family in UI_FONT_CANDIDATES:
        if family in families:
            font = QFont(app.font())
            font.setFamily(family)
            app.setFont(font)
            break
    app.setProperty("_ppt_tools_ui_font_configured", True)


def install_control_help(root: QWidget) -> None:
    """Fill only missing help; keep purpose-written tooltips intact."""

    for button in root.findChildren(QPushButton):
        text = button.text().strip().replace("&", "")
        if text and not button.toolTip():
            button.setToolTip(tr("点击执行“{}”。").format(text))
    for combo in root.findChildren(QComboBox):
        if combo.toolTip():
            continue

        def update_combo_help(_index: int = -1, control: QComboBox = combo) -> None:
            control.setToolTip(
                tr("当前选择：{}。点击切换其他选项。").format(control.currentText())
            )

        update_combo_help()
        combo.currentIndexChanged.connect(update_combo_help)


def apply_shared_dialog_contract(dialog: QDialog) -> None:
    """Apply the shared dialog baseline once without replacing local styling."""

    if dialog.property("_ppt_tools_shared_dialog"):
        return
    dialog.setStyleSheet(f"{dialog.styleSheet()}\n{SHARED_DIALOG_QSS}")
    install_control_help(dialog)
    dialog.setProperty("_ppt_tools_shared_dialog", True)


def format_user_file_size(size_bytes: int) -> str:
    """Return a compact file size without exposing implementation-level bytes."""
    size = max(0, int(size_bytes))
    if size == 0:
        return "0 KB"
    if size < 1024:
        return "< 1 KB"
    if size < 1024**2:
        return f"{size / 1024:.1f} KB"
    if size < 1024**3:
        return f"{size / 1024**2:.2f} MB"
    return f"{size / 1024**3:.2f} GB"


SHARED_MAIN_QSS = """
QMainWindow, QWidget#central {
    background: #07101a;
}
QFrame#headerCard, QFrame#leftCard, QFrame#sideCard,
QFrame#detailsCard, QFrame#queueCard, QFrame#previewCard,
QFrame#rightCard {
    background: #0d1926;
    border: 1px solid #294057;
    border-radius: 10px;
}
QLabel#title, QLabel#pageTitle {
    font-size: 16px;
}
QLabel#sectionTitle {
    font-size: 13px;
}
QPushButton {
    background: #091522;
    color: #dce7f3;
    border: 1px solid #294057;
    border-radius: 8px;
    font-size: 13px;
    padding: 0 10px;
}
QPushButton:hover {
    background: #122235;
    border-color: #55718c;
}
QPushButton:disabled {
    background: #32475f;
    color: #8ba0b4;
    border-color: #425a71;
}
QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {
    background: #091522;
    color: #dce7f3;
    border: 1px solid #294057;
    border-radius: 8px;
    font-size: 12px;
}
QHeaderView::section {
    background: #152536;
    color: #cbd5e1;
    border: 0;
    border-bottom: 1px solid #294057;
}
QScrollBar:vertical, QScrollBar:horizontal {
    background: transparent;
    border: 0;
}
QScrollBar:vertical { width: 8px; }
QScrollBar:horizontal { height: 8px; }
QScrollBar::handle:vertical, QScrollBar::handle:horizontal {
    background: #3d566e;
    border-radius: 4px;
    min-height: 24px;
    min-width: 24px;
}
QScrollBar::handle:vertical:hover, QScrollBar::handle:horizontal:hover {
    background: #58748e;
}
QScrollBar::add-line, QScrollBar::sub-line,
QScrollBar::add-page, QScrollBar::sub-page {
    background: transparent;
    border: 0;
    width: 0;
    height: 0;
}
QLabel#estimatePill {
    border-radius: 9px;
    padding: 2px 10px;
    font-size: 12px;
}
QPushButton#helpIconButton {
    border-radius: 15px;
    min-width: 30px;
    max-width: 30px;
    min-height: 30px;
    max-height: 30px;
}
"""

SHARED_DIALOG_QSS = """
QDialog {
    background: #07101a;
}
QDialog QLabel#dialogTitle {
    font-size: 16px;
}
QDialog QPushButton {
    background: #091522;
    color: #dce7f3;
    border: 1px solid #294057;
    border-radius: 8px;
    font-size: 13px;
    min-height: 30px;
    min-width: 72px;
    padding: 0 18px;
}
QDialog QPushButton:hover {
    background: #122235;
    border-color: #55718c;
}
QDialog QLineEdit, QDialog QComboBox, QDialog QSpinBox,
QDialog QDoubleSpinBox {
    background: #091522;
    color: #dce7f3;
    border: 1px solid #294057;
    border-radius: 8px;
    font-size: 12px;
}
QDialog QHeaderView::section {
    font-size: 11px;
}
"""
