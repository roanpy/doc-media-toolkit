from __future__ import annotations

import logging
import threading
from collections.abc import Callable
from typing import Any

from PySide6.QtCore import QObject, Signal, Slot

from pptx_tools.ui_theme import SHARED_DIALOG_QSS, SHARED_MAIN_QSS


LOGGER = logging.getLogger("pptx_tools.media_manager_ui")

MEDIA_MANAGER_STYLESHEET = (
    """
QMainWindow, QDialog, QWidget { background: #0b1017; color: #cbd5e1; font-size: 12px; }
QFrame#headerCard { background: #111827; border: 1px solid #334155; border-radius: 10px; }
QFrame#projectBar { background: #0f1720; border: 1px solid #273244; border-radius: 8px; }
QFrame#headerCard QLabel, QLabel#pageSubtitle, QLabel#projectPath { background: transparent; }
QLabel#pageTitle { color: #f8fafc; font-size: 17px; font-weight: 600; }
QLabel#pageSubtitle { color: #94a3b8; font-size: 11px; }
QLabel#projectPath { color: #94a3b8; font-size: 11px; }
QLabel#statMeta { background: transparent; color: #94a3b8; padding: 3px 3px; font-size: 11px; font-weight: 600; }
QLabel#inputSummary { background: #111827; border: 1px solid #273244; border-radius: 6px; padding: 4px 8px; color: #cbd5e1; font-size: 12px; }
QFrame#workflowFiles { background: #0f1720; border: 1px dashed #334155; border-radius: 6px; }
QFrame#workflowFiles QLabel { background: transparent; border: 0; color: #cbd5e1; font-size: 12px; padding: 3px 6px; }
QFrame#workflowFiles QPushButton#inputChip { background: #18212d; border: 1px solid #334155; border-radius: 5px; color: #f8fafc; padding: 4px 7px; font-size: 12px; }
QFrame#workflowSettings { background: #0f1720; border: 1px solid #273244; border-radius: 6px; }
QFrame#workflowSettings QComboBox, QFrame#workflowSettings QLineEdit { min-height: 26px; }
QLabel#emptyState { background: #0f1720; border: 1px dashed #334155; border-radius: 8px; color: #94a3b8; padding: 20px; font-size: 12px; font-weight: 500; }
QFrame#videoDetailDrawer { background: #111827; border: 1px solid #334155; border-radius: 10px; }
QFrame#videoDetailDrawer QLabel { background: transparent; }
QLabel#detailTitle { color: #f8fafc; font-size: 16px; font-weight: 600; }
QLabel#detailMeta { color: #cbd5e1; font-size: 12px; }
QLabel#detailLabel { color: #94a3b8; font-size: 11px; font-weight: 600; }
QLabel#detailValue { color: #dbe4f0; font-size: 12px; }
QFrame#detailMetricCard { background: #0f1720; border: 1px solid #273244; border-radius: 6px; }
QFrame#detailMetricCard QLabel { border: 0; }
QLabel#detailMetricLabel { color: #94a3b8; font-size: 11px; font-weight: 600; }
QLabel#detailMetricValue { color: #f1f5f9; font-size: 12px; font-weight: 600; }
QLabel#detailStatus[healthy="true"] { color: #4ade80; }
QLabel#detailStatus[healthy="false"] { color: #fca5a5; }
QLabel#detailBadge { color: #fed7aa; background: #7c2d12; border: 1px solid #c2410c; border-radius: 5px; padding: 2px 6px; font-size: 11px; font-weight: 700; }
QLabel#detailPreview { background: #0b1017; border: 1px solid #334155; border-radius: 8px; color: #64748b; }
QDialog QLabel#dialogTitle { color: #f8fafc; font-size: 16px; font-weight: 600; }
QDialog QLabel#dialogSubtitle { color: #94a3b8; font-size: 11px; }
QDialog QLabel#dialogSummary { background: #111827; border: 1px solid #273244; border-radius: 6px; color: #cbd5e1; font-size: 12px; padding: 8px 10px; }
QDialog QLabel#previewHeading { color: #cbd5e1; font-size: 12px; font-weight: 600; }
QDialog QTreeWidget::item { min-height: 28px; }
QDialog QDialogButtonBox QPushButton {
    min-height: 28px;
    min-width: 78px;
    font-size: 11px;
    font-weight: 500;
}
QPushButton#dangerAction { background: #7f1d1d; border-color: #ef4444; color: white; }
QPushButton#dangerAction:hover { background: #991b1b; }
QPushButton#dangerAction:disabled { background: #111827; border-color: #27364a; color: #64748b; }
QGroupBox { border: 1px solid #334155; border-radius: 8px; margin-top: 10px; padding-top: 8px; font-weight: 600; }
QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; color: #fb923c; font-size: 14px; }
QTreeWidget, QPlainTextEdit, QComboBox, QLineEdit { background: #0f1720; border: 1px solid #334155; border-radius: 6px; padding: 4px; font-size: 12px; }
QTreeWidget:focus, QPlainTextEdit:focus, QComboBox:focus, QLineEdit:focus, QPushButton:focus { border: 1px solid #fb923c; }
QMenu { background: #111827; color: #cbd5e1; border: 1px solid #334155; padding: 4px; font-size: 12px; }
QMenu::item { border-radius: 5px; padding: 5px 22px 5px 8px; }
QMenu::item:selected { background: #243244; }
QMenu::item:disabled { color: #64748b; }
QMenu::separator { height: 1px; background: #334155; margin: 4px 6px; }
QTreeWidget::item { min-height: 28px; }
QTreeWidget::item:selected { background: #12385f; color: white; }
QHeaderView::section { background: #18212d; color: #cbd5e1; border: 0; padding: 6px 5px; font-size: 12px; font-weight: 600; }
QScrollBar:vertical, QScrollBar:horizontal { background: #0f1720; border: 0; border-radius: 4px; }
QScrollBar:vertical { width: 8px; margin: 2px 2px; }
QScrollBar:horizontal { height: 8px; margin: 2px 2px; }
QScrollBar::handle:vertical, QScrollBar::handle:horizontal { background: #334155; border-radius: 4px; min-height: 24px; min-width: 24px; }
QScrollBar::handle:vertical:hover, QScrollBar::handle:horizontal:hover { background: #475569; }
QScrollBar::add-line, QScrollBar::sub-line, QScrollBar::add-page, QScrollBar::sub-page { background: transparent; border: 0; width: 0; height: 0; }
QPushButton { background: #18212d; border: 1px solid #334155; border-radius: 6px; padding: 5px 12px; font-size: 12px; font-weight: 500; min-width: 0px; }
QPushButton::menu-indicator { image: none; width: 0px; height: 0px; margin: 0px; padding: 0px; }
QWidget#libraryActions QPushButton { padding: 5px 10px; font-size: 12px; }
QWidget#libraryActions[compact="true"] QPushButton { padding: 4px 6px; }
QPushButton#statFilter { background: #111827; border-color: #334155; padding: 4px 10px; font-size: 12px; }
QWidget#libraryActions[compact="true"] QPushButton#statFilter { padding: 4px 6px; }
QPushButton#statFilter:hover { background: #243244; border-color: #64748b; }
QPushButton#statFilter:checked { background: #c2410c; border-color: #fb923c; color: white; }
QPushButton:hover { background: #243244; }
QPushButton:disabled { color: #64748b; background: #111827; }
QPushButton#primaryAction { background: #f97316; border-color: #fb923c; color: white; padding: 5px 14px; font-weight: 600; }
QPushButton#primaryAction:hover { background: #ea580c; }
QPushButton#primaryAction:disabled { background: #111827; border-color: #27364a; color: #64748b; }
QPushButton#helpIconButton {
    background: #0f1720; color: #cbd5e1; border: 1px solid #334155;
    border-radius: 15px; padding: 0; min-width: 30px; max-width: 30px;
    min-height: 30px; max-height: 30px; font-size: 12px; font-weight: 700;
}
QPushButton#helpIconButton:hover { background: #18212d; color: #f8fafc; }
QProgressBar { background: #111827; border: 1px solid #334155; border-radius: 4px; height: 8px; }
QProgressBar::chunk { background: #f97316; border-radius: 3px; }
QRadioButton::indicator, QCheckBox::indicator { width: 14px; height: 14px; border: 1px solid #64748b; background: #0f1720; }
QRadioButton::indicator { border-radius: 7px; }
QRadioButton::indicator:checked, QCheckBox::indicator:checked { background: #f97316; border-color: #fb923c; }
QRadioButton, QCheckBox { spacing: 8px; color: #cbd5e1; }
QRadioButton:disabled, QCheckBox:disabled { color: #64748b; }
QWidget#cleanupActionPanel { background: #090f16; }
QLabel#cleanupActionLabel { color: #cbd5e1; font-weight: 600; }
QWidget#cleanupForcePanel { background: #111a24; border-top: 1px solid #27364a; }
QCheckBox#cleanupForceCheck { color: #fed7aa; font-weight: 600; }
QLabel#cleanupForceHint { color: #94a3b8; font-size: 11px; }
QDialog#libraryHealthDialog QLabel#healthSummary { color: #f8fafc; font-size: 12px; font-weight: 600; padding: 3px 0; }
QDialog#libraryHealthDialog QLabel#healthHint { color: #94a3b8; font-size: 11px; padding: 3px 0; }
QTreeWidget#healthTree { font-size: 11px; }
QTreeWidget#healthTree::item { min-height: 26px; }
QMainWindow, QDialog, QWidget, QLabel, QPushButton, QTreeWidget,
QPlainTextEdit, QComboBox, QLineEdit, QMenu, QHeaderView::section {
    font-size: 12px;
}
QPushButton, QPushButton#dangerAction, QPushButton#inputChip,
QPushButton#primaryAction, QPushButton#statFilter {
    font-size: 12px;
}
QFrame#workflowFiles QPushButton#inputChip,
QWidget#libraryActions QPushButton,
QWidget#libraryActions QPushButton#primaryAction,
QWidget#libraryActions QPushButton#statFilter {
    font-size: 12px;
    font-weight: 500;
    padding: 5px 8px;
}
QWidget#libraryActions[compact="true"] QPushButton,
QWidget#libraryActions[compact="true"] QPushButton#statFilter {
    padding: 5px 6px;
}
QLabel#pageTitle { font-size: 18px; }
QLabel#detailTitle, QDialog QLabel#dialogTitle { font-size: 15px; }
QGroupBox::title { font-size: 13px; }
QLabel#pageSubtitle, QLabel#projectPath, QLabel#statMeta,
QLabel#detailLabel, QLabel#detailMetricLabel, QLabel#detailBadge,
QDialog QLabel#dialogSubtitle, QDialog#libraryHealthDialog QLabel#healthHint {
    font-size: 11px;
}
QLabel#detailMeta, QLabel#detailValue, QLabel#detailMetricValue,
QDialog QLabel#dialogSummary, QDialog QLabel#previewHeading,
QDialog#libraryHealthDialog QLabel#healthSummary {
    font-size: 11px;
}
QHeaderView::section, QPlainTextEdit#operationLog { font-size: 11px; }
"""
    + SHARED_MAIN_QSS
    + SHARED_DIALOG_QSS
)


class OperationWorker(QObject):
    message = Signal(str)
    finished = Signal(object)
    failed = Signal(str)

    def __init__(
        self, operation: Callable[[Callable[[str], None], Callable[[], bool]], Any]
    ) -> None:
        super().__init__()
        self.operation = operation
        self.cancelled = threading.Event()
        self.thread_id: int | None = None

    @Slot()
    def run(self) -> None:
        self.thread_id = threading.get_ident()
        try:
            result = self.operation(self.message.emit, self.cancelled.is_set)
        except Exception as exc:
            LOGGER.exception("Media manager operation failed")
            self.failed.emit(str(exc) or type(exc).__name__)
        else:
            self.finished.emit(result)

    def cancel(self) -> None:
        self.cancelled.set()
