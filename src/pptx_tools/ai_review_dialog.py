from __future__ import annotations

from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from pptx_tools.manager_i18n import tr
from pptx_tools.ui_theme import SHARED_DIALOG_QSS


class AISuggestionDialog(QDialog):
    def __init__(
        self,
        parent: QWidget,
        result: dict[str, Any],
        current: dict[str, Any],
        field_labels: dict[str, str],
        item_names: dict[str, str],
    ) -> None:
        super().__init__(parent)
        self.result = result
        self.current = current
        self.checks: dict[str, QCheckBox] = {}
        self.setWindowTitle(tr("AI 整理建议"))
        self.setMinimumWidth(620)
        self.setMaximumHeight(720)
        self.setStyleSheet(SHARED_DIALOG_QSS)
        root = QVBoxLayout(self)
        root.addWidget(QLabel(tr("逐项核对后再应用；未勾选的内容不会修改。")))

        scroll = QScrollArea()
        scroll.setObjectName("aiReviewScroll")
        scroll.setWidgetResizable(True)
        body = QWidget()
        body.setObjectName("aiReviewBody")
        body_layout = QVBoxLayout(body)
        for key, label in field_labels.items():
            value = result.get(key)
            if not value:
                continue
            display = tr("、").join(value) if isinstance(value, list) else str(value)
            old = current.get(key)
            old_display = (
                tr("、").join(old) if isinstance(old, list) else str(old or tr("未填写"))
            )
            checkbox = QCheckBox(f"{label}{tr('：')}{old_display}  →  {display}")
            checkbox.setChecked(True)
            self.checks[key] = checkbox
            body_layout.addWidget(checkbox)

        groups = result.get("merge_groups") or []
        if groups:
            body_layout.addWidget(
                QLabel(tr("疑似同源与主资源建议（仅进入核对，不自动合并）："))
            )
        for group in groups:
            names = [item_names.get(item_id, item_id) for item_id in group["item_ids"]]
            primary = item_names.get(group["primary_id"], group["primary_id"])
            detail = QLabel(
                f"{' + '.join(names)}\n" +
                f"{tr('建议主资源：')}{primary}{tr('（')}{group['confidence']:.0%}{tr('）')}\n" +
                f"{tr('理由：')}{group['reason'] or tr('未说明')}"
            )
            detail.setWordWrap(True)
            detail.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            body_layout.addWidget(detail)
        body_layout.addStretch(1)
        scroll.setWidget(body)
        root.addWidget(scroll, 1)

        actions = QHBoxLayout()
        actions.addStretch(1)
        close = QPushButton(tr("关闭"))
        close.clicked.connect(self.reject)
        apply_button = QPushButton(tr("应用勾选字段"))
        apply_button.setObjectName("primaryAction")
        apply_button.setEnabled(bool(self.checks))
        apply_button.clicked.connect(self.accept)
        actions.addWidget(close)
        actions.addWidget(apply_button)
        root.addLayout(actions)
        self.setStyleSheet(
            SHARED_DIALOG_QSS
            + """
QScrollArea#aiReviewScroll {
    background: #0d1926;
    border: 1px solid #294057;
    border-radius: 8px;
}
QScrollArea#aiReviewScroll QWidget#qt_scrollarea_viewport,
QWidget#aiReviewBody {
    background: #0d1926;
    color: #cbd5e1;
}
QWidget#aiReviewBody QLabel,
QWidget#aiReviewBody QCheckBox {
    color: #cbd5e1;
}
"""
        )

    def selected_values(self) -> dict[str, Any]:
        return {
            key: self.result[key]
            for key, checkbox in self.checks.items()
            if checkbox.isChecked()
        }
