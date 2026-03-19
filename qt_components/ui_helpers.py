from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget


def set_button_variant(button: QPushButton, variant: str) -> QPushButton:
    button.setProperty("variant", variant)
    button.style().unpolish(button)
    button.style().polish(button)
    button.update()
    return button


def create_info_card(parent: QWidget, title: str, hint: str, object_name: str = "PanelCard") -> QFrame:
    card = QFrame(parent)
    card.setObjectName(object_name)
    layout = QVBoxLayout(card)
    layout.setContentsMargins(18, 16, 18, 16)
    layout.setSpacing(6)

    title_label = QLabel(title, card)
    title_label.setObjectName("SectionTitle")
    hint_label = QLabel(hint, card)
    hint_label.setObjectName("SectionHint")
    hint_label.setWordWrap(True)

    layout.addWidget(title_label)
    layout.addWidget(hint_label)
    return card


def create_toolbar_card(parent: QWidget, title: str, hint: str | None = None) -> tuple[QFrame, QHBoxLayout]:
    card = QFrame(parent)
    card.setObjectName("SectionCard")
    layout = QHBoxLayout(card)
    layout.setContentsMargins(20, 18, 20, 18)
    layout.setSpacing(16)

    left = QVBoxLayout()
    left.setSpacing(5)
    title_label = QLabel(title, card)
    title_label.setObjectName("SectionTitle")
    left.addWidget(title_label)
    if hint:
        hint_label = QLabel(hint, card)
        hint_label.setObjectName("SectionHint")
        hint_label.setWordWrap(True)
        left.addWidget(hint_label)
    layout.addLayout(left, 1)
    return card, layout


def create_status_pill(parent: QWidget, text: str, tone: str = "neutral") -> QLabel:
    label = QLabel(text, parent)
    label.setAlignment(Qt.AlignCenter)
    label.setProperty("pillTone", tone)
    label.setObjectName("StatusPill")
    return label


def create_field_label(text: str, parent: QWidget) -> QLabel:
    label = QLabel(text, parent)
    label.setObjectName("FieldLabel")
    label.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
    label.setMinimumWidth(92)
    return label


def create_metric_card(parent: QWidget, title: str, value: str, tone: str = "neutral") -> tuple[QFrame, QLabel]:
    card = QFrame(parent)
    card.setObjectName("MetricCard")
    card.setProperty("metricTone", tone)
    layout = QVBoxLayout(card)
    layout.setContentsMargins(16, 14, 16, 14)
    layout.setSpacing(4)

    title_label = QLabel(title, card)
    title_label.setObjectName("MetricTitle")
    value_label = QLabel(value, card)
    value_label.setObjectName("MetricValue")

    layout.addWidget(title_label)
    layout.addWidget(value_label)
    return card, value_label
