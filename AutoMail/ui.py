from __future__ import annotations

import json
import threading
from pathlib import Path
from typing import Any

from config import CONFIG_FILE, load_config, save_config
from eml_reader import read_eml
from scheduler import MailScheduler
from sender import send_mail
from state import load_state

try:
    from PySide6.QtCore import Qt
    from PySide6.QtGui import QAction, QTextCharFormat, QTextCursor, QTextListFormat
    from PySide6.QtWidgets import (
        QApplication,
        QCheckBox,
        QComboBox,
        QFileDialog,
        QFormLayout,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QMainWindow,
        QMessageBox,
        QPushButton,
        QSpinBox,
        QTextEdit,
        QToolBar,
        QVBoxLayout,
        QWidget,
    )
except ImportError as exc:  # pragma: no cover - chỉ xảy ra khi thiếu dependency UI
    raise RuntimeError("Cần cài PySide6 để mở giao diện AutoMail: pip install -r requirements.txt") from exc


def _join(value: Any) -> str:
    return "; ".join(value) if isinstance(value, list) else str(value or "")


def _split(value: str) -> list[str]:
    return [x.strip() for x in value.replace(",", ";").split(";") if x.strip()]


class AutoMailWindow(QMainWindow):
    """Giao diện cấu hình mail dùng PySide6 QTextEdit + toolbar định dạng giống Outlook."""

    def __init__(self, scheduler: MailScheduler) -> None:
        super().__init__()
        self.scheduler = scheduler
        self.config = load_config()
        self.setWindowTitle("AutoMail - Outlook style editor")
        self.resize(1100, 780)
        self._build_ui()
        self._load_to_form()

    def _build_ui(self) -> None:
        root = QWidget(self)
        layout = QVBoxLayout(root)
        self.setCentralWidget(root)

        form = QFormLayout()
        self.account = QLineEdit()
        self.to = QLineEdit()
        self.cc = QLineEdit()
        self.bcc = QLineEdit()
        self.subject = QLineEdit()
        form.addRow("From/account", self.account)
        form.addRow("Tới", self.to)
        form.addRow("Cc", self.cc)
        form.addRow("Bcc", self.bcc)
        form.addRow("Tiêu đề", self.subject)
        layout.addLayout(form)

        self.editor = QTextEdit()
        self.editor.setAcceptRichText(True)
        self.editor.setPlaceholderText("Soạn nội dung mail tại đây hoặc chọn file .eml để nạp nội dung...")
        layout.addWidget(self.editor, 1)
        self._build_toolbar()

        schedule_row = QHBoxLayout()
        self.schedule_enabled = QCheckBox("Bật schedule")
        self.schedule_type = QComboBox()
        self.schedule_type.addItems(["daily", "weekly", "interval"])
        self.schedule_time = QLineEdit("08:00")
        self.interval = QSpinBox()
        self.interval.setRange(1, 1440)
        self.interval.setValue(60)
        schedule_row.addWidget(self.schedule_enabled)
        schedule_row.addWidget(QLabel("Kiểu"))
        schedule_row.addWidget(self.schedule_type)
        schedule_row.addWidget(QLabel("Giờ HH:MM"))
        schedule_row.addWidget(self.schedule_time)
        schedule_row.addWidget(QLabel("Interval phút"))
        schedule_row.addWidget(self.interval)
        schedule_row.addStretch()
        layout.addLayout(schedule_row)

        weekday_row = QHBoxLayout()
        weekday_row.addWidget(QLabel("Gửi vào thứ"))
        self.weekday_checks = []
        for label, value in [("T2", 0), ("T3", 1), ("T4", 2), ("T5", 3), ("T6", 4), ("T7", 5), ("CN", 6)]:
            check = QCheckBox(label)
            check.setProperty("weekday", value)
            weekday_row.addWidget(check)
            self.weekday_checks.append(check)
        weekday_row.addStretch()
        layout.addLayout(weekday_row)

        buttons = QHBoxLayout()
        pick_eml = QPushButton("Chọn .eml và nạp nội dung")
        pick_eml.clicked.connect(self.pick_eml)
        save = QPushButton("Lưu config")
        save.clicked.connect(self.save)
        send = QPushButton("Gửi thử")
        send.clicked.connect(self.send_test)
        state = QPushButton("Xem trạng thái")
        state.clicked.connect(self.show_state)
        for btn in (pick_eml, save, send, state):
            buttons.addWidget(btn)
        buttons.addStretch()
        layout.addLayout(buttons)
        self.statusBar().showMessage(f"Config: {CONFIG_FILE}")

    def _build_toolbar(self) -> None:
        if not hasattr(self, "editor"):
            raise RuntimeError("AutoMail editor must be created before building the formatting toolbar.")
        bar = QToolBar("Định dạng văn bản", self)
        self.addToolBar(Qt.ToolBarArea.TopToolBarArea, bar)
        self.font_family = QComboBox()
        self.font_family.addItems(["Calibri", "Arial", "Times New Roman", "Tahoma", "Verdana"])
        self.font_family.currentTextChanged.connect(self.editor.setFontFamily)
        self.font_size = QSpinBox()
        self.font_size.setRange(8, 48)
        self.font_size.setValue(11)
        self.font_size.valueChanged.connect(lambda s: self.editor.setFontPointSize(float(s)))
        bar.addWidget(QLabel("Font "))
        bar.addWidget(self.font_family)
        bar.addWidget(QLabel(" Size "))
        bar.addWidget(self.font_size)
        for title, slot in [
            ("B", lambda: self._merge_format(weight=True)),
            ("I", lambda: self.editor.setFontItalic(not self.editor.fontItalic())),
            ("U", lambda: self.editor.setFontUnderline(not self.editor.fontUnderline())),
            ("• List", lambda: self._insert_list(QTextListFormat.Style.ListDisc)),
            ("1. List", lambda: self._insert_list(QTextListFormat.Style.ListDecimal)),
            ("Left", lambda: self.editor.setAlignment(Qt.AlignmentFlag.AlignLeft)),
            ("Center", lambda: self.editor.setAlignment(Qt.AlignmentFlag.AlignCenter)),
            ("Right", lambda: self.editor.setAlignment(Qt.AlignmentFlag.AlignRight)),
        ]:
            action = QAction(title, self)
            action.triggered.connect(slot)
            bar.addAction(action)

    def _insert_list(self, style: QTextListFormat.Style) -> None:
        cursor = self.editor.textCursor()
        fmt = QTextListFormat()
        fmt.setStyle(style)
        cursor.createList(fmt)

    def _merge_format(self, *, weight: bool = False) -> None:
        fmt = QTextCharFormat()
        if weight:
            current = self.editor.fontWeight()
            fmt.setFontWeight(400 if current > 400 else 700)
        cursor = self.editor.textCursor()
        if not cursor.hasSelection():
            cursor.select(QTextCursor.SelectionType.WordUnderCursor)
        cursor.mergeCharFormat(fmt)
        self.editor.mergeCurrentCharFormat(fmt)

    def _load_to_form(self) -> None:
        mail = self.config["mail"]
        self.account.setText(mail.get("account", ""))
        self.to.setText(_join(mail.get("to")))
        self.cc.setText(_join(mail.get("cc")))
        self.bcc.setText(_join(mail.get("bcc")))
        self.subject.setText(mail.get("subject", ""))
        self.font_family.setCurrentText(mail.get("font_family", "Calibri"))
        self.font_size.setValue(int(mail.get("font_size", 11)))
        if mail.get("body_format", "html") == "html":
            self.editor.setHtml(mail.get("body", ""))
        else:
            self.editor.setPlainText(mail.get("body", ""))
        schedule = self.config["schedule"]
        self.schedule_enabled.setChecked(bool(schedule.get("enabled")))
        self.schedule_type.setCurrentText(schedule.get("type", "daily"))
        self.schedule_time.setText(schedule.get("time", "08:00"))
        self.interval.setValue(int(schedule.get("interval_minutes", 60)))
        weekdays = set(schedule.get("weekdays", [0, 1, 2, 3, 4]))
        for check in self.weekday_checks:
            check.setChecked(int(check.property("weekday")) in weekdays)

    def _form_config(self) -> dict[str, Any]:
        cfg = load_config()
        cfg["mail"].update({
            "account": self.account.text().strip(),
            "to": _split(self.to.text()),
            "cc": _split(self.cc.text()),
            "bcc": _split(self.bcc.text()),
            "subject": self.subject.text().strip(),
            "body": self.editor.toHtml(),
            "body_format": "html",
            "font_family": self.font_family.currentText(),
            "font_size": self.font_size.value(),
        })
        selected_weekdays = [int(check.property("weekday")) for check in self.weekday_checks if check.isChecked()]
        cfg["schedule"] = {
            "enabled": self.schedule_enabled.isChecked(),
            "type": self.schedule_type.currentText(),
            "time": self.schedule_time.text().strip() or "08:00",
            "weekdays": selected_weekdays or [0, 1, 2, 3, 4],
            "interval_minutes": self.interval.value(),
        }
        return cfg

    def save(self) -> None:
        self.config = self._form_config()
        save_config(self.config)
        self.statusBar().showMessage("Đã lưu config.json", 4000)

    def pick_eml(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Chọn template .eml", str(Path(__file__).parent / "templates"), "Email (*.eml)")
        if not path:
            return
        data = read_eml(path)
        if data.get("subject"):
            self.subject.setText(data["subject"])
        if data.get("body_format") == "html":
            self.editor.setHtml(data.get("body", ""))
        else:
            self.editor.setPlainText(data.get("body", ""))
        self.config["mail"]["template"] = path
        self.config["mail"]["subject"] = self.subject.text()
        self.config["mail"]["body"] = self.editor.toHtml()
        self.config["mail"]["body_format"] = "html"
        save_config(self.config)
        self.statusBar().showMessage(f"Đã nạp nội dung từ {path}", 6000)

    def send_test(self) -> None:
        self.save()

        def worker() -> None:
            result = send_mail(self.config["mail"])
            QMessageBox.information(self, "AutoMail", result["message"])

        threading.Thread(target=worker, daemon=True).start()

    def show_state(self) -> None:
        QMessageBox.information(self, "State", json.dumps(load_state(), ensure_ascii=False, indent=2))


def run_ui(scheduler: MailScheduler) -> None:
    app = QApplication.instance() or QApplication([])
    window = AutoMailWindow(scheduler)
    window.show()
    app.exec()


# Giữ alias tên cũ để main/đoạn tích hợp cũ không bị vỡ nếu import AutoMailUI.
AutoMailUI = AutoMailWindow
