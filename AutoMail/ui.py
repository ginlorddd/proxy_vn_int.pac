from __future__ import annotations

import csv
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
    from PySide6.QtGui import QAction, QColor, QTextCharFormat, QTextCursor, QTextListFormat
    from PySide6.QtWidgets import (
        QApplication,
        QCalendarWidget,
        QCheckBox,
        QColorDialog,
        QComboBox,
        QFileDialog,
        QFontComboBox,
        QFormLayout,
        QFrame,
        QGroupBox,
        QHBoxLayout,
        QHeaderView,
        QLabel,
        QLineEdit,
        QMainWindow,
        QMessageBox,
        QPushButton,
        QScrollArea,
        QSizePolicy,
        QSpinBox,
        QTableWidget,
        QTableWidgetItem,
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


def _extract_email(value: str) -> str:
    text = value.strip()
    if "<" in text and ">" in text:
        return text.split("<", 1)[1].split(">", 1)[0].strip()
    return text


def _account_smtp(account: Any) -> str:
    smtp = str(getattr(account, "SmtpAddress", "") or "").strip()
    if smtp:
        return smtp
    try:
        user = account.CurrentUser
        entry = user.AddressEntry
        if str(entry.Type).upper() == "EX":
            exchange_user = entry.GetExchangeUser()
            smtp = str(getattr(exchange_user, "PrimarySmtpAddress", "") or "").strip()
            if smtp:
                return smtp
        return str(getattr(entry, "Address", "") or "").strip()
    except Exception:
        return ""


def _add_account_value(accounts: list[str], display: str, smtp: str) -> None:
    display = display.strip()
    smtp = smtp.strip()
    value = f"{display} <{smtp}>" if smtp and display and display.lower() != smtp.lower() else smtp or display
    if value and value not in accounts:
        accounts.append(value)


def _iter_com_collection(collection: Any) -> list[Any]:
    try:
        return [collection.Item(index) for index in range(1, int(collection.Count) + 1)]
    except Exception:
        try:
            return list(collection)
        except Exception:
            return []


def get_outlook_accounts() -> list[str]:
    """Đọc đầy đủ account Outlook đã đăng nhập từ MAPI, Exchange và CurrentUser."""
    pythoncom = None
    try:
        import pythoncom as _pythoncom  # type: ignore
        import win32com.client  # type: ignore

        pythoncom = _pythoncom
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        except Exception:
            outlook = win32com.client.Dispatch("Outlook.Application")
        session = outlook.Session or outlook.GetNamespace("MAPI")
        accounts: list[str] = []

        for account in _iter_com_collection(session.Accounts):
            smtp = _account_smtp(account)
            display = str(getattr(account, "DisplayName", "") or smtp).strip()
            _add_account_value(accounts, display, smtp)

        current_user = getattr(session, "CurrentUser", None)
        if current_user is not None:
            try:
                entry = current_user.AddressEntry
                smtp = ""
                if str(getattr(entry, "Type", "")).upper() == "EX":
                    exchange_user = entry.GetExchangeUser()
                    smtp = str(getattr(exchange_user, "PrimarySmtpAddress", "") or "").strip()
                smtp = smtp or str(getattr(entry, "Address", "") or "").strip()
                _add_account_value(accounts, str(getattr(current_user, "Name", "") or smtp), smtp)
            except Exception:
                pass

        return accounts
    except Exception:
        return []
    finally:
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


MODERN_STYLE = """
QMainWindow { background: #f3f6fb; }
QWidget { color: #111827; font-family: Segoe UI, Arial; font-size: 10pt; }
QLabel { background: transparent; color: #111827; }
QGroupBox#card { background: #ffffff; border: 1px solid #dbe3ef; border-radius: 14px; margin-top: 14px; padding: 16px; font-weight: 700; }
QGroupBox#card::title { subcontrol-origin: margin; left: 16px; padding: 0 8px; color: #1d4ed8; }
QFrame#hero { background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #1d4ed8, stop:1 #06b6d4); border-radius: 18px; }
QLabel#heroTitle { color: white; font-size: 22pt; font-weight: 800; }
QLabel#heroSubtitle { color: #e0f2fe; font-size: 10.5pt; }
QLabel#pill { background: rgba(255,255,255,0.20); color: white; border-radius: 10px; padding: 6px 10px; font-weight: 700; }
QLineEdit, QTextEdit, QComboBox, QSpinBox, QTableWidget { background: white; border: 1px solid #d7deea; border-radius: 9px; padding: 5px 8px; selection-background-color: #bfdbfe; }
QLineEdit:focus, QTextEdit:focus, QComboBox:focus, QSpinBox:focus { border: 1px solid #2563eb; }
QPushButton { background: #2563eb; color: white; border: none; border-radius: 9px; padding: 9px 14px; font-weight: 700; }
QPushButton:hover { background: #1d4ed8; }
QPushButton#secondary { background: #e8eef8; color: #1f2937; }
QPushButton#success { background: #059669; }
QToolBar { background: #ffffff; border: 1px solid #d7deea; spacing: 6px; padding: 6px; }
QHeaderView::section { background: #e8eef8; padding: 8px; border: 0; font-weight: 700; }
"""


class AutoMailWindow(QMainWindow):
    """Giao diện cấu hình mail dùng PySide6 QTextEdit + toolbar định dạng giống Outlook."""

    def __init__(self, scheduler: MailScheduler) -> None:
        super().__init__()
        self.scheduler = scheduler
        self.config = load_config()
        self.setWindowTitle("AutoMail - Outlook style editor")
        self.resize(1180, 820)
        self.setStyleSheet(MODERN_STYLE)
        self._build_ui()
        self._load_to_form()

    def _build_ui(self) -> None:
        root = QWidget(self)
        page_layout = QVBoxLayout(root)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(0)
        viewport = QScrollArea(root)
        viewport.setWidgetResizable(True)
        viewport.setFrameShape(QFrame.Shape.NoFrame)
        content = QWidget(viewport)
        layout = QVBoxLayout(content)
        layout.setContentsMargins(18, 18, 18, 12)
        layout.setSpacing(12)
        viewport.setWidget(content)
        page_layout.addWidget(viewport, 1)
        self.setCentralWidget(root)

        hero = QFrame()
        hero.setObjectName("hero")
        hero_layout = QHBoxLayout(hero)
        hero_layout.setContentsMargins(22, 18, 22, 18)
        title_col = QVBoxLayout()
        hero_title = QLabel("AutoMail Outlook")
        hero_title.setObjectName("heroTitle")
        hero_subtitle = QLabel("Soạn mail rich-text, chọn tài khoản Outlook đã đăng nhập và lập lịch gửi tự động.")
        hero_subtitle.setObjectName("heroSubtitle")
        title_col.addWidget(hero_title)
        title_col.addWidget(hero_subtitle)
        hero_layout.addLayout(title_col, 1)
        self.account_count = QLabel("Outlook: đang tải")
        self.account_count.setObjectName("pill")
        hero_layout.addWidget(self.account_count)
        layout.addWidget(hero)

        mail_card = QGroupBox("Thông tin gửi mail")
        mail_card.setObjectName("card")
        mail_card.setMaximumHeight(310)
        form = QFormLayout(mail_card)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignLeft)
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(8)
        self.account = QComboBox()
        self.account.setEditable(False)
        self.account.setPlaceholderText("Chọn account Outlook đã đăng nhập")
        self.account.setMinimumWidth(360)
        self.account.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.refresh_accounts()
        self.to = QLineEdit()
        self.cc = QLineEdit()
        self.bcc = QLineEdit()
        self.subject = QLineEdit()
        for field in (self.account, self.to, self.cc, self.bcc, self.subject):
            field.setMinimumHeight(34)
        account_widget = QWidget()
        account_widget.setObjectName("inlineActions")
        account_row = QHBoxLayout(account_widget)
        account_row.setContentsMargins(0, 0, 0, 0)
        account_row.setSpacing(8)
        account_row.addWidget(self.account, 1)
        refresh_accounts = QPushButton("Tải account Outlook")
        refresh_accounts.setObjectName("secondary")
        refresh_accounts.setMinimumWidth(150)
        refresh_accounts.clicked.connect(self.refresh_accounts)
        account_row.addWidget(refresh_accounts)
        form.addRow("From/account", account_widget)
        form.addRow("Tới", self.to)
        form.addRow("Cc", self.cc)
        form.addRow("Bcc", self.bcc)
        import_recipients = QPushButton("Import To/Cc/Bcc")
        import_recipients.setObjectName("secondary")
        import_recipients.clicked.connect(self.import_recipients)
        form.addRow("Import người nhận", import_recipients)
        form.addRow("Tiêu đề", self.subject)
        layout.addWidget(mail_card)

        editor_card = QGroupBox("Nội dung email")
        editor_card.setObjectName("card")
        editor_card.setMinimumHeight(500)
        editor_card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        editor_layout = QVBoxLayout(editor_card)
        self.editor = QTextEdit()
        self.editor.setAcceptRichText(True)
        self.editor.setMinimumHeight(440)
        self.editor.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.editor.setPlaceholderText("Soạn nội dung mail tại đây hoặc chọn file .eml để nạp nội dung...")
        editor_layout.addWidget(self.editor, 1)
        layout.addWidget(editor_card, 5)
        self._build_toolbar()

        schedule_card = QGroupBox("Lịch gửi")
        schedule_card.setObjectName("card")
        schedule_card.setMaximumHeight(500)
        schedule_layout = QVBoxLayout(schedule_card)
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
        schedule_layout.addLayout(schedule_row)

        weekday_row = QHBoxLayout()
        weekday_row.addWidget(QLabel("Gửi vào thứ"))
        self.weekday_checks = []
        for label, value in [("T2", 0), ("T3", 1), ("T4", 2), ("T5", 3), ("T6", 4), ("T7", 5), ("CN", 6)]:
            check = QCheckBox(label)
            check.setProperty("weekday", value)
            weekday_row.addWidget(check)
            self.weekday_checks.append(check)
        weekday_row.addStretch()
        schedule_layout.addLayout(weekday_row)

        self.calendar = QCalendarWidget()
        self.calendar.setMaximumHeight(130)
        self.calendar.setGridVisible(True)
        self.calendar.selectionChanged.connect(self.add_selected_date_schedule)
        schedule_layout.addWidget(QLabel("Lịch gửi theo ngày cụ thể (chọn ngày trên calendar để thêm dòng gửi):"))
        schedule_layout.addWidget(self.calendar)
        self.date_schedule_table = QTableWidget(0, 3)
        self.date_schedule_table.setMinimumHeight(120)
        self.date_schedule_table.setMaximumHeight(150)
        self.date_schedule_table.setHorizontalHeaderLabels(["Ngày", "Giờ", "Template/Nội dung mail"])
        self.date_schedule_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        schedule_layout.addWidget(self.date_schedule_table)
        layout.addWidget(schedule_card)

        buttons = QHBoxLayout()
        buttons.setSpacing(10)
        pick_eml = QPushButton("Chọn .eml và nạp nội dung")
        pick_eml.clicked.connect(self.pick_eml)
        save = QPushButton("Lưu config")
        save.clicked.connect(self.save)
        send = QPushButton("Gửi thử")
        send.setObjectName("success")
        send.clicked.connect(self.send_test)
        state = QPushButton("Xem trạng thái")
        state.clicked.connect(self.show_state)
        import_recipients_quick = QPushButton("Import To/Cc/Bcc")
        import_recipients_quick.setObjectName("secondary")
        import_recipients_quick.clicked.connect(self.import_recipients)
        for btn in (pick_eml, import_recipients_quick, save, send, state):
            btn.setMinimumHeight(34)
            buttons.addWidget(btn)
        buttons.addStretch()
        buttons.setContentsMargins(18, 10, 18, 10)
        page_layout.addLayout(buttons)
        self.statusBar().showMessage(f"Config: {CONFIG_FILE}")

    def _build_toolbar(self) -> None:
        bar = QToolBar("Định dạng văn bản", self)
        self.addToolBar(Qt.ToolBarArea.TopToolBarArea, bar)
        self.font_family = QFontComboBox()
        self.font_family.currentFontChanged.connect(lambda font: self._set_editor_font_family(font.family()))
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
            ("Màu chữ", self.pick_text_color),
            ("Màu nền", self.pick_background_color),
            ("• List", lambda: self._insert_list(QTextListFormat.Style.ListDisc)),
            ("1. List", lambda: self._insert_list(QTextListFormat.Style.ListDecimal)),
            ("Left", lambda: self.editor.setAlignment(Qt.AlignmentFlag.AlignLeft)),
            ("Center", lambda: self.editor.setAlignment(Qt.AlignmentFlag.AlignCenter)),
            ("Right", lambda: self.editor.setAlignment(Qt.AlignmentFlag.AlignRight)),
        ]:
            action = QAction(title, self)
            action.triggered.connect(slot)
            bar.addAction(action)

    def _set_editor_font_family(self, font_family: str) -> None:
        if hasattr(self, "editor"):
            self.editor.setFontFamily(font_family)

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
        account = str(mail.get("account", "") or "")
        if account:
            match = self.account.findData(account)
            if match == -1:
                match = self.account.findText(account)
            if match == -1:
                self.account.addItem(account, _extract_email(account))
                match = self.account.count() - 1
            self.account.setCurrentIndex(match)
        else:
            self.account.setCurrentIndex(0)
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
        self._load_date_schedules(schedule.get("date_schedules", []))

    def _form_config(self) -> dict[str, Any]:
        cfg = load_config()
        cfg["mail"].update({
            "account": "" if self.account.currentIndex() <= 0 else str(self.account.currentData() or _extract_email(self.account.currentText())),
            "to": _split(self.to.text()),
            "cc": _split(self.cc.text()),
            "bcc": _split(self.bcc.text()),
            "subject": self.subject.text().strip(),
            "body": self.editor.toHtml(),
            "body_format": "html",
            "font_family": self.font_family.currentFont().family(),
            "font_size": self.font_size.value(),
        })
        selected_weekdays = [int(check.property("weekday")) for check in self.weekday_checks if check.isChecked()]
        cfg["schedule"] = {
            "enabled": self.schedule_enabled.isChecked(),
            "type": self.schedule_type.currentText(),
            "time": self.schedule_time.text().strip() or "08:00",
            "weekdays": selected_weekdays or [0, 1, 2, 3, 4],
            "interval_minutes": self.interval.value(),
            "date_schedules": self._collect_date_schedules(),
        }
        return cfg

    def save(self) -> None:
        self.config = self._form_config()
        save_config(self.config)
        self.statusBar().showMessage("Đã lưu config.json", 4000)

    def refresh_accounts(self) -> None:
        current = ""
        if hasattr(self, "account") and self.account.currentIndex() > 0:
            current = str(self.account.currentData() or self.account.currentText()).strip()
        accounts = get_outlook_accounts()
        self.account.clear()
        self.account.addItem("Chọn account Outlook đã đăng nhập")
        self.account.setItemData(0, "", Qt.ItemDataRole.UserRole)
        for account in accounts:
            self.account.addItem(account, _extract_email(account))
        if current:
            match = self.account.findData(current)
            if match == -1:
                match = self.account.findText(current)
            if match == -1:
                extracted = _extract_email(current)
                match = next((index for index in range(self.account.count()) if self.account.itemData(index) == extracted), -1)
            if match == -1:
                self.account.addItem(current, _extract_email(current))
                match = self.account.count() - 1
            self.account.setCurrentIndex(match)
        elif accounts:
            self.account.setCurrentIndex(1)
        if hasattr(self, "account_count"):
            self.account_count.setText(f"Outlook: {len(accounts)} account" if accounts else "Outlook: chưa tìm thấy")
        if hasattr(self, "statusBar"):
            self.statusBar().showMessage(f"Đã tải {len(accounts)} account Outlook", 4000)

    def pick_text_color(self) -> None:
        color = QColorDialog.getColor(QColor("#111827"), self, "Chọn màu chữ")
        if color.isValid():
            self.editor.setTextColor(color)

    def pick_background_color(self) -> None:
        color = QColorDialog.getColor(QColor("#fff3bf"), self, "Chọn màu nền")
        if color.isValid():
            self.editor.setTextBackgroundColor(color)

    def import_recipients(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Import người nhận", "", "CSV/Text (*.csv *.txt);;All files (*.*)")
        if not path:
            return
        to_values: list[str] = []
        cc_values: list[str] = []
        bcc_values: list[str] = []
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            sample = f.read(2048)
            f.seek(0)
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t") if sample.strip() else csv.excel
            reader = csv.DictReader(f, dialect=dialect)
            fields = {name.lower() for name in (reader.fieldnames or [])}
            if fields & {"to", "cc", "bcc"}:
                for row in reader:
                    to_values.extend(_split(row.get("to", "") or row.get("To", "")))
                    cc_values.extend(_split(row.get("cc", "") or row.get("Cc", "")))
                    bcc_values.extend(_split(row.get("bcc", "") or row.get("Bcc", "")))
            else:
                f.seek(0)
                for line in f:
                    to_values.extend(_split(line))
        self.to.setText(_join(_split(self.to.text()) + to_values))
        self.cc.setText(_join(_split(self.cc.text()) + cc_values))
        self.bcc.setText(_join(_split(self.bcc.text()) + bcc_values))

    def add_selected_date_schedule(self) -> None:
        date_text = self.calendar.selectedDate().toString("yyyy-MM-dd")
        row = self.date_schedule_table.rowCount()
        self.date_schedule_table.insertRow(row)
        for col, value in enumerate([date_text, self.schedule_time.text().strip() or "08:00", self.config.get("mail", {}).get("template", "")]):
            self.date_schedule_table.setItem(row, col, QTableWidgetItem(value))

    def _load_date_schedules(self, schedules: list[dict[str, Any]]) -> None:
        self.date_schedule_table.setRowCount(0)
        for item in schedules:
            row = self.date_schedule_table.rowCount()
            self.date_schedule_table.insertRow(row)
            self.date_schedule_table.setItem(row, 0, QTableWidgetItem(str(item.get("date", ""))))
            self.date_schedule_table.setItem(row, 1, QTableWidgetItem(str(item.get("time", "08:00"))))
            self.date_schedule_table.setItem(row, 2, QTableWidgetItem(str(item.get("template", ""))))

    def _collect_date_schedules(self) -> list[dict[str, str]]:
        schedules = []
        for row in range(self.date_schedule_table.rowCount()):
            date_item = self.date_schedule_table.item(row, 0)
            time_item = self.date_schedule_table.item(row, 1)
            template_item = self.date_schedule_table.item(row, 2)
            if date_item and date_item.text().strip():
                schedules.append({
                    "date": date_item.text().strip(),
                    "time": time_item.text().strip() if time_item else "08:00",
                    "template": template_item.text().strip() if template_item else "",
                })
        return schedules

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
