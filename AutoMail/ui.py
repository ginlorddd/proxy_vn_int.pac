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

PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

try:
    from PySide6.QtCore import Qt
    from PySide6.QtGui import QAction, QColor, QTextCharFormat, QTextCursor, QTextListFormat
    from PySide6.QtWidgets import (
        QAbstractItemView,
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


def _property_accessor_smtp(obj: Any) -> str:
    try:
        return str(obj.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) or "").strip()
    except Exception:
        return ""


def _address_entry_smtp(entry: Any) -> str:
    try:
        if str(getattr(entry, "Type", "")).upper() == "EX":
            exchange_user = entry.GetExchangeUser()
            smtp = str(getattr(exchange_user, "PrimarySmtpAddress", "") or "").strip()
            if smtp:
                return smtp
        return str(getattr(entry, "Address", "") or "").strip()
    except Exception:
        return ""


def _account_smtp(account: Any) -> str:
    for candidate in (
        str(getattr(account, "SmtpAddress", "") or "").strip(),
        str(getattr(account, "UserName", "") or "").strip(),
    ):
        if "@" in candidate:
            return candidate
    try:
        smtp = _address_entry_smtp(account.CurrentUser.AddressEntry)
        if smtp:
            return smtp
    except Exception:
        pass
    try:
        smtp = _property_accessor_smtp(account.DeliveryStore)
        if smtp:
            return smtp
    except Exception:
        pass
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


def _outlook_application(win32com: Any) -> Any:
    try:
        return win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        pass
    try:
        return win32com.client.Dispatch("Outlook.Application")
    except Exception:
        return win32com.client.gencache.EnsureDispatch("Outlook.Application")


def _session_current_user_smtp(session: Any) -> tuple[str, str]:
    current_user = getattr(session, "CurrentUser", None)
    if current_user is None:
        return "", ""
    try:
        smtp = _address_entry_smtp(current_user.AddressEntry)
        return str(getattr(current_user, "Name", "") or smtp), smtp
    except Exception:
        return "", ""


def get_outlook_accounts() -> list[str]:
    """Đọc đầy đủ account Outlook đã đăng nhập từ MAPI, Exchange, Stores và CurrentUser."""
    pythoncom = None
    try:
        import pythoncom as _pythoncom  # type: ignore
        import win32com.client  # type: ignore

        pythoncom = _pythoncom
        pythoncom.CoInitialize()
        outlook = _outlook_application(win32com)
        session = outlook.GetNamespace("MAPI")
        try:
            session.Logon("", "", False, False)
        except Exception:
            pass
        accounts: list[str] = []

        for account in _iter_com_collection(session.Accounts):
            smtp = _account_smtp(account)
            display = str(getattr(account, "DisplayName", "") or getattr(account, "UserName", "") or smtp).strip()
            _add_account_value(accounts, display, smtp)

        for store in _iter_com_collection(getattr(session, "Stores", [])):
            smtp = _property_accessor_smtp(store)
            display = str(getattr(store, "DisplayName", "") or smtp).strip()
            _add_account_value(accounts, display, smtp)

        display, smtp = _session_current_user_smtp(session)
        _add_account_value(accounts, display, smtp)
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
QLabel#statusRunning { background: #dcfce7; color: #166534; border-radius: 10px; padding: 6px 10px; font-weight: 800; }
QLabel#statusStopping { background: #fef3c7; color: #92400e; border-radius: 10px; padding: 6px 10px; font-weight: 800; }
QLabel#statusStopped { background: #fee2e2; color: #991b1b; border-radius: 10px; padding: 6px 10px; font-weight: 800; }
QLineEdit, QTextEdit, QComboBox, QSpinBox, QTableWidget { background: white; border: 1px solid #d7deea; border-radius: 9px; padding: 5px 8px; selection-background-color: #bfdbfe; }
QComboBox::drop-down { subcontrol-origin: padding; subcontrol-position: top right; width: 30px; border-left: 1px solid #d7deea; border-top-right-radius: 9px; border-bottom-right-radius: 9px; background: #eef4ff; }
QComboBox::down-arrow { width: 10px; height: 10px; }
QSpinBox::up-button { subcontrol-origin: border; subcontrol-position: top right; width: 26px; border-left: 1px solid #d7deea; border-top-right-radius: 9px; background: #eef4ff; }
QSpinBox::down-button { subcontrol-origin: border; subcontrol-position: bottom right; width: 26px; border-left: 1px solid #d7deea; border-bottom-right-radius: 9px; background: #eef4ff; }
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
        self._marked_schedule_dates: set[str] = set()
        self.setWindowTitle("AutoMail - Outlook style editor")
        self.resize(1180, 820)
        self.setStyleSheet(MODERN_STYLE)
        self._build_ui()
        self._load_to_form()
        self.update_scheduler_status()

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
        self.scheduler_status = QLabel("Scheduler: stopped")
        hero_layout.addWidget(self.scheduler_status)
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
        self.calendar.selectionChanged.connect(self.show_selected_date_info)
        schedule_layout.addWidget(QLabel("Lịch gửi master: chọn ngày để xem/thêm cấu hình gửi mail tự động:"))
        schedule_layout.addWidget(self.calendar)
        template_row = QHBoxLayout()
        template_row.addWidget(QLabel("Template có sẵn"))
        self.template_combo = QComboBox()
        self.template_combo.setMinimumWidth(360)
        template_row.addWidget(self.template_combo, 1)
        add_library_template = QPushButton("Thêm template có sẵn")
        add_library_template.setObjectName("secondary")
        add_library_template.clicked.connect(self.add_template_library_item)
        template_row.addWidget(add_library_template)
        schedule_layout.addLayout(template_row)
        self.date_schedule_table = QTableWidget(0, 8)
        self.date_schedule_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.date_schedule_table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.date_schedule_table.setMinimumHeight(150)
        self.date_schedule_table.setMaximumHeight(190)
        self.date_schedule_table.setHorizontalHeaderLabels(["Ngày", "Giờ", "Lặp", "From", "Template/Nội dung mail", "To", "Cc", "Bcc"])
        self.date_schedule_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        schedule_layout.addWidget(self.date_schedule_table)
        schedule_buttons = QHBoxLayout()
        new_config = QPushButton("Thêm mới cấu hình ngày")
        new_config.setObjectName("secondary")
        new_config.clicked.connect(self.new_day_mail_config)
        add_template = QPushButton("Thêm template vào lịch")
        add_template.setObjectName("secondary")
        add_template.clicked.connect(self.add_template_schedule)
        delete_rows = QPushButton("Xóa dòng đã chọn")
        delete_rows.setObjectName("secondary")
        delete_rows.clicked.connect(self.delete_selected_date_schedules)
        schedule_buttons.addWidget(new_config)
        schedule_buttons.addWidget(add_template)
        schedule_buttons.addWidget(delete_rows)
        schedule_buttons.addStretch()
        schedule_layout.addLayout(schedule_buttons)
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
        self.refresh_template_combo()
        self._load_date_schedules(schedule.get("date_schedules", []))
        self.refresh_calendar_markers()

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
        cfg["templates"] = [self.template_combo.itemData(index) or self.template_combo.itemText(index) for index in range(self.template_combo.count()) if index > 0]
        cfg["schedule"] = {
            "enabled": self.schedule_enabled.isChecked(),
            "type": self.schedule_type.currentText(),
            "time": self.schedule_time.text().strip() or "08:00",
            "weekdays": selected_weekdays or [0, 1, 2, 3, 4],
            "interval_minutes": self.interval.value(),
            "date_schedules": self._collect_date_schedules(),
        }
        return cfg

    def update_scheduler_status(self, status: str | None = None) -> None:
        status = status or (self.scheduler.status() if hasattr(self.scheduler, "status") else "stopped")
        labels = {"running": "Scheduler: running", "stopping": "Scheduler: stopping", "stopped": "Scheduler: stopped"}
        objects = {"running": "statusRunning", "stopping": "statusStopping", "stopped": "statusStopped"}
        self.scheduler_status.setText(labels.get(status, f"Scheduler: {status}"))
        self.scheduler_status.setObjectName(objects.get(status, "statusStopped"))
        self.scheduler_status.style().unpolish(self.scheduler_status)
        self.scheduler_status.style().polish(self.scheduler_status)

    def save(self) -> None:
        self.config = self._form_config()
        save_config(self.config)
        if self.config.get("schedule", {}).get("enabled"):
            self.scheduler.start()
        else:
            self.update_scheduler_status("stopping")
            QApplication.processEvents()
            self.scheduler.stop()
        self.update_scheduler_status()
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

    def refresh_template_combo(self) -> None:
        if not hasattr(self, "template_combo"):
            return
        current = self.template_combo.currentData() or self.template_combo.currentText()
        self.template_combo.clear()
        self.template_combo.addItem("Không dùng template có sẵn", "")
        templates = list(dict.fromkeys([*self.config.get("templates", []), self.config.get("mail", {}).get("template", "")]))
        for template in templates:
            if template:
                self.template_combo.addItem(Path(template).name, template)
        if current:
            index = self.template_combo.findData(current)
            if index >= 0:
                self.template_combo.setCurrentIndex(index)

    def add_template_library_item(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Thêm template có sẵn", str(Path(__file__).parent / "templates"), "Email (*.eml)")
        if not path:
            return
        cfg = load_config()
        templates = list(dict.fromkeys([*cfg.get("templates", []), path]))
        cfg["templates"] = templates
        save_config(cfg)
        self.config = cfg
        self.refresh_template_combo()
        self.template_combo.setCurrentIndex(self.template_combo.findData(path))
        self.statusBar().showMessage(f"Đã thêm template {Path(path).name}", 4000)

    def show_selected_date_info(self) -> None:
        self.refresh_calendar_markers()
        date_text = self.calendar.selectedDate().toString("yyyy-MM-dd")
        events = [item for item in self._collect_date_schedules() if item.get("date") == date_text]
        if events:
            summary = " | ".join(f"{item.get('time')} {Path(item.get('template', '')).name or item.get('subject', 'Mail')}" for item in events)
            self.calendar.setToolTip(summary)
            self.statusBar().showMessage(f"{date_text}: {summary}", 6000)
        else:
            self.calendar.setToolTip("Không có lịch gửi cho ngày này")

    def refresh_calendar_markers(self) -> None:
        default_format = QTextCharFormat()
        for date_text in self._marked_schedule_dates:
            self.calendar.setDateTextFormat(self.calendar.selectedDate().fromString(date_text, "yyyy-MM-dd"), default_format)
        self._marked_schedule_dates.clear()
        marker = QTextCharFormat()
        marker.setBackground(QColor("#dbeafe"))
        marker.setForeground(QColor("#1d4ed8"))
        marker.setFontWeight(700)
        for item in self._collect_date_schedules():
            date_text = item.get("date", "")
            if date_text:
                self.calendar.setDateTextFormat(self.calendar.selectedDate().fromString(date_text, "yyyy-MM-dd"), marker)
                self._marked_schedule_dates.add(date_text)

    def _current_account_value(self) -> str:
        return "" if self.account.currentIndex() <= 0 else str(self.account.currentData() or _extract_email(self.account.currentText()))

    def new_day_mail_config(self) -> None:
        self.to.clear()
        self.cc.clear()
        self.bcc.clear()
        self.subject.clear()
        self.editor.clear()
        self.template_combo.setCurrentIndex(0)
        self._insert_schedule_row(
            date=self.calendar.selectedDate().toString("yyyy-MM-dd"),
            time=self.schedule_time.text().strip() or "08:00",
            repeat="once",
            account=self._current_account_value(),
        )
        self.refresh_calendar_markers()

    def delete_selected_date_schedules(self) -> None:
        selected_rows = sorted({index.row() for index in self.date_schedule_table.selectedIndexes()}, reverse=True)
        if not selected_rows and self.date_schedule_table.currentRow() >= 0:
            selected_rows = [self.date_schedule_table.currentRow()]
        for row in selected_rows:
            self.date_schedule_table.removeRow(row)
        if selected_rows:
            self.refresh_calendar_markers()
            self.statusBar().showMessage(f"Đã xóa {len(selected_rows)} dòng lịch", 4000)

    def _insert_schedule_row(
        self,
        *,
        date: str,
        time: str,
        repeat: str = "once",
        account: str = "",
        template: str = "",
        to: str = "",
        cc: str = "",
        bcc: str = "",
    ) -> None:
        row = self.date_schedule_table.rowCount()
        self.date_schedule_table.insertRow(row)
        for col, value in enumerate([date, time, repeat, account, template, to, cc, bcc]):
            self.date_schedule_table.setItem(row, col, QTableWidgetItem(value))

    def add_selected_date_schedule(self) -> None:
        self._insert_schedule_row(
            date=self.calendar.selectedDate().toString("yyyy-MM-dd"),
            time=self.schedule_time.text().strip() or "08:00",
            repeat="once",
            account=self._current_account_value(),
            template=self.config.get("mail", {}).get("template", ""),
            to=self.to.text(),
            cc=self.cc.text(),
            bcc=self.bcc.text(),
        )
        self.refresh_calendar_markers()

    def add_template_schedule(self) -> None:
        path = str(self.template_combo.currentData() or "") if hasattr(self, "template_combo") else ""
        if not path:
            path, _ = QFileDialog.getOpenFileName(self, "Thêm template vào lịch", str(Path(__file__).parent / "templates"), "Email (*.eml)")
        if not path:
            return
        self._insert_schedule_row(
            date=self.calendar.selectedDate().toString("yyyy-MM-dd"),
            time=self.schedule_time.text().strip() or "08:00",
            repeat=self.schedule_type.currentText() if self.schedule_type.currentText() in {"daily", "weekly"} else "once",
            account=self._current_account_value(),
            template=path,
            to=self.to.text(),
            cc=self.cc.text(),
            bcc=self.bcc.text(),
        )
        self.refresh_calendar_markers()

    def _load_date_schedules(self, schedules: list[dict[str, Any]]) -> None:
        self.date_schedule_table.setRowCount(0)
        for item in schedules:
            self._insert_schedule_row(
                date=str(item.get("date", "")),
                time=str(item.get("time", "08:00")),
                repeat=str(item.get("repeat", "once")),
                account=str(item.get("account", "")),
                template=str(item.get("template", "")),
                to=_join(item.get("to", [])),
                cc=_join(item.get("cc", [])),
                bcc=_join(item.get("bcc", [])),
            )

    def _collect_date_schedules(self) -> list[dict[str, Any]]:
        schedules: list[dict[str, Any]] = []
        for row in range(self.date_schedule_table.rowCount()):
            values = [self.date_schedule_table.item(row, col).text().strip() if self.date_schedule_table.item(row, col) else "" for col in range(8)]
            if values[0]:
                schedules.append({
                    "date": values[0],
                    "time": values[1] or "08:00",
                    "repeat": values[2] or "once",
                    "account": values[3],
                    "template": values[4],
                    "to": _split(values[5]),
                    "cc": _split(values[6]),
                    "bcc": _split(values[7]),
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
