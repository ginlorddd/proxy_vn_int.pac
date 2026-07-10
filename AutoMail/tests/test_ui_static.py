from __future__ import annotations

import ast
from pathlib import Path

UI_PATH = Path(__file__).resolve().parents[1] / "ui.py"


def _method_body(class_node: ast.ClassDef, method_name: str) -> list[ast.stmt]:
    for node in class_node.body:
        if isinstance(node, ast.FunctionDef) and node.name == method_name:
            return node.body
    raise AssertionError(f"Missing method {method_name}")


def _is_self_call(node: ast.AST, method_name: str) -> bool:
    return (
        isinstance(node, ast.Expr)
        and isinstance(node.value, ast.Call)
        and isinstance(node.value.func, ast.Attribute)
        and node.value.func.attr == method_name
        and isinstance(node.value.func.value, ast.Name)
        and node.value.func.value.id == "self"
    )


def _assigns_self_editor(node: ast.AST) -> bool:
    return (
        isinstance(node, ast.Assign)
        and any(
            isinstance(target, ast.Attribute)
            and target.attr == "editor"
            and isinstance(target.value, ast.Name)
            and target.value.id == "self"
            for target in node.targets
        )
    )


def test_editor_is_created_before_toolbar_is_built() -> None:
    tree = ast.parse(UI_PATH.read_text(encoding="utf-8"))
    window = next(node for node in tree.body if isinstance(node, ast.ClassDef) and node.name == "AutoMailWindow")
    build_ui_body = _method_body(window, "_build_ui")

    editor_index = next(i for i, node in enumerate(build_ui_body) if _assigns_self_editor(node))
    toolbar_index = next(i for i, node in enumerate(build_ui_body) if _is_self_call(node, "_build_toolbar"))

    assert editor_index < toolbar_index


def test_toolbar_does_not_bind_to_editor_before_runtime() -> None:
    source = UI_PATH.read_text(encoding="utf-8")
    assert "currentTextChanged.connect(self.editor.setFontFamily)" not in source
    assert "currentFontChanged.connect" in source and "_set_editor_font_family" in source


def test_weekday_controls_are_persisted() -> None:
    source = UI_PATH.read_text(encoding="utf-8")
    assert "self.weekday_checks" in source
    assert "selected_weekdays" in source
    assert '"weekdays": selected_weekdays or [0, 1, 2, 3, 4]' in source


def test_modern_editor_features_are_present() -> None:
    source = UI_PATH.read_text(encoding="utf-8")
    assert "QFontComboBox" in source
    assert "QColorDialog" in source
    assert "get_outlook_accounts" in source
    assert "PrimarySmtpAddress" in source
    assert "Tải account Outlook" in source
    assert "import_recipients" in source
    assert "QCalendarWidget" in source
    assert "date_schedules" in source


def test_from_account_is_non_editable_outlook_dropdown() -> None:
    source = UI_PATH.read_text(encoding="utf-8")
    assert "self.account = QComboBox()" in source
    assert "self.account.setEditable(False)" in source
    assert "self.account.addItem(\"Chọn account Outlook đã đăng nhập\")" in source
    assert "self.account.addItem(account, _extract_email(account))" in source
    assert "currentData()" in source


def test_outlook_accounts_use_count_item_iteration() -> None:
    source = UI_PATH.read_text(encoding="utf-8")
    assert "def _iter_com_collection" in source
    assert "collection.Item(index)" in source
    assert "for account in _iter_com_collection(session.Accounts)" in source


def test_ui_keeps_action_buttons_visible_in_scrollable_layout() -> None:
    source = UI_PATH.read_text(encoding="utf-8")
    assert "QScrollArea" in source
    assert "viewport.setWidgetResizable(True)" in source
    assert "account_widget = QWidget()" in source
    assert "refresh_accounts.setMinimumWidth(150)" in source
    assert "import_recipients_quick = QPushButton(\"Import To/Cc/Bcc\")" in source


def test_editor_gets_priority_over_compact_sections() -> None:
    source = UI_PATH.read_text(encoding="utf-8")
    assert "editor_card.setMinimumHeight(500)" in source
    assert "self.editor.setMinimumHeight(440)" in source
    assert "layout.addWidget(editor_card, 5)" in source
    assert "mail_card.setMaximumHeight(310)" in source
    assert "schedule_card.setMaximumHeight(500)" in source
    assert "self.calendar.setMaximumHeight(130)" in source
    assert "self.date_schedule_table.setMinimumHeight(150)" in source
    assert "self.date_schedule_table.setMaximumHeight(190)" in source


def test_schedule_supports_multiple_templates_repeats_and_recipients() -> None:
    source = UI_PATH.read_text(encoding="utf-8")
    assert 'QComboBox::drop-down' in source
    assert 'QSpinBox::up-button' in source and 'QSpinBox::down-button' in source
    assert 'self.date_schedule_table = QTableWidget(0, 7)' in source
    assert '["Ngày", "Giờ", "Lặp", "Template/Nội dung mail", "To", "Cc", "Bcc"]' in source
    assert 'def add_template_schedule' in source
    assert '"repeat": values[2] or "once"' in source
    assert '"to": _split(values[4])' in source
