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


def test_weekday_controls_are_persisted() -> None:
    source = UI_PATH.read_text(encoding="utf-8")
    assert "self.weekday_checks" in source
    assert "selected_weekdays" in source
    assert '"weekdays": selected_weekdays or [0, 1, 2, 3, 4]' in source
