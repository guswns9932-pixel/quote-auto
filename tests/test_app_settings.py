"""설정 저장 — 통합양식 위치는 PC 마다 따로 기억한다 (같은 계정의 로밍 설정을 공유하더라도)."""
import pytest

pytest.importorskip("PySide6.QtCore")
from PySide6.QtCore import QSettings

import app_settings
from app_settings import Key


@pytest.fixture()
def settings(tmp_path):
    QSettings.setPath(QSettings.IniFormat, QSettings.UserScope, str(tmp_path / "settings"))
    return tmp_path


def _on(monkeypatch, pc):
    monkeypatch.setenv("COMPUTERNAME", pc)


def test_each_pc_remembers_its_own_template(settings, monkeypatch):
    a = settings / "A_양식.xlsx"; a.write_bytes(b"")
    b = settings / "B_양식.xlsx"; b.write_bytes(b"")

    _on(monkeypatch, "PC-A")
    app_settings.set_pc_str(Key.TEMPLATE_PATH, str(a))
    _on(monkeypatch, "PC-B")
    app_settings.set_pc_str(Key.TEMPLATE_PATH, str(b))

    _on(monkeypatch, "PC-A")
    assert app_settings.get_pc_path(Key.TEMPLATE_PATH) == str(a)
    _on(monkeypatch, "PC-B")
    assert app_settings.get_pc_path(Key.TEMPLATE_PATH) == str(b)


def test_first_run_after_update_uses_the_old_shared_value(settings, monkeypatch):
    old = settings / "예전.xlsx"; old.write_bytes(b"")
    app_settings.set_str(Key.TEMPLATE_PATH, str(old))          # 예전 버전이 저장한 공용 값
    _on(monkeypatch, "PC-A")
    assert app_settings.get_pc_path(Key.TEMPLATE_PATH) == str(old)


def test_missing_file_on_this_pc_does_not_fall_back_to_another_pcs(settings, monkeypatch):
    other = settings / "다른PC.xlsx"; other.write_bytes(b"")
    app_settings.set_str(Key.TEMPLATE_PATH, str(other))
    _on(monkeypatch, "PC-A")
    app_settings.set_pc_str(Key.TEMPLATE_PATH, str(settings / "Z드라이브" / "양식.xlsx"))  # 지금은 없는 경로
    assert app_settings.get_pc_path(Key.TEMPLATE_PATH) == ""
    assert app_settings.get_pc_str(Key.TEMPLATE_PATH).endswith("양식.xlsx")  # 지우지 않는다 — 드라이브가 붙으면 다시 쓴다


def test_pc_name_is_safe_as_a_settings_key(monkeypatch):
    monkeypatch.setenv("COMPUTERNAME", "lot-pc/01 \\x")
    assert "/" not in app_settings.pc_name() and "\\" not in app_settings.pc_name()
