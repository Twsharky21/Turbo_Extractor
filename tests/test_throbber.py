"""
test_throbber.py — Unit tests for Throbber widget and ThrobberMixin.

Tests:
  - test_throbber_starts_with_idle_char
  - test_throbber_start_sets_running
  - test_throbber_stop_resets_to_idle
  - test_throbber_start_is_idempotent
  - test_throbber_mixin_start_stop_no_widget (graceful no-op)
"""
from __future__ import annotations

import pytest

try:
    import tkinter as tk
    _root = tk.Tk()
    tk.Frame(_root)
    _root.destroy()
    _TCL_OK = True
except Exception:
    _TCL_OK = False

pytestmark = pytest.mark.skipif(not _TCL_OK, reason="Tcl/Tk not available")

from gui.mixins.throbber_mixin import Throbber, ThrobberMixin


def test_throbber_starts_with_idle_char():
    root = tk.Tk()
    t = Throbber(root)
    assert t.running is False
    # Idle state draws a small oval — canvas should have items
    root.update_idletasks()
    assert len(t.find_all()) > 0
    t.destroy()
    root.destroy()


def test_throbber_start_sets_running():
    root = tk.Tk()
    t = Throbber(root)
    t.start()
    assert t.running is True
    root.update_idletasks()
    # Spinning state draws background ring + arc = at least 2 items
    assert len(t.find_all()) >= 2
    t.stop()
    t.destroy()
    root.destroy()


def test_throbber_stop_resets_to_idle():
    root = tk.Tk()
    t = Throbber(root)
    t.start()
    root.update_idletasks()
    t.stop()
    assert t.running is False
    root.update_idletasks()
    # Back to idle — single oval
    assert len(t.find_all()) == 1
    t.destroy()
    root.destroy()


def test_throbber_start_is_idempotent():
    root = tk.Tk()
    t = Throbber(root)
    t.start()
    first_after_id = t._after_id
    t.start()  # second call should be a no-op
    assert t._after_id == first_after_id
    assert t.running is True
    t.stop()
    t.destroy()
    root.destroy()


def test_throbber_mixin_start_stop_no_widget():
    """ThrobberMixin must not crash when self.throbber doesn't exist."""

    class FakeApp(ThrobberMixin):
        pass

    app = FakeApp()
    # No self.throbber attribute at all — should be graceful no-ops
    app.throbber_start()
    app.throbber_stop()

    # Also test with throbber = None
    app.throbber = None
    app.throbber_start()
    app.throbber_stop()
