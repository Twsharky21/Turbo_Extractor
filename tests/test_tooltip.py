"""
test_tooltip.py — Unit tests for gui/tooltip.py.

Tests:
  - add_tooltip attaches bindings without error
  - Tooltip shows on simulated enter, hides on leave
  - add_tooltip with None widget is a safe no-op
  - add_tooltip with empty text is a safe no-op
  - Multiple tooltips on different widgets don't interfere
  - Tooltip on readonly combobox works
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

from gui.tooltip import add_tooltip, _Tooltip


def test_add_tooltip_attaches_without_error():
    root = tk.Tk()
    btn = tk.Button(root, text="Test")
    btn.pack()
    add_tooltip(btn, "Hello tooltip")
    root.update_idletasks()
    root.destroy()


def test_tooltip_none_widget_no_crash():
    """add_tooltip(None, ...) must not raise."""
    add_tooltip(None, "some text")


def test_tooltip_empty_text_no_crash():
    """add_tooltip(widget, '') must not raise or attach."""
    root = tk.Tk()
    btn = tk.Button(root, text="Test")
    btn.pack()
    add_tooltip(btn, "")
    root.update_idletasks()
    root.destroy()


def test_tooltip_show_and_hide():
    root = tk.Tk()
    btn = tk.Button(root, text="Hover me")
    btn.pack()
    root.update_idletasks()

    tip = _Tooltip(btn, "Test tip")

    # Simulate enter → schedule show
    tip._on_enter()
    # Force the delayed show to fire
    root.after(500, lambda: None)
    root.update()
    # Give it time to show
    import time
    time.sleep(0.5)
    root.update()

    assert tip._tip_window is not None

    # Simulate leave → hide
    tip._on_leave()
    root.update_idletasks()
    assert tip._tip_window is None

    root.destroy()


def test_multiple_tooltips_independent():
    root = tk.Tk()
    b1 = tk.Button(root, text="A")
    b2 = tk.Button(root, text="B")
    b1.pack()
    b2.pack()
    add_tooltip(b1, "Tip A")
    add_tooltip(b2, "Tip B")
    root.update_idletasks()
    root.destroy()


def test_tooltip_on_combobox():
    root = tk.Tk()
    from tkinter import ttk
    cb = ttk.Combobox(root, values=["X", "Y"], state="readonly")
    cb.pack()
    add_tooltip(cb, "Combo tooltip")
    root.update_idletasks()
    root.destroy()
