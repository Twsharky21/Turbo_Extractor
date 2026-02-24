"""
gui/tooltip.py — Lightweight hover tooltip for Tkinter widgets.

Usage:
    from gui.tooltip import add_tooltip
    add_tooltip(some_widget, "This field does X. Optional.")

Purely additive — no layout, style, or event changes to the target widget.
"""
from __future__ import annotations

import tkinter as tk


_DELAY_MS = 400        # hover delay before showing
_BG       = "#ffffe1"  # classic tooltip yellow
_FG       = "#222222"
_FONT     = ("Segoe UI", 9, "normal")
_PAD_X    = 8
_PAD_Y    = 4
_WRAP     = 280        # wraplength in px


class _Tooltip:
    """Attach a hover tooltip to an existing widget."""

    def __init__(self, widget: tk.Widget, text: str) -> None:
        self._widget = widget
        self._text = text
        self._tip_window: tk.Toplevel | None = None
        self._after_id: str | None = None

        widget.bind("<Enter>", self._on_enter, add="+")
        widget.bind("<Leave>", self._on_leave, add="+")
        widget.bind("<ButtonPress>", self._on_leave, add="+")

    def _on_enter(self, event=None) -> None:
        self._cancel()
        self._after_id = self._widget.after(_DELAY_MS, self._show)

    def _on_leave(self, event=None) -> None:
        self._cancel()
        self._hide()

    def _cancel(self) -> None:
        if self._after_id is not None:
            try:
                self._widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    def _show(self) -> None:
        if self._tip_window is not None:
            return
        try:
            x = self._widget.winfo_rootx() + 20
            y = self._widget.winfo_rooty() + self._widget.winfo_height() + 4
        except Exception:
            return

        tw = tk.Toplevel(self._widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")

        # Remove window shadow on some platforms
        try:
            tw.wm_attributes("-topmost", True)
        except Exception:
            pass

        label = tk.Label(
            tw,
            text=self._text,
            justify="left",
            background=_BG,
            foreground=_FG,
            font=_FONT,
            relief="solid",
            borderwidth=1,
            padx=_PAD_X,
            pady=_PAD_Y,
            wraplength=_WRAP,
        )
        label.pack()
        self._tip_window = tw

    def _hide(self) -> None:
        if self._tip_window is not None:
            try:
                self._tip_window.destroy()
            except Exception:
                pass
            self._tip_window = None


def add_tooltip(widget: tk.Widget, text: str) -> None:
    """Attach a tooltip to any Tkinter widget. Safe no-op if widget is None."""
    if widget is not None and text:
        _Tooltip(widget, text)
