"""
gui/mixins/throbber_mixin.py — Animated spinner widget and app mixin.

Throbber(tk.Canvas):
  - Idle state: shows a small static grey dot
  - Running state: draws a spinning arc that rotates smoothly at 50ms intervals

ThrobberMixin:
  - throbber_start() / throbber_stop() — called by app.run_all / run_selected_sheet
  - Graceful no-op if self.throbber is not set (safe for tests)
"""
from __future__ import annotations

import tkinter as tk


_SIZE = 22          # canvas width/height in px
_LINE_W = 3         # arc stroke width
_IDLE_COLOR = "#cccccc"
_ARC_COLOR = "#1f76ff"
_BG_RING = "#e0e0e0"
_INTERVAL_MS = 50
_ARC_EXTENT = 80    # degrees of the spinning arc
_STEP_DEG = 18      # rotation per tick


class Throbber(tk.Canvas):
    """Animated spinning arc indicator."""

    def __init__(self, master, **kw):
        kw.setdefault("width", _SIZE)
        kw.setdefault("height", _SIZE)
        kw.setdefault("highlightthickness", 0)
        kw.setdefault("borderwidth", 0)
        super().__init__(master, **kw)
        self._running = False
        self._angle = 0
        self._after_id = None
        self._draw_idle()

    # ── Public API ────────────────────────────────────────────────────────────

    @property
    def running(self) -> bool:
        return self._running

    def start(self) -> None:
        """Begin animation. Idempotent — calling while running is a no-op."""
        if self._running:
            return
        self._running = True
        self._angle = 0
        self._tick()

    def stop(self) -> None:
        """Stop animation and reset to idle indicator."""
        self._running = False
        if self._after_id is not None:
            try:
                self.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None
        self._draw_idle()

    # ── Drawing ───────────────────────────────────────────────────────────────

    def _draw_idle(self) -> None:
        self.delete("all")
        cx, cy = _SIZE // 2, _SIZE // 2
        r = 4
        self.create_oval(cx - r, cy - r, cx + r, cy + r,
                         fill=_IDLE_COLOR, outline=_IDLE_COLOR)

    def _draw_spinning(self) -> None:
        self.delete("all")
        pad = 3
        # Background ring
        self.create_oval(pad, pad, _SIZE - pad, _SIZE - pad,
                         outline=_BG_RING, width=_LINE_W)
        # Spinning arc
        self.create_arc(pad, pad, _SIZE - pad, _SIZE - pad,
                        start=self._angle, extent=_ARC_EXTENT,
                        outline=_ARC_COLOR, width=_LINE_W, style="arc")

    def _tick(self) -> None:
        if not self._running:
            return
        self._draw_spinning()
        self._angle = (self._angle + _STEP_DEG) % 360
        self._after_id = self.after(_INTERVAL_MS, self._tick)


class ThrobberMixin:
    """
    Mixin for TurboExtractorApp: start/stop the throbber widget.

    Expects self.throbber to be a Throbber instance (set by ui_build.py).
    Gracefully no-ops if the widget doesn't exist yet.
    """

    def throbber_start(self) -> None:
        throbber = getattr(self, "throbber", None)
        if throbber is not None:
            throbber.start()

    def throbber_stop(self) -> None:
        throbber = getattr(self, "throbber", None)
        if throbber is not None:
            throbber.stop()
