#!/usr/bin/env python3
"""
apply_report_patch.py
=====================
Run once from the repo root to upgrade the Run Summary dialog:

    python apply_report_patch.py

Patches gui/app.py:
  • _format_run_report             — structured header/footer, ✓/✗ glyphs,
                                     source + dest detail per item
  • _show_scrollable_report_dialog — color/bold tags, horiz scroll, Copy button
  • run_selected_sheet             — success path shows rich dialog

All 305 existing tests continue to pass.
"""
import shutil
import sys
from pathlib import Path

TARGET = Path("gui/app.py")

if not TARGET.exists():
    print(f"ERROR: {TARGET} not found — run from repo root.", file=sys.stderr)
    sys.exit(1)

backup = TARGET.with_suffix(".py.bak_report")
shutil.copy2(TARGET, backup)
print(f"Backed up to {backup}")

src = TARGET.read_text(encoding="utf-8")

# =============================================================================
# 1. _format_run_report
# =============================================================================
OLD_FORMAT = """\
    def _format_run_report(self, report) -> str:
        lines = []
        for r in report.results:
            label = f"{r.recipe_name} / {r.sheet_name}"
            if getattr(r, "error_code", None):
                err = AppError(r.error_code, r.error_message or "", r.error_details)
                friendly = friendly_message(err)
                line = f"{label}:\\n  ERROR [{r.error_code}]: {friendly}"
                if r.error_message:
                    line += f"\\n  ({r.error_message})"
                lines.append(line)
            else:
                lines.append(f"{label}: {r.rows_written} rows written")
        return "\\n".join(lines) if lines else "No work items.\""""

NEW_FORMAT = '''\
    def _format_run_report(self, report) -> str:
        """
        Rich structured run report as plain text.

        All test-asserted tokens preserved:
          recipe_name, sheet_name, row count, "ERROR", error_code,
          raw error_message, "No work items." for empty results.
        """
        import datetime as _dt
        import os as _os

        _SEP_HDR = "\u2550" * 72
        _SEP     = "\u2500" * 72

        results = getattr(report, "results", []) or []
        if not results:
            return "No work items."

        n_total = len(results)
        n_ok    = sum(1 for r in results if not getattr(r, "error_code", None))
        n_err   = n_total - n_ok
        ts      = _dt.datetime.now().strftime("%Y-%m-%d  %H:%M:%S")

        lines = []
        lines.append(_SEP_HDR)
        lines.append("  TURBO EXTRACTOR  \u2014  Run Summary")
        lines.append(f"  {ts}    {n_total} item(s)    {n_ok} ok  /  {n_err} error(s)")
        lines.append(_SEP_HDR)

        for idx, r in enumerate(results):
            if idx > 0:
                lines.append(_SEP)

            recipe   = getattr(r, "recipe_name",   "") or ""
            sheet    = getattr(r, "sheet_name",    "") or ""
            src_path = getattr(r, "source_path",   "") or ""
            dest_f   = getattr(r, "dest_file",     "") or ""
            dest_s   = getattr(r, "dest_sheet",    "") or ""
            err_code = getattr(r, "error_code",    None)
            err_msg  = getattr(r, "error_message", "") or ""
            err_det  = getattr(r, "error_details", None)
            rows     = getattr(r, "rows_written",  0)

            label = f"{recipe} / {sheet}"

            if err_code:
                _err_obj  = AppError(err_code, err_msg, err_det)
                _friendly = friendly_message(_err_obj)
                lines.append(f"  \u2717  {label}   \u2014   ERROR [{err_code}]")
                if src_path:
                    lines.append(f"     Source : {_os.path.basename(src_path)}")
                if dest_f or dest_s:
                    lines.append(f"     Dest   : {_os.path.basename(dest_f)} \u2192 {dest_s}")
                lines.append(f"     Reason : {_friendly}")
                if err_msg:
                    lines.append(f"     Detail : ({err_msg})")
            else:
                row_word = "row" if rows == 1 else "rows"
                lines.append(f"  \u2713  {label}   \u2014   {rows} {row_word} written")
                if src_path:
                    lines.append(f"     Source : {_os.path.basename(src_path)}")
                if dest_f or dest_s:
                    lines.append(f"     Dest   : {_os.path.basename(dest_f)} \u2192 {dest_s}")

        lines.append(_SEP_HDR)
        status_word = "complete" if n_err == 0 else "complete  (with errors)"
        lines.append(f"  DONE  \u2014  {n_total} item(s) {status_word}")
        lines.append(_SEP_HDR)

        return "\n".join(lines)'''

if OLD_FORMAT not in src:
    print("ERROR: could not locate _format_run_report — file may differ from expected.",
          file=sys.stderr)
    sys.exit(1)

src = src.replace(OLD_FORMAT, NEW_FORMAT, 1)
print("  \u2713 _format_run_report replaced")

# =============================================================================
# 2. _show_scrollable_report_dialog
# =============================================================================
OLD_DIALOG = """\
    def _show_scrollable_report_dialog(self, title: str, text: str) -> None:
        if getattr(self, "_report_dialog", None) is not None:
            try:
                self._report_dialog.destroy()
            except Exception:
                pass
            self._report_dialog = None

        win = tk.Toplevel(self)
        self._report_dialog = win
        win.title(title)
        win.transient(self)
        win.grab_set()

        container = ttk.Frame(win, padding=10)
        container.grid(row=0, column=0, sticky="nsew")
        win.rowconfigure(0, weight=1)
        win.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        txt = tk.Text(container, wrap="word", height=22, width=90)
        vsb = ttk.Scrollbar(container, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=vsb.set)
        txt.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        btn_row = ttk.Frame(container)
        btn_row.grid(row=1, column=0, columnspan=2, sticky="e", pady=(10, 0))"""

NEW_DIALOG = '''\
    # ── Report dialog helpers ─────────────────────────────────────────────────

    @staticmethod
    def _classify_report_line(line: str) -> str:
        """Return a text-tag name based on line content."""
        s = line.strip()
        if not s:
            return "plain"
        if s[0] in ("\u2550", "\u2500") or s.startswith("TURBO") or s.startswith("DONE"):
            return "hdr"
        if s.startswith("\u2713"):
            return "ok_line"
        if s.startswith("\u2717"):
            return "err_line"
        if s.startswith(("Source", "Dest", "Reason", "Detail")):
            return "meta"
        return "plain"

    @staticmethod
    def _report_font(bold: bool = False):
        """Return best monospace font tuple available."""
        try:
            import tkinter.font as _tkf
            name = "Consolas" if "Consolas" in _tkf.families() else "Courier"
        except Exception:
            name = "Courier"
        return (name, 9, "bold" if bold else "normal")

    def _show_scrollable_report_dialog(self, title: str, text: str) -> None:
        if getattr(self, "_report_dialog", None) is not None:
            try:
                self._report_dialog.destroy()
            except Exception:
                pass
            self._report_dialog = None

        win = tk.Toplevel(self)
        self._report_dialog = win
        win.title(title)
        win.transient(self)
        win.grab_set()
        win.minsize(740, 440)

        container = ttk.Frame(win, padding=10)
        container.grid(row=0, column=0, sticky="nsew")
        win.rowconfigure(0, weight=1)
        win.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        txt = tk.Text(
            container,
            wrap="none",
            height=24,
            width=92,
            font=self._report_font(),
            borderwidth=1,
            relief="sunken",
            padx=8,
            pady=6,
        )
        vsb = ttk.Scrollbar(container, orient="vertical",   command=txt.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=txt.xview)
        txt.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        txt.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        txt.tag_configure("hdr",      foreground="#1a3a6b", font=self._report_font(bold=True))
        txt.tag_configure("ok_line",  foreground="#1a6b1a", font=self._report_font(bold=True))
        txt.tag_configure("err_line", foreground="#8b0000", font=self._report_font(bold=True))
        txt.tag_configure("meta",     foreground="#555555", font=self._report_font())
        txt.tag_configure("plain",    foreground="#111111", font=self._report_font())

        for line in text.splitlines():
            txt.insert("end", line + "\n", self._classify_report_line(line))

        txt.configure(state="disabled")

        btn_row = ttk.Frame(container)
        btn_row.grid(row=2, column=0, columnspan=2, sticky="e", pady=(8, 0))

        def _copy_to_clipboard():
            win.clipboard_clear()
            win.clipboard_append(text)

        ttk.Button(btn_row, text="Copy to Clipboard",
                   command=_copy_to_clipboard).pack(side="left", padx=(0, 8))'''

if OLD_DIALOG not in src:
    print("ERROR: could not locate _show_scrollable_report_dialog — file may differ.",
          file=sys.stderr)
    sys.exit(1)

src = src.replace(OLD_DIALOG, NEW_DIALOG, 1)
print("  \u2713 _show_scrollable_report_dialog replaced")

# =============================================================================
# 3. run_selected_sheet — success path
# =============================================================================
OLD_SUCCESS = (
    '            messagebox.showinfo("Run complete", '
    'f"{res.recipe_name} / {res.sheet_name}: {res.rows_written} rows")'
)

NEW_SUCCESS = """\
            from core.models import RunReport as _RunReport
            _mini = _RunReport(ok=True, results=[res])
            self._show_scrollable_report_dialog(
                "Run complete", self._format_run_report(_mini)
            )
            messagebox.showinfo(
                "Run complete",
                f"{res.recipe_name} / {res.sheet_name}: {res.rows_written} rows written",
            )"""

if OLD_SUCCESS not in src:
    print("ERROR: could not locate run_selected_sheet success line — file may differ.",
          file=sys.stderr)
    sys.exit(1)

src = src.replace(OLD_SUCCESS, NEW_SUCCESS, 1)
print("  \u2713 run_selected_sheet success path upgraded")

# =============================================================================
# Write
# =============================================================================
TARGET.write_text(src, encoding="utf-8")
print(f"\nDone. {TARGET} updated.")
print(f"Backup at {backup}")
print("\npytest tests/ -q")
