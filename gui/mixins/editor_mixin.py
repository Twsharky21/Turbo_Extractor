from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from core.models import SheetConfig, Destination, Rule


class EditorMixin:
    """
    Mixin for TurboExtractorApp: right-panel sheet editor and rules UI.

    Covers: _make_default_sheet, _load_sheet_into_editor,
    _do_load_sheet_into_editor, _clear_editor, _push_editor_to_sheet,
    _rebuild_rules, _build_rule_row, add_rule, _remove_rule.
    """

    def _make_default_sheet(self, name: str) -> SheetConfig:
        return SheetConfig(
            name=name,
            workbook_sheet=name,
            source_start_row="",
            columns_spec="",
            rows_spec="",
            paste_mode="pack",
            rules_combine="AND",
            rules=[],
            destination=Destination(
                file_path="",
                sheet_name="Sheet1",
                start_col="A",
                start_row="",
            ),
        )

    def _load_sheet_into_editor(self, sheet: SheetConfig) -> None:
        self._loading = True
        try:
            self._do_load_sheet_into_editor(sheet)
        finally:
            self._loading = False

    def _do_load_sheet_into_editor(self, sheet: SheetConfig) -> None:
        self.columns_var.set(sheet.columns_spec)
        self.rows_var.set(sheet.rows_spec)
        self.source_start_row_var.set(getattr(sheet, "source_start_row", ""))
        paste_val = sheet.paste_mode
        if paste_val == "pack":
            paste_val = "Pack Together"
        elif paste_val == "keep":
            paste_val = "Keep Format"
        self.paste_var.set(paste_val)
        self.combine_var.set(sheet.rules_combine)

        self.dest_file_var.set(sheet.destination.file_path)
        self.dest_sheet_var.set(sheet.destination.sheet_name)
        self.start_col_var.set(sheet.destination.start_col)
        self.start_row_var.set(sheet.destination.start_row)

        self._rebuild_rules()

    def _clear_editor(self) -> None:
        self.columns_var.set("")
        self.rows_var.set("")
        self.source_start_row_var.set("")
        self.paste_var.set("")
        self.combine_var.set("")
        self.dest_file_var.set("")
        self.dest_sheet_var.set("")
        self.start_col_var.set("")
        self.start_row_var.set("")

        for child in self.rules_frame.winfo_children():
            child.destroy()

    def _push_editor_to_sheet(self, *args) -> None:
        if self._loading:
            return
        sel = self.tree.selection()
        if not sel:
            return
        path = self._get_tree_path(sel[0])
        if len(path) != 3:
            return

        sheet = self.project.sources[path[0]].recipes[path[1]].sheets[path[2]]

        sheet.columns_spec = self.columns_var.get()
        sheet.rows_spec = self.rows_var.get()
        sheet.source_start_row = self.source_start_row_var.get()
        val = self.paste_var.get().strip()
        if val:
            if val.lower().startswith("pack"):
                sheet.paste_mode = "pack"
            elif val.lower().startswith("keep"):
                sheet.paste_mode = "keep"
            else:
                sheet.paste_mode = val
        if self.combine_var.get():
            sheet.rules_combine = self.combine_var.get()

        sheet.destination.file_path = self.dest_file_var.get()
        sheet.destination.sheet_name = self.dest_sheet_var.get()
        sheet.destination.start_col = self.start_col_var.get()
        sheet.destination.start_row = self.start_row_var.get()

        self._mark_dirty()

    def _rebuild_rules(self) -> None:
        for child in self.rules_frame.winfo_children():
            child.destroy()

        sel = self.tree.selection()
        if not sel:
            return
        path = self._get_tree_path(sel[0])
        if len(path) != 3:
            return
        sheet = self.project.sources[path[0]].recipes[path[1]].sheets[path[2]]

        for idx, rule in enumerate(sheet.rules):
            self._build_rule_row(idx, rule)

    def _build_rule_row(self, idx: int, rule: Rule) -> None:
        row = ttk.Frame(self.rules_frame)
        row.grid(row=idx, column=0, sticky="ew", pady=2)
        row.columnconfigure(0, minsize=100)
        row.columnconfigure(1, minsize=64)
        row.columnconfigure(2, minsize=112)
        row.columnconfigure(3, weight=1)

        # Display capitalized values; model stores lowercase
        mode_display = rule.mode.capitalize()
        op_display = rule.operator.capitalize() if rule.operator in ("equals", "contains") else rule.operator

        mode_var = tk.StringVar(value=mode_display)
        col_var  = tk.StringVar(value=rule.column)
        op_var   = tk.StringVar(value=op_display)
        val_var  = tk.StringVar(value=rule.value)

        ttk.Combobox(row, textvariable=mode_var, values=["Include", "Exclude"],
                     state="readonly", style="White.TCombobox", width=9).grid(row=0, column=0, sticky="w")
        col_entry = ttk.Entry(row, textvariable=col_var, width=6)
        col_entry.grid(row=0, column=1, sticky="w", padx=(6, 0))
        ttk.Combobox(row, textvariable=op_var, values=["Equals", "Contains", "<", ">"],
                     state="readonly", style="White.TCombobox", width=10).grid(row=0, column=2, sticky="w", padx=(6, 0))
        ttk.Entry(row, textvariable=val_var).grid(row=0, column=3, sticky="ew", padx=(6, 0))

        ttk.Button(row, text="X", command=lambda i=idx: self._remove_rule(i),
                   width=3).grid(row=0, column=4, padx=(6, 0))

        # Auto-capitalize column letters
        def _cap_rule_col(*_):
            v = col_var.get()
            up = v.upper()
            if v != up:
                col_var.set(up)

        def push(*_):
            # Map display values back to model (lowercase)
            mode_val = mode_var.get().strip().lower()
            if mode_val in ("include", "exclude"):
                rule.mode = mode_val
            op_val = op_var.get().strip()
            if op_val.lower() in ("equals", "contains"):
                rule.operator = op_val.lower()
            else:
                rule.operator = op_val
            rule.column   = col_var.get()
            rule.value    = val_var.get()
            self._mark_dirty()

        mode_var.trace_add("write", push)
        col_var.trace_add("write", _cap_rule_col)
        col_var.trace_add("write",  push)
        op_var.trace_add("write",   push)
        val_var.trace_add("write",  push)

    def add_rule(self) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        path = self._get_tree_path(sel[0])
        if len(path) != 3:
            return
        sheet = self.project.sources[path[0]].recipes[path[1]].sheets[path[2]]
        sheet.rules.append(Rule(mode="include", column="A", operator="contains", value=""))
        self._rebuild_rules()
        self._mark_dirty()

    def _remove_rule(self, idx: int) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        path = self._get_tree_path(sel[0])
        if len(path) != 3:
            return
        sheet = self.project.sources[path[0]].recipes[path[1]].sheets[path[2]]
        if 0 <= idx < len(sheet.rules):
            del sheet.rules[idx]
        self._rebuild_rules()
        self._mark_dirty()
