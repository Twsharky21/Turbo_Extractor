\
from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional

from core.project import ProjectConfig, SourceConfig, RecipeConfig
from core.models import SheetConfig, Destination, Rule


class TurboExtractorApp(tk.Tk):
    """
    V3 GUI — Rules UI added.
    """

    def __init__(self) -> None:
        super().__init__()
        self.title("Excel Turbo Extractor V3")
        self.minsize(1100, 700)

        self.project: ProjectConfig = ProjectConfig()
        self.current_sheet: Optional[SheetConfig] = None

        self._build_ui()

    # ---------------- UI ----------------

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        root = ttk.Frame(self, padding=10)
        root.grid(row=0, column=0, sticky="nsew")
        root.columnconfigure(0, weight=1)
        root.columnconfigure(1, weight=3)
        root.rowconfigure(0, weight=1)

        # LEFT
        left = ttk.LabelFrame(root, text="Sources / Recipes / Sheets", padding=8)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(left, show="tree", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        # RIGHT
        right = ttk.Frame(root)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)

        # ---- Sheet Section ----
        sheet_box = ttk.LabelFrame(right, text="Selected Sheet", padding=8)
        sheet_box.grid(row=0, column=0, sticky="ew")
        sheet_box.columnconfigure(1, weight=1)

        ttk.Label(sheet_box, text="Columns:").grid(row=0, column=0, sticky="w")
        self.columns_var = tk.StringVar()
        ttk.Entry(sheet_box, textvariable=self.columns_var).grid(row=0, column=1, sticky="ew")
        self.columns_var.trace_add("write", self._push_editor_to_sheet)

        ttk.Label(sheet_box, text="Rows:").grid(row=1, column=0, sticky="w")
        self.rows_var = tk.StringVar()
        ttk.Entry(sheet_box, textvariable=self.rows_var).grid(row=1, column=1, sticky="ew")
        self.rows_var.trace_add("write", self._push_editor_to_sheet)

        ttk.Label(sheet_box, text="Combine:").grid(row=2, column=0, sticky="w")
        self.combine_var = tk.StringVar()
        combine = ttk.Combobox(sheet_box, textvariable=self.combine_var,
                               values=["AND", "OR"], state="readonly")
        combine.grid(row=2, column=1, sticky="ew")
        combine.bind("<<ComboboxSelected>>", self._push_editor_to_sheet)

        # ---- Rules Section ----
        rules_box = ttk.LabelFrame(right, text="Rules", padding=8)
        rules_box.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        rules_box.columnconfigure(0, weight=1)
        rules_box.rowconfigure(0, weight=1)

        self.rules_canvas = tk.Canvas(rules_box, height=200)
        self.rules_canvas.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(rules_box, orient="vertical",
                                  command=self.rules_canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.rules_canvas.configure(yscrollcommand=scrollbar.set)

        self.rules_frame = ttk.Frame(self.rules_canvas)
        self.rules_canvas.create_window((0, 0), window=self.rules_frame, anchor="nw")
        self.rules_frame.bind("<Configure>",
                              lambda e: self.rules_canvas.configure(
                                  scrollregion=self.rules_canvas.bbox("all")))

        ttk.Button(rules_box, text="+ Add Rule",
                   command=self.add_rule).grid(row=1, column=0, sticky="w", pady=(6, 0))

    # ---------------- Tree ----------------

    def _on_tree_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        path = self._get_tree_path(sel[0])
        if len(path) == 3:
            self.current_sheet = self.project.sources[path[0]].recipes[path[1]].sheets[path[2]]
            self._load_sheet()
        else:
            self.current_sheet = None
            self._clear_editor()

    def _get_tree_path(self, item_id):
        path = []
        current = item_id
        while current:
            parent = self.tree.parent(current)
            siblings = self.tree.get_children(parent)
            path.insert(0, list(siblings).index(current))
            current = parent
        return path

    def refresh_tree(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for source in self.project.sources:
            s_id = self.tree.insert("", "end", text=source.path)
            for recipe in source.recipes:
                r_id = self.tree.insert(s_id, "end", text=recipe.name)
                for sheet in recipe.sheets:
                    self.tree.insert(r_id, "end", text=sheet.name)

    # ---------------- Editor ----------------

    def _load_sheet(self):
        s = self.current_sheet
        self.columns_var.set(s.columns_spec)
        self.rows_var.set(s.rows_spec)
        self.combine_var.set(s.rules_combine)
        self._rebuild_rules()

    def _clear_editor(self):
        self.columns_var.set("")
        self.rows_var.set("")
        self.combine_var.set("")
        for child in self.rules_frame.winfo_children():
            child.destroy()

    def _push_editor_to_sheet(self, *args):
        if not self.current_sheet:
            return
        self.current_sheet.columns_spec = self.columns_var.get()
        self.current_sheet.rows_spec = self.rows_var.get()
        if self.combine_var.get():
            self.current_sheet.rules_combine = self.combine_var.get()

    # ---------------- Rules UI ----------------

    def add_rule(self):
        if not self.current_sheet:
            return
        rule = Rule(mode="include", column="A",
                    operator="equals", value="")
        self.current_sheet.rules.append(rule)
        self._rebuild_rules()

    def _rebuild_rules(self):
        for child in self.rules_frame.winfo_children():
            child.destroy()

        if not self.current_sheet:
            return

        for idx, rule in enumerate(self.current_sheet.rules):
            self._build_rule_row(idx, rule)

    def _build_rule_row(self, idx: int, rule: Rule):
        frame = ttk.Frame(self.rules_frame)
        frame.grid(row=idx, column=0, sticky="ew", pady=2)
        frame.columnconfigure(3, weight=1)

        mode_var = tk.StringVar(value=rule.mode)
        col_var = tk.StringVar(value=rule.column)
        op_var = tk.StringVar(value=rule.operator)
        val_var = tk.StringVar(value=rule.value)

        ttk.Combobox(frame, textvariable=mode_var,
                     values=["include", "exclude"],
                     state="readonly").grid(row=0, column=0)

        ttk.Entry(frame, textvariable=col_var, width=5).grid(row=0, column=1)
        ttk.Combobox(frame, textvariable=op_var,
                     values=["equals", "contains", "<", ">"],
                     state="readonly").grid(row=0, column=2)
        ttk.Entry(frame, textvariable=val_var).grid(row=0, column=3, sticky="ew")

        ttk.Button(frame, text="X",
                   command=lambda i=idx: self._remove_rule(i)).grid(row=0, column=4)

        def push(*_):
            rule.mode = mode_var.get()
            rule.column = col_var.get()
            rule.operator = op_var.get()
            rule.value = val_var.get()

        mode_var.trace_add("write", push)
        col_var.trace_add("write", push)
        op_var.trace_add("write", push)
        val_var.trace_add("write", push)

    def _remove_rule(self, idx: int):
        if not self.current_sheet:
            return
        del self.current_sheet.rules[idx]
        self._rebuild_rules()


def main():
    app = TurboExtractorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
