\
from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional

from core.project import ProjectConfig, SourceConfig, RecipeConfig
from core.models import SheetConfig, Destination, Rule


class TurboExtractorApp(tk.Tk):
    """
    V3 GUI (merged): Tree structure + minimal sheet editor + rules UI.

    Goals:
    - Keep ProjectConfig as the single source of truth.
    - Tree reflects Sources -> Recipes -> Sheets.
    - Selecting a Sheet loads editor; selecting Source/Recipe clears editor.
    - Add/Remove structure implemented.
    - Rules UI supports add/remove + live binding.
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

        # ----- LEFT: Tree + buttons -----
        left = ttk.LabelFrame(root, text="Sources / Recipes / Sheets", padding=8)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(left, show="tree", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        yscroll = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=yscroll.set)

        btns = ttk.Frame(left)
        btns.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))

        ttk.Button(btns, text="Add Source(s)...", command=self.add_sources).pack(side="left")
        ttk.Button(btns, text="Add Recipe", command=self.add_recipe).pack(side="left", padx=(6, 0))
        ttk.Button(btns, text="Add Sheet", command=self.add_sheet).pack(side="left", padx=(6, 0))
        ttk.Button(btns, text="Remove Selected", command=self.remove_selected).pack(side="left", padx=(6, 0))

        # ----- RIGHT: Editor -----
        right = ttk.Frame(root)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        # Sheet box
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

        ttk.Label(sheet_box, text="Paste Mode:").grid(row=2, column=0, sticky="w")
        self.paste_var = tk.StringVar()
        self.paste_combo = ttk.Combobox(sheet_box, textvariable=self.paste_var, values=["pack", "keep"], state="readonly")
        self.paste_combo.grid(row=2, column=1, sticky="ew")
        self.paste_combo.bind("<<ComboboxSelected>>", self._push_editor_to_sheet)

        ttk.Label(sheet_box, text="Rules Combine:").grid(row=3, column=0, sticky="w")
        self.combine_var = tk.StringVar()
        self.combine_combo = ttk.Combobox(sheet_box, textvariable=self.combine_var, values=["AND", "OR"], state="readonly")
        self.combine_combo.grid(row=3, column=1, sticky="ew")
        self.combine_combo.bind("<<ComboboxSelected>>", self._push_editor_to_sheet)

        # Destination minimal (kept small for now)
        dest_box = ttk.LabelFrame(right, text="Destination", padding=8)
        dest_box.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        dest_box.columnconfigure(1, weight=1)

        ttk.Label(dest_box, text="File:").grid(row=0, column=0, sticky="w")
        self.dest_file_var = tk.StringVar()
        ttk.Entry(dest_box, textvariable=self.dest_file_var).grid(row=0, column=1, sticky="ew")
        self.dest_file_var.trace_add("write", self._push_editor_to_sheet)

        ttk.Label(dest_box, text="Sheet Name:").grid(row=1, column=0, sticky="w")
        self.dest_sheet_var = tk.StringVar()
        ttk.Entry(dest_box, textvariable=self.dest_sheet_var).grid(row=1, column=1, sticky="ew")
        self.dest_sheet_var.trace_add("write", self._push_editor_to_sheet)

        ttk.Label(dest_box, text="Start Col:").grid(row=2, column=0, sticky="w")
        self.start_col_var = tk.StringVar()
        ttk.Entry(dest_box, textvariable=self.start_col_var, width=8).grid(row=2, column=1, sticky="w")
        self.start_col_var.trace_add("write", self._push_editor_to_sheet)

        ttk.Label(dest_box, text="Start Row:").grid(row=3, column=0, sticky="w")
        self.start_row_var = tk.StringVar()
        ttk.Entry(dest_box, textvariable=self.start_row_var, width=8).grid(row=3, column=1, sticky="w")
        self.start_row_var.trace_add("write", self._push_editor_to_sheet)

        # Rules section (scrollable)
        rules_box = ttk.LabelFrame(right, text="Rules", padding=8)
        rules_box.grid(row=2, column=0, sticky="nsew", pady=(10, 0))
        rules_box.columnconfigure(0, weight=1)
        rules_box.rowconfigure(0, weight=1)
        right.rowconfigure(2, weight=1)

        self.rules_canvas = tk.Canvas(rules_box, height=220)
        self.rules_canvas.grid(row=0, column=0, sticky="nsew")

        rules_scroll = ttk.Scrollbar(rules_box, orient="vertical", command=self.rules_canvas.yview)
        rules_scroll.grid(row=0, column=1, sticky="ns")
        self.rules_canvas.configure(yscrollcommand=rules_scroll.set)

        self.rules_frame = ttk.Frame(self.rules_canvas)
        self.rules_canvas.create_window((0, 0), window=self.rules_frame, anchor="nw")

        self.rules_frame.bind(
            "<Configure>",
            lambda e: self.rules_canvas.configure(scrollregion=self.rules_canvas.bbox("all")),
        )

        ttk.Button(rules_box, text="+ Add Rule", command=self.add_rule).grid(row=1, column=0, sticky="w", pady=(6, 0))

        self._clear_editor()

    # ---------------- Tree helpers ----------------

    def refresh_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        for source in self.project.sources:
            s_id = self.tree.insert("", "end", text=source.path)
            for recipe in source.recipes:
                r_id = self.tree.insert(s_id, "end", text=recipe.name)
                for sheet in recipe.sheets:
                    self.tree.insert(r_id, "end", text=sheet.name)

    def _get_tree_path(self, item_id):
        path = []
        current = item_id
        while current:
            parent = self.tree.parent(current)
            siblings = self.tree.get_children(parent)
            path.insert(0, list(siblings).index(current))
            current = parent
        return path

    # ---------------- Selection ----------------

    def _on_tree_select(self, event=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return

        path = self._get_tree_path(sel[0])
        if len(path) == 3:
            self.current_sheet = self.project.sources[path[0]].recipes[path[1]].sheets[path[2]]
            self._load_sheet_into_editor(self.current_sheet)
        else:
            self.current_sheet = None
            self._clear_editor()

    # ---------------- Structure actions ----------------

    def add_sources(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Add source file(s)",
            filetypes=[("Excel/CSV", "*.xlsx *.xlsm *.csv"), ("All files", "*.*")],
        )
        if not paths:
            return

        for p in paths:
            sheet = self._make_default_sheet(name="Sheet1")
            recipe = RecipeConfig(name="Recipe1", sheets=[sheet])
            self.project.sources.append(SourceConfig(path=p, recipes=[recipe]))

        self.refresh_tree()

    def add_recipe(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Select Source", "Select a Source to add a Recipe.")
            return

        path = self._get_tree_path(sel[0])
        if len(path) != 1:
            messagebox.showwarning("Select Source", "Select a Source to add a Recipe.")
            return

        source = self.project.sources[path[0]]
        source.recipes.append(RecipeConfig(name=f"Recipe{len(source.recipes) + 1}", sheets=[]))
        self.refresh_tree()

    def add_sheet(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Select Recipe", "Select a Recipe to add a Sheet.")
            return

        path = self._get_tree_path(sel[0])
        if len(path) != 2:
            messagebox.showwarning("Select Recipe", "Select a Recipe to add a Sheet.")
            return

        recipe = self.project.sources[path[0]].recipes[path[1]]
        recipe.sheets.append(self._make_default_sheet(name=f"Sheet{len(recipe.sheets) + 1}"))
        self.refresh_tree()

    def remove_selected(self) -> None:
        sel = self.tree.selection()
        if not sel:
            return

        path = self._get_tree_path(sel[0])

        if len(path) == 1:
            del self.project.sources[path[0]]
        elif len(path) == 2:
            del self.project.sources[path[0]].recipes[path[1]]
        elif len(path) == 3:
            source = self.project.sources[path[0]]
            recipe = source.recipes[path[1]]
            del recipe.sheets[path[2]]
            # auto-delete empty recipe
            if not recipe.sheets:
                del source.recipes[path[1]]

        self.current_sheet = None
        self.refresh_tree()
        self._clear_editor()

    def _make_default_sheet(self, name: str) -> SheetConfig:
        return SheetConfig(
            name=name,
            workbook_sheet="Sheet1",
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

    # ---------------- Editor binding ----------------

    def _load_sheet_into_editor(self, sheet: SheetConfig) -> None:
        self.columns_var.set(sheet.columns_spec)
        self.rows_var.set(sheet.rows_spec)
        self.paste_var.set(sheet.paste_mode)
        self.combine_var.set(sheet.rules_combine)

        self.dest_file_var.set(sheet.destination.file_path)
        self.dest_sheet_var.set(sheet.destination.sheet_name)
        self.start_col_var.set(sheet.destination.start_col)
        self.start_row_var.set(sheet.destination.start_row)

        self._rebuild_rules()

    def _clear_editor(self) -> None:
        self.columns_var.set("")
        self.rows_var.set("")
        self.paste_var.set("")
        self.combine_var.set("")
        self.dest_file_var.set("")
        self.dest_sheet_var.set("")
        self.start_col_var.set("")
        self.start_row_var.set("")

        for child in self.rules_frame.winfo_children():
            child.destroy()

    def _push_editor_to_sheet(self, *args) -> None:
        if not self.current_sheet:
            return

        self.current_sheet.columns_spec = self.columns_var.get()
        self.current_sheet.rows_spec = self.rows_var.get()
        if self.paste_var.get():
            self.current_sheet.paste_mode = self.paste_var.get()
        if self.combine_var.get():
            self.current_sheet.rules_combine = self.combine_var.get()

        self.current_sheet.destination.file_path = self.dest_file_var.get()
        self.current_sheet.destination.sheet_name = self.dest_sheet_var.get()
        self.current_sheet.destination.start_col = self.start_col_var.get()
        self.current_sheet.destination.start_row = self.start_row_var.get()

    # ---------------- Rules UI ----------------

    def add_rule(self) -> None:
        if not self.current_sheet:
            return
        self.current_sheet.rules.append(Rule(mode="include", column="A", operator="equals", value=""))
        self._rebuild_rules()

    def _remove_rule(self, idx: int) -> None:
        if not self.current_sheet:
            return
        del self.current_sheet.rules[idx]
        self._rebuild_rules()

    def _rebuild_rules(self) -> None:
        for child in self.rules_frame.winfo_children():
            child.destroy()

        if not self.current_sheet:
            return

        for idx, rule in enumerate(self.current_sheet.rules):
            self._build_rule_row(idx, rule)

    def _build_rule_row(self, idx: int, rule: Rule) -> None:
        row = ttk.Frame(self.rules_frame)
        row.grid(row=idx, column=0, sticky="ew", pady=2)
        row.columnconfigure(3, weight=1)

        mode_var = tk.StringVar(value=rule.mode)
        col_var = tk.StringVar(value=rule.column)
        op_var = tk.StringVar(value=rule.operator)
        val_var = tk.StringVar(value=rule.value)

        ttk.Combobox(row, textvariable=mode_var, values=["include", "exclude"], state="readonly", width=9).grid(row=0, column=0)
        ttk.Entry(row, textvariable=col_var, width=6).grid(row=0, column=1, padx=(6, 0))
        ttk.Combobox(row, textvariable=op_var, values=["equals", "contains", "<", ">"], state="readonly", width=10).grid(row=0, column=2, padx=(6, 0))
        ttk.Entry(row, textvariable=val_var).grid(row=0, column=3, sticky="ew", padx=(6, 0))

        ttk.Button(row, text="X", command=lambda i=idx: self._remove_rule(i), width=3).grid(row=0, column=4, padx=(6, 0))

        def push(*_):
            rule.mode = mode_var.get()
            rule.column = col_var.get()
            rule.operator = op_var.get()
            rule.value = val_var.get()

        mode_var.trace_add("write", push)
        col_var.trace_add("write", push)
        op_var.trace_add("write", push)
        val_var.trace_add("write", push)


def main() -> None:
    app = TurboExtractorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
