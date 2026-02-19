\
from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional

from core.engine import run_all
from core.project import ProjectConfig, SourceConfig, RecipeConfig
from core.models import SheetConfig, Destination


class TurboExtractorApp(tk.Tk):
    """
    V3 Minimal GUI with Sheet Editor wiring.
    """

    def __init__(self) -> None:
        super().__init__()
        self.title("Excel Turbo Extractor V3")
        self.minsize(1000, 650)

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

        # ----- LEFT TREE -----
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

        # ----- RIGHT PANEL -----
        right = ttk.LabelFrame(root, text="Selected Sheet", padding=8)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(1, weight=1)

        ttk.Label(right, text="Columns:").grid(row=0, column=0, sticky="w")
        self.columns_var = tk.StringVar()
        self.columns_entry = ttk.Entry(right, textvariable=self.columns_var)
        self.columns_entry.grid(row=0, column=1, sticky="ew")
        self.columns_var.trace_add("write", self._push_editor_to_sheet)

        ttk.Label(right, text="Rows:").grid(row=1, column=0, sticky="w")
        self.rows_var = tk.StringVar()
        self.rows_entry = ttk.Entry(right, textvariable=self.rows_var)
        self.rows_entry.grid(row=1, column=1, sticky="ew")
        self.rows_var.trace_add("write", self._push_editor_to_sheet)

        ttk.Label(right, text="Paste Mode:").grid(row=2, column=0, sticky="w")
        self.paste_var = tk.StringVar()
        self.paste_combo = ttk.Combobox(
            right, textvariable=self.paste_var,
            values=["pack", "keep"], state="readonly"
        )
        self.paste_combo.grid(row=2, column=1, sticky="ew")
        self.paste_combo.bind("<<ComboboxSelected>>", self._push_editor_to_sheet)

        ttk.Label(right, text="Dest File:").grid(row=3, column=0, sticky="w")
        self.dest_file_var = tk.StringVar()
        self.dest_file_entry = ttk.Entry(right, textvariable=self.dest_file_var)
        self.dest_file_entry.grid(row=3, column=1, sticky="ew")
        self.dest_file_var.trace_add("write", self._push_editor_to_sheet)

        ttk.Label(right, text="Dest Sheet:").grid(row=4, column=0, sticky="w")
        self.dest_sheet_var = tk.StringVar()
        self.dest_sheet_entry = ttk.Entry(right, textvariable=self.dest_sheet_var)
        self.dest_sheet_entry.grid(row=4, column=1, sticky="ew")
        self.dest_sheet_var.trace_add("write", self._push_editor_to_sheet)

        ttk.Label(right, text="Start Col:").grid(row=5, column=0, sticky="w")
        self.start_col_var = tk.StringVar()
        self.start_col_entry = ttk.Entry(right, textvariable=self.start_col_var)
        self.start_col_entry.grid(row=5, column=1, sticky="ew")
        self.start_col_var.trace_add("write", self._push_editor_to_sheet)

        ttk.Label(right, text="Start Row:").grid(row=6, column=0, sticky="w")
        self.start_row_var = tk.StringVar()
        self.start_row_entry = ttk.Entry(right, textvariable=self.start_row_var)
        self.start_row_entry.grid(row=6, column=1, sticky="ew")
        self.start_row_var.trace_add("write", self._push_editor_to_sheet)

    # ---------------- Tree Binding ----------------

    def _on_tree_select(self, event=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        item_id = sel[0]
        path = self._get_tree_path(item_id)

        # path depth: [source], [source, recipe], [source, recipe, sheet]
        if len(path) == 3:
            source_idx, recipe_idx, sheet_idx = path
            sheet = self.project.sources[source_idx].recipes[recipe_idx].sheets[sheet_idx]
            self.current_sheet = sheet
            self._load_sheet_into_editor(sheet)
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

    # ---------------- Editor Binding ----------------

    def _load_sheet_into_editor(self, sheet: SheetConfig) -> None:
        self.columns_var.set(sheet.columns_spec)
        self.rows_var.set(sheet.rows_spec)
        self.paste_var.set(sheet.paste_mode)
        self.dest_file_var.set(sheet.destination.file_path)
        self.dest_sheet_var.set(sheet.destination.sheet_name)
        self.start_col_var.set(sheet.destination.start_col)
        self.start_row_var.set(sheet.destination.start_row)

    def _clear_editor(self) -> None:
        self.columns_var.set("")
        self.rows_var.set("")
        self.paste_var.set("")
        self.dest_file_var.set("")
        self.dest_sheet_var.set("")
        self.start_col_var.set("")
        self.start_row_var.set("")

    def _push_editor_to_sheet(self, *args) -> None:
        if not self.current_sheet:
            return
        self.current_sheet.columns_spec = self.columns_var.get()
        self.current_sheet.rows_spec = self.rows_var.get()
        if self.paste_var.get():
            self.current_sheet.paste_mode = self.paste_var.get()
        self.current_sheet.destination.file_path = self.dest_file_var.get()
        self.current_sheet.destination.sheet_name = self.dest_sheet_var.get()
        self.current_sheet.destination.start_col = self.start_col_var.get()
        self.current_sheet.destination.start_row = self.start_row_var.get()

    # ---------------- Utility ----------------

    def refresh_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        for s_idx, source in enumerate(self.project.sources):
            s_id = self.tree.insert("", "end", text=source.path)
            for r_idx, recipe in enumerate(source.recipes):
                r_id = self.tree.insert(s_id, "end", text=recipe.name)
                for sh_idx, sheet in enumerate(recipe.sheets):
                    self.tree.insert(r_id, "end", text=sheet.name)


def main() -> None:
    app = TurboExtractorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
