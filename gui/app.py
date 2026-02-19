\
from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional

from core.project import ProjectConfig, SourceConfig, RecipeConfig
from core.models import SheetConfig, Destination


class TurboExtractorApp(tk.Tk):
    """
    V3 GUI — Tree structure management added.
    Supports:
    - Add Source(s)
    - Add Recipe
    - Add Sheet
    - Remove Selected (with auto-delete empty recipe)
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

        btns = ttk.Frame(left)
        btns.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))

        ttk.Button(btns, text="Add Source(s)...", command=self.add_sources).pack(side="left")
        ttk.Button(btns, text="Add Recipe", command=self.add_recipe).pack(side="left", padx=(6, 0))
        ttk.Button(btns, text="Add Sheet", command=self.add_sheet).pack(side="left", padx=(6, 0))
        ttk.Button(btns, text="Remove Selected", command=self.remove_selected).pack(side="left", padx=(6, 0))

        # ----- RIGHT PANEL -----
        right = ttk.LabelFrame(root, text="Selected Sheet", padding=8)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(1, weight=1)

        ttk.Label(right, text="Columns:").grid(row=0, column=0, sticky="w")
        self.columns_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.columns_var).grid(row=0, column=1, sticky="ew")
        self.columns_var.trace_add("write", self._push_editor_to_sheet)

    # ---------------- Tree Structure ----------------

    def add_sources(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Add source file(s)",
            filetypes=[("Excel/CSV", "*.xlsx *.xlsm *.csv"), ("All files", "*.*")]
        )
        if not paths:
            return

        for p in paths:
            default_sheet = SheetConfig(
                name="Sheet1",
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
            recipe = RecipeConfig(name="Recipe1", sheets=[default_sheet])
            source = SourceConfig(path=p, recipes=[recipe])
            self.project.sources.append(source)

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
        source.recipes.append(RecipeConfig(name=f"Recipe{len(source.recipes)+1}", sheets=[]))
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

        source = self.project.sources[path[0]]
        recipe = source.recipes[path[1]]

        new_sheet = SheetConfig(
            name=f"Sheet{len(recipe.sheets)+1}",
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
        recipe.sheets.append(new_sheet)
        self.refresh_tree()

    def remove_selected(self) -> None:
        sel = self.tree.selection()
        if not sel:
            return

        path = self._get_tree_path(sel[0])

        if len(path) == 1:
            del self.project.sources[path[0]]

        elif len(path) == 2:
            source = self.project.sources[path[0]]
            del source.recipes[path[1]]

        elif len(path) == 3:
            source = self.project.sources[path[0]]
            recipe = source.recipes[path[1]]
            del recipe.sheets[path[2]]

            if not recipe.sheets:
                del source.recipes[path[1]]

        self.current_sheet = None
        self.refresh_tree()

    # ---------------- Tree Helpers ----------------

    def refresh_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        for s_idx, source in enumerate(self.project.sources):
            s_id = self.tree.insert("", "end", text=source.path)
            for r_idx, recipe in enumerate(source.recipes):
                r_id = self.tree.insert(s_id, "end", text=recipe.name)
                for sh_idx, sheet in enumerate(recipe.sheets):
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

    # ---------------- Editor Binding ----------------

    def _on_tree_select(self, event=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        path = self._get_tree_path(sel[0])
        if len(path) == 3:
            self.current_sheet = self.project.sources[path[0]].recipes[path[1]].sheets[path[2]]
        else:
            self.current_sheet = None

    def _push_editor_to_sheet(self, *args) -> None:
        if self.current_sheet:
            self.current_sheet.columns_spec = self.columns_var.get()


def main() -> None:
    app = TurboExtractorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
