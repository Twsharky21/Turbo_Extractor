from __future__ import annotations

import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional

from core.project import ProjectConfig, SourceConfig, RecipeConfig
from core.models import SheetConfig, Destination, Rule
from core import templates as tpl
from core.engine import run_all as engine_run_all, run_sheet as engine_run_sheet
from core.errors import AppError
from core.autosave import resolve_autosave_path, save_project_atomic, load_project_if_exists


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
        self.current_source_path: Optional[str] = None
        self.current_recipe_name: Optional[str] = None

        # Inline rename state
        self._rename_entry: Optional[ttk.Entry] = None
        self._rename_item_id: Optional[str] = None
        self._rename_path: Optional[list[int]] = None
        self._rename_kind: Optional[str] = None  # 'recipe' | 'sheet'

        # Autosave state
        self._autosave_dirty: bool = False
        self._autosave_after_id: Optional[str] = None
        self._autosave_periodic_id: Optional[str] = None
        self._autosave_path: str = resolve_autosave_path()

        self._build_ui()

        # Load autosave (if present) AFTER UI exists, then refresh.
        self._try_load_autosave()

        # Periodic safety save
        self._schedule_periodic_autosave()

        # Ensure save on close
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------------- UI ----------------

    def _build_ui(self) -> None:
            # Overall layout: top toolbar, then left tree + right editor (monolith-like)
            self.columnconfigure(0, weight=1)
            self.rowconfigure(0, weight=1)

            root = ttk.Frame(self, padding=8)
            root.grid(row=0, column=0, sticky="nsew")
            root.columnconfigure(0, weight=1)
            root.columnconfigure(1, weight=3)
            root.rowconfigure(1, weight=1)

            # ----- TOP TOOLBAR -----
            topbar = ttk.Frame(root)
            topbar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
            for i in range(8):
                topbar.columnconfigure(i, weight=0)
            topbar.columnconfigure(7, weight=1)

            ttk.Button(topbar, text="Add Source (XLSX/CSV)", style="RunAccent.TButton", command=self.add_sources).grid(row=0, column=0, padx=(0, 6))
            ttk.Button(topbar, text="Remove Selected", command=self.remove_selected).grid(row=0, column=5, padx=(0, 6))
            ttk.Button(topbar, text="Move Source Up", command=self.move_source_up).grid(row=0, column=1, padx=(0, 6))
            ttk.Button(topbar, text="Move Source Down", command=self.move_source_down).grid(row=0, column=2, padx=(0, 6))
            ttk.Button(topbar, text="Add Recipe", command=self.add_recipe).grid(row=0, column=3, padx=(0, 6))
            ttk.Button(topbar, text="Add Sheet", command=self.add_sheet).grid(row=0, column=4, padx=(0, 6))

            # ----- LEFT: TREE -----
            left = ttk.Frame(root)
            left.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
            left.columnconfigure(0, weight=1)
            left.rowconfigure(0, weight=1)

            self.tree = ttk.Treeview(left, show="tree", selectmode="browse")
            self.tree.grid(row=0, column=0, sticky="nsew")
            self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
            self.tree.bind("<Button-3>", self._on_tree_right_click)

            yscroll = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
            yscroll.grid(row=0, column=1, sticky="ns")
            self.tree.configure(yscrollcommand=yscroll.set)

            # ----- RIGHT: EDITOR -----
            right = ttk.Frame(root)
            right.grid(row=1, column=1, sticky="nsew")
            right.columnconfigure(0, weight=1)
            right.rowconfigure(1, weight=1)

            # Selected Sheet (within Recipe)
            sheet_box = ttk.LabelFrame(right, text="Selected Sheet (within Recipe)", padding=10)
            sheet_box.grid(row=0, column=0, sticky="ew")
            sheet_box.columnconfigure(1, weight=1)

            ttk.Label(sheet_box, text="Columns (e.g., A,C,AC-ZZ):").grid(row=0, column=0, sticky="w")
            self.columns_var = tk.StringVar()
            ttk.Entry(sheet_box, textvariable=self.columns_var).grid(row=0, column=1, sticky="ew", padx=(10, 0))
            self.columns_var.trace_add("write", self._push_editor_to_sheet)

            ttk.Label(sheet_box, text="Rows (e.g., 1-3,9-80,117):").grid(row=1, column=0, sticky="w", pady=(6, 0))
            self.rows_var = tk.StringVar()
            ttk.Entry(sheet_box, textvariable=self.rows_var).grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=(6, 0))
            self.rows_var.trace_add("write", self._push_editor_to_sheet)

            ttk.Label(sheet_box, text="Source Start Row:").grid(row=2, column=0, sticky="w", pady=(6, 0))
            self.source_start_row_var = tk.StringVar()
            ttk.Entry(sheet_box, textvariable=self.source_start_row_var, width=10).grid(row=2, column=1, sticky="w", padx=(10, 0), pady=(6, 0))
            self.source_start_row_var.trace_add("write", self._push_editor_to_sheet)

            ttk.Label(sheet_box, text="Column paste mode:").grid(row=3, column=0, sticky="w", pady=(6, 0))
            self.paste_var = tk.StringVar()
            self.paste_combo = ttk.Combobox(
                sheet_box,
                textvariable=self.paste_var,
                values=["Pack Together", "Keep Format"],
                state="readonly",
                width=18,
            )
            self.paste_combo.grid(row=3, column=1, sticky="w", padx=(10, 0), pady=(6, 0))
            self.paste_combo.bind("<<ComboboxSelected>>", self._push_editor_to_sheet)

            # ----- RULES -----
            rules_box = ttk.LabelFrame(right, text="Rules", padding=10)
            rules_box.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
            rules_box.columnconfigure(0, weight=1)
            rules_box.rowconfigure(2, weight=1)
            right.rowconfigure(1, weight=1)

            top_rules = ttk.Frame(rules_box)
            top_rules.grid(row=0, column=0, sticky="ew")
            top_rules.columnconfigure(1, weight=1)

            ttk.Label(top_rules, text="Rules combine:").grid(row=0, column=0, sticky="w")
            self.combine_var = tk.StringVar()
            self.combine_combo = ttk.Combobox(top_rules, textvariable=self.combine_var, values=["AND", "OR"], state="readonly", width=8)
            self.combine_combo.grid(row=0, column=1, sticky="w", padx=(10, 0))
            self.combine_combo.bind("<<ComboboxSelected>>", self._push_editor_to_sheet)

            ttk.Button(top_rules, text="+ Add rule", command=self.add_rule).grid(row=0, column=2, sticky="w", padx=(20, 0))

            # Header row
            hdr = ttk.Frame(rules_box)
            hdr.grid(row=1, column=0, sticky="ew", pady=(8, 4))
            ttk.Label(hdr, text="Include/Exclude").grid(row=0, column=0, sticky="w", padx=(0, 10))
            ttk.Label(hdr, text="Column").grid(row=0, column=1, sticky="w", padx=(0, 10))
            ttk.Label(hdr, text="Operator").grid(row=0, column=2, sticky="w", padx=(0, 10))
            ttk.Label(hdr, text="Value").grid(row=0, column=3, sticky="w", padx=(0, 10))

            # Scrollable rules area
            rules_area = ttk.Frame(rules_box)
            rules_area.grid(row=2, column=0, sticky="nsew")
            rules_area.columnconfigure(0, weight=1)
            rules_area.rowconfigure(0, weight=1)

            self.rules_canvas = tk.Canvas(rules_area, height=260)
            self.rules_canvas.grid(row=0, column=0, sticky="nsew")

            rules_scroll = ttk.Scrollbar(rules_area, orient="vertical", command=self.rules_canvas.yview)
            rules_scroll.grid(row=0, column=1, sticky="ns")
            self.rules_canvas.configure(yscrollcommand=rules_scroll.set)

            self.rules_frame = ttk.Frame(self.rules_canvas)
            self.rules_canvas.create_window((0, 0), window=self.rules_frame, anchor="nw")

            self.rules_frame.bind(
                "<Configure>",
                lambda e: self.rules_canvas.configure(scrollregion=self.rules_canvas.bbox("all")),
            )

            # ----- DESTINATION -----
            dest_box = ttk.LabelFrame(right, text="Destination", padding=10)
            dest_box.grid(row=2, column=0, sticky="ew", pady=(10, 0))
            dest_box.columnconfigure(1, weight=1)

            ttk.Label(dest_box, text="File:").grid(row=0, column=0, sticky="w")
            self.dest_file_var = tk.StringVar()
            ttk.Entry(dest_box, textvariable=self.dest_file_var).grid(row=0, column=1, sticky="ew", padx=(10, 10))
            ttk.Button(dest_box, text="Browse", command=self.browse_destination).grid(row=0, column=2, sticky="ew")
            self.dest_file_var.trace_add("write", self._push_editor_to_sheet)

            ttk.Label(dest_box, text="Sheet name:").grid(row=1, column=0, sticky="w", pady=(6, 0))
            self.dest_sheet_var = tk.StringVar()
            ttk.Entry(dest_box, textvariable=self.dest_sheet_var).grid(row=1, column=1, columnspan=2, sticky="ew", padx=(10, 0), pady=(6, 0))
            self.dest_sheet_var.trace_add("write", self._push_editor_to_sheet)

            ttk.Label(dest_box, text="Start column (e.g., A, D, AA):").grid(row=2, column=0, sticky="w", pady=(6, 0))
            start_row = ttk.Frame(dest_box)
            start_row.grid(row=2, column=1, columnspan=2, sticky="w", padx=(10, 0), pady=(6, 0))

            self.start_col_var = tk.StringVar()
            ttk.Entry(start_row, textvariable=self.start_col_var, width=8).grid(row=0, column=0, sticky="w")
            self.start_col_var.trace_add("write", self._push_editor_to_sheet)

            ttk.Label(start_row, text="Start row:").grid(row=0, column=1, sticky="w", padx=(15, 6))
            self.start_row_var = tk.StringVar()
            ttk.Entry(start_row, textvariable=self.start_row_var, width=10).grid(row=0, column=2, sticky="w")
            self.start_row_var.trace_add("write", self._push_editor_to_sheet)

            # ----- BOTTOM: STATUS + RUN BUTTONS -----
            bottom = ttk.Frame(right)
            bottom.grid(row=3, column=0, sticky="ew", pady=(10, 0))
            bottom.columnconfigure(0, weight=1)

            self.status_var = tk.StringVar(value="Idle")
            ttk.Label(bottom, textvariable=self.status_var).grid(row=0, column=0, sticky="w")

            # Make RUN buttons visually closer to monolith (blue accent).
            try:
                style = ttk.Style()
                if style.theme_use() != "clam":
                    style.theme_use("clam")
                style.configure("RunAccent.TButton", padding=(16, 6))
                style.map(
                    "RunAccent.TButton",
                    background=[("active", "#1e6bd6"), ("!disabled", "#1f76ff")],
                    foreground=[("!disabled", "white")],
                )
            except Exception:
                # If theme/style cannot be applied, fall back to default styling.
                pass

            run_btns = ttk.Frame(bottom)
            run_btns.grid(row=0, column=1, sticky="e")

            ttk.Button(run_btns, text="RUN", style="RunAccent.TButton", command=self.run_selected_sheet).pack(side="right", padx=(6, 0))
            ttk.Button(run_btns, text="RUN ALL", style="RunAccent.TButton", command=self.run_all).pack(side="right")

            # Context menu (Source only)
            self._source_menu = tk.Menu(self, tearoff=0)
            self._source_menu.add_command(label="Save Template...", command=self._ctx_save_template)
            self._source_menu.add_command(label="Load Template...", command=self._ctx_load_template)
            self._source_menu.add_separator()
            self._source_menu.add_command(label="Set Default", command=self._ctx_set_default)
            self._source_menu.add_command(label="Reset Default", command=self._ctx_reset_default)
            self._ctx_source_index: Optional[int] = None

            # Context menu (Recipe)
            self._recipe_menu = tk.Menu(self, tearoff=0)
            self._recipe_menu.add_command(label="Rename Recipe", command=self._ctx_rename_recipe)
            self._ctx_recipe_path: Optional[list[int]] = None

            # Context menu (Sheet)
            self._sheet_menu = tk.Menu(self, tearoff=0)
            self._sheet_menu.add_command(label="Rename Sheet", command=self._ctx_rename_sheet)
            self._ctx_sheet_path: Optional[list[int]] = None
    def _mark_dirty(self) -> None:
        self._autosave_dirty = True
        self._schedule_debounced_autosave()

    def _schedule_debounced_autosave(self) -> None:
        if self._autosave_after_id is not None:
            try:
                self.after_cancel(self._autosave_after_id)
            except Exception:
                pass
        # Debounce ~1.2s
        self._autosave_after_id = self.after(1200, self._autosave_now)

    def _schedule_periodic_autosave(self) -> None:
        # ~45s safety save
        self._autosave_periodic_id = self.after(45000, self._periodic_autosave_tick)

    def _periodic_autosave_tick(self) -> None:
        if self._autosave_dirty:
            self._autosave_now()
        self._schedule_periodic_autosave()

    def _autosave_now(self) -> None:
        if not self._autosave_dirty:
            return
        try:
            save_project_atomic(self.project, self._autosave_path)
            self._autosave_dirty = False
        except Exception:
            # Autosave should never crash the app.
            pass

    def _try_load_autosave(self) -> None:
        try:
            loaded = load_project_if_exists(self._autosave_path)
            if loaded is not None:
                self.project = loaded
                self.refresh_tree()
                self._clear_editor()
        except Exception:
            pass

    def _on_close(self) -> None:
        try:
            self._autosave_now()
        finally:
            self.destroy()

    # ---------------- Tree helpers ----------------

    def refresh_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        for source in self.project.sources:
            s_id = self.tree.insert("", "end", text=source.path)
            self.tree.item(s_id, open=True)
            for recipe in source.recipes:
                r_id = self.tree.insert(s_id, "end", text=recipe.name)
                self.tree.item(r_id, open=True)
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
            self.current_source_path = self.project.sources[path[0]].path
            self.current_recipe_name = self.project.sources[path[0]].recipes[path[1]].name
            self._load_sheet_into_editor(self.current_sheet)
        else:
            self.current_sheet = None
            self.current_source_path = None
            self.current_recipe_name = None
            self._clear_editor()

    def _on_tree_right_click(self, event) -> None:
        item = self.tree.identify_row(event.y)
        if not item:
            return
        self.tree.selection_set(item)
        path = self._get_tree_path(item)

        # Source
        if len(path) == 1:
            self._ctx_source_index = path[0]
            try:
                self._source_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self._source_menu.grab_release()
            return

        # Recipe
        if len(path) == 2:
            self._ctx_recipe_path = path
            try:
                self._recipe_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self._recipe_menu.grab_release()
            return

        # Sheet
        if len(path) == 3:
            self._ctx_sheet_path = path
            try:
                self._sheet_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self._sheet_menu.grab_release()
            return

    # ---------------- Templates / Default (Source context menu) ----------------

    def _get_ctx_source(self) -> Optional[SourceConfig]:
        if self._ctx_source_index is None:
            return None
        if 0 <= self._ctx_source_index < len(self.project.sources):
            return self.project.sources[self._ctx_source_index]
        return None

    def _ctx_save_template(self) -> None:
        src = self._get_ctx_source()
        if not src:
            return
        path = filedialog.asksaveasfilename(
            title="Save Source Template",
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        tpl.save_template_json(tpl.source_to_template(src), path)

    def _ctx_load_template(self) -> None:
        src = self._get_ctx_source()
        if not src:
            return
        path = filedialog.askopenfilename(
            title="Load Source Template",
            filetypes=[("JSON", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        template = tpl.load_template_json(path)
        tpl.apply_template_to_source(src, template)
        self.refresh_tree()
        self._mark_dirty()

    def _ctx_set_default(self) -> None:
        src = self._get_ctx_source()
        if not src:
            return
        tpl.set_default_template(tpl.source_to_template(src))

    def _ctx_reset_default(self) -> None:
        tpl.reset_default_template()


    # ---------------- Inline Rename (Recipe/Sheet) ----------------

    def _ctx_rename_recipe(self) -> None:
        if not self._ctx_recipe_path:
            return
        self._begin_inline_rename(kind="recipe", path=self._ctx_recipe_path)

    def _ctx_rename_sheet(self) -> None:
        if not self._ctx_sheet_path:
            return
        self._begin_inline_rename(kind="sheet", path=self._ctx_sheet_path)

    def _begin_inline_rename(self, kind: str, path: list[int]) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        item_id = sel[0]

        self._cancel_inline_rename()

        bbox = self.tree.bbox(item_id)
        if not bbox:
            return
        x, y, w, h = bbox

        current_text = self.tree.item(item_id, "text")

        var = tk.StringVar(value=current_text)
        entry = ttk.Entry(self.tree, textvariable=var)
        entry.place(x=x, y=y, width=max(w, 80), height=h)

        self._rename_entry = entry
        self._rename_item_id = item_id
        self._rename_path = list(path)
        self._rename_kind = kind

        entry.focus_set()
        entry.icursor("end")

        entry.bind("<Return>", lambda e: self._commit_inline_rename())
        entry.bind("<Escape>", lambda e: self._cancel_inline_rename())
        entry.bind("<FocusOut>", lambda e: self._commit_inline_rename())

    def _cancel_inline_rename(self) -> None:
        if self._rename_entry is not None:
            try:
                self._rename_entry.destroy()
            except Exception:
                pass
        self._rename_entry = None
        self._rename_item_id = None
        self._rename_path = None
        self._rename_kind = None

    def _commit_inline_rename(self) -> None:
        if self._rename_entry is None or self._rename_path is None or self._rename_kind is None:
            return
        new_name = self._rename_entry.get().strip()
        kind = self._rename_kind
        path = self._rename_path

        self._cancel_inline_rename()

        if not new_name:
            return

        if kind == "recipe" and len(path) == 2:
            self._apply_recipe_rename(path, new_name)
            self.refresh_tree()
            self._select_tree_by_indices(path)
            self._mark_dirty()
            return

        if kind == "sheet" and len(path) == 3:
            self._apply_sheet_rename(path, new_name)
            self.refresh_tree()
            self._select_tree_by_indices(path)
            self._mark_dirty()
            return

    def _select_tree_by_indices(self, path: list[int]) -> None:
        if not path:
            return
        roots = list(self.tree.get_children(""))
        if path[0] < 0 or path[0] >= len(roots):
            return
        item = roots[path[0]]
        if len(path) >= 2:
            kids = list(self.tree.get_children(item))
            if path[1] < 0 or path[1] >= len(kids):
                return
            item = kids[path[1]]
        if len(path) >= 3:
            kids = list(self.tree.get_children(item))
            if path[2] < 0 or path[2] >= len(kids):
                return
            item = kids[path[2]]

        self.tree.selection_set(item)
        self.tree.see(item)
        self._on_tree_select()

    def _apply_recipe_rename(self, path: list[int], new_name: str) -> None:
        s, r = path[0], path[1]
        self.project.sources[s].recipes[r].name = new_name

    def _apply_sheet_rename(self, path: list[int], new_name: str) -> None:
        s, r, sh = path[0], path[1], path[2]
        sheet = self.project.sources[s].recipes[r].sheets[sh]
        sheet.name = new_name
        sheet.workbook_sheet = new_name
    # ---------------- Structure actions ----------------

    
    def browse_destination(self) -> None:
        # Choose destination XLSX (create or select). GUI only.
        path = filedialog.asksaveasfilename(
            title="Select destination XLSX",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not path:
            return
        self.dest_file_var.set(path)

    def add_sources(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Add source file(s)",
            filetypes=[("Excel/CSV", "*.xlsx *.xlsm *.csv"), ("All files", "*.*")],
        )
        if not paths:
            return

        default_template = tpl.load_default_template()

        for p in paths:
            src = SourceConfig(path=p, recipes=[])
            if default_template:
                tpl.apply_template_to_source(src, default_template)
            else:
                sheet = self._make_default_sheet(name="Sheet1")
                recipe = RecipeConfig(name="Recipe1", sheets=[sheet])
                src.recipes = [recipe]
            self.project.sources.append(src)

        self.refresh_tree()
        self._mark_dirty()

    def move_source_up(self) -> None:
        sel = self.tree.selection()
        if not sel:
            return

        path = self._get_tree_path(sel[0])
        if len(path) != 1:
            messagebox.showinfo("Move Source", "Please select a Source (top-level) to move.")
            return

        idx = path[0]
        if idx <= 0:
            return

        self.project.sources[idx - 1], self.project.sources[idx] = self.project.sources[idx], self.project.sources[idx - 1]
        moved_path = self.project.sources[idx - 1].path
        self.refresh_tree()
        self._select_source_by_path(moved_path)
        self._mark_dirty()

    def move_source_down(self) -> None:
        sel = self.tree.selection()
        if not sel:
            return

        path = self._get_tree_path(sel[0])
        if len(path) != 1:
            messagebox.showinfo("Move Source", "Please select a Source (top-level) to move.")
            return

        idx = path[0]
        if idx >= len(self.project.sources) - 1:
            return

        self.project.sources[idx + 1], self.project.sources[idx] = self.project.sources[idx], self.project.sources[idx + 1]
        moved_path = self.project.sources[idx + 1].path
        self.refresh_tree()
        self._select_source_by_path(moved_path)
        self._mark_dirty()

    def _select_source_by_path(self, source_path: str) -> None:
        for item_id in self.tree.get_children(""):
            if self.tree.item(item_id, "text") == source_path:
                self.tree.selection_set(item_id)
                self.tree.see(item_id)
                self._on_tree_select()
                return

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
        self._mark_dirty()

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
        self._mark_dirty()

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
        self.current_source_path = None
        self.current_recipe_name = None
        self.refresh_tree()
        self._clear_editor()
        self._mark_dirty()

    # ---------------- Run wiring ----------------


    def _feedback_clear(self) -> None:
        # GUI-only feedback. In the monolith-like layout we primarily use the popup summary.
        # Keep these helpers as no-ops / minimal status updates for backward compatibility.
        if hasattr(self, "status_var"):
            self.status_var.set("Running...")

    def _feedback_key(self, source_path: str, recipe_name: str, sheet_name: str) -> str:
        base = os.path.basename(source_path)
        return f"{base} | {recipe_name} / {sheet_name}"

    def _feedback_set_row(self, key: str, status: str, rows: str, message: str) -> None:
        # If a feedback tree is present (older layout), update it. Otherwise ignore.
        tree = getattr(self, "feedback_tree", None)
        if tree is None:
            return
        # Find existing row by key
        existing = None
        for item in tree.get_children():
            if tree.item(item, "text") == key:
                existing = item
                break
        if existing is None:
            existing = tree.insert("", "end", text=key, values=(status, rows, message))
        else:
            tree.item(existing, values=(status, rows, message))

    def _feedback_progress_callback(self, event, payload=None, *args) -> None:
        """Progress hook used by RUN ALL + local RUN, tolerant of legacy call shapes."""
        # Supported call forms:
        # 1) callback(progress_item)
        # 2) callback(event, payload)
        try:
            progress_item = event if payload is None and hasattr(event, 'source_path') else payload
            if progress_item is None:
                return
            key = self._feedback_key(progress_item.source_path, progress_item.recipe_name, progress_item.sheet_name)
            status = getattr(progress_item, 'status', None) or getattr(progress_item, 'message', '') or ''
            rows_written = getattr(progress_item, 'rows_written', None)
            if rows_written is None:
                rows_written = getattr(progress_item, 'rows_written', None)
            rows = '' if getattr(progress_item, 'rows_written', None) is None else str(getattr(progress_item, 'rows_written'))
            msg = getattr(progress_item, 'message', '') or ''
            self._feedback_set_row(key, str(status), rows, msg)
        except Exception:
            return
    def _format_run_report(self, report) -> str:
        lines = []
        for r in report.results:
            label = f"{r.recipe_name} / {r.sheet_name}"
            if getattr(r, "error_code", None):
                lines.append(f"{label}: ERROR {r.error_code} - {r.error_message}")
            else:
                lines.append(f"{label}: {r.rows_written} rows")
        return "\n".join(lines) if lines else "No work items."

    def run_all(self) -> None:
        items = self.project.build_run_items()
        self._feedback_clear()
        # Backwards-compatible call:
        # - Newer core.engine.run_all may accept an on_progress callback.
        # - Tests (and older implementations) may monkeypatch run_all with a
        #   signature that does NOT accept extra kwargs.
        try:
            report = engine_run_all(items, on_progress=self._feedback_progress_callback)
        except TypeError:
            report = engine_run_all(items)
            # Populate feedback panel from the final report (best-effort).
            try:
                for r in getattr(report, "results", []) or []:
                    self._feedback_progress_callback("result", r)
            except Exception:
                pass
        title = "Run complete" if report.ok else "Run complete (with errors)"
        self._show_scrollable_report_dialog(title, self._format_run_report(report))

    def run_selected_sheet(self) -> None:
        if not self.current_sheet or not self.current_source_path or not self.current_recipe_name:
            messagebox.showwarning("Select Sheet", "Select a Sheet to run.")
            return
        self._feedback_clear()
        try:
            res = engine_run_sheet(self.current_source_path, self.current_sheet, recipe_name=self.current_recipe_name)
            self._feedback_progress_callback(
                "result",
                res,
            )
            messagebox.showinfo("Run complete", f"{res.recipe_name} / {res.sheet_name}: {res.rows_written} rows")
        except AppError as e:
            from core.models import SheetResult

            err_res = SheetResult(
                source_path=self.current_source_path,
                recipe_name=self.current_recipe_name,
                sheet_name=self.current_sheet.name,
                dest_file=self.current_sheet.destination.file_path,
                dest_sheet=self.current_sheet.destination.sheet_name,
                rows_written=0,
                message="ERROR",
                error_code=e.code,
                error_message=e.message,
                error_details=e.details,
            )
            self._feedback_progress_callback("error", err_res)
            messagebox.showerror("Run failed", f"{e.code}: {e.message}")

    # -----------------------------
    # Scrollable report dialog
    # -----------------------------

    def _show_scrollable_report_dialog(self, title: str, text: str) -> None:
        """Show a single centered, scrollable summary dialog.

        The RUN ALL summary can be long; messagebox cannot scroll.
        """
        # Ensure only one report window exists at a time.
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
        btn_row.grid(row=1, column=0, columnspan=2, sticky="e", pady=(10, 0))
        close_btn = ttk.Button(btn_row, text="Close", command=win.destroy)
        close_btn.grid(row=0, column=0, sticky="e")

        txt.insert("1.0", text or "")
        txt.configure(state="disabled")

        # Center on screen
        win.update_idletasks()
        w = win.winfo_width()
        h = win.winfo_height()
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        x = max(0, int((sw - w) / 2))
        y = max(0, int((sh - h) / 2))
        win.geometry(f"{w}x{h}+{x}+{y}")

        try:
            close_btn.focus_set()
        except Exception:
            pass

    def _make_default_sheet(self, name: str) -> SheetConfig:
        return SheetConfig(
            name=name,
            workbook_sheet="Sheet1",
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

    # ---------------- Editor binding ----------------

    def _load_sheet_into_editor(self, sheet: SheetConfig) -> None:
        self.columns_var.set(sheet.columns_spec)
        self.rows_var.set(sheet.rows_spec)
        self.source_start_row_var.set(getattr(sheet, "source_start_row", ""))
        self.paste_var.set("Pack Together" if sheet.paste_mode == "pack" else "Keep Format" if sheet.paste_mode == "keep" else sheet.paste_mode)
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
        if not self.current_sheet:
            return

        self.current_sheet.columns_spec = self.columns_var.get()
        self.current_sheet.rows_spec = self.rows_var.get()
        self.current_sheet.source_start_row = self.source_start_row_var.get()
        val = self.paste_var.get().strip()
        if val:
            if val.lower().startswith("pack"):
                self.current_sheet.paste_mode = "pack"
            elif val.lower().startswith("keep"):
                self.current_sheet.paste_mode = "keep"
            else:
                # Backwards compatibility (older UI stored raw values)
                self.current_sheet.paste_mode = val
        if self.combine_var.get():
            self.current_sheet.rules_combine = self.combine_var.get()

        self.current_sheet.destination.file_path = self.dest_file_var.get()
        self.current_sheet.destination.sheet_name = self.dest_sheet_var.get()
        self.current_sheet.destination.start_col = self.start_col_var.get()
        self.current_sheet.destination.start_row = self.start_row_var.get()

        self._mark_dirty()

    # ---------------- Rules UI ----------------

    def add_rule(self) -> None:
        if not self.current_sheet:
            return
        self.current_sheet.rules.append(Rule(mode="include", column="A", operator="equals", value=""))
        self._rebuild_rules()
        self._mark_dirty()

    def _remove_rule(self, idx: int) -> None:
        if not self.current_sheet:
            return
        del self.current_sheet.rules[idx]
        self._rebuild_rules()
        self._mark_dirty()

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
            self._mark_dirty()

        mode_var.trace_add("write", push)
        col_var.trace_add("write", push)
        op_var.trace_add("write", push)
        val_var.trace_add("write", push)


def main() -> None:
    app = TurboExtractorApp()
    app.mainloop()


if __name__ == "__main__":
    main()