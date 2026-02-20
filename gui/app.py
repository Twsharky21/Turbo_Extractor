from __future__ import annotations

import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional

from gui.ui_build import build_ui

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

        # Editor loading guard: suppresses _push_editor_to_sheet while loading
        self._loading: bool = False

        # Editor loading guard: suppresses _push_editor_to_sheet while loading
        self._loading: bool = False

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
        build_ui(self)

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
        # Only auto-load if TURBO_AUTOSAVE_PATH is explicitly set in the environment.
        # This preserves test isolation: tests that need autoload set the env var via
        # monkeypatch; tests that don't set it start with a clean project.
        # Real-app usage: main() sets os.environ[ENV_AUTOSAVE_PATH] before launching.
        import os
        from core.autosave import ENV_AUTOSAVE_PATH
        if not os.environ.get(ENV_AUTOSAVE_PATH):
            return
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

    def _source_label(self, src: SourceConfig) -> str:
        name = getattr(src, "name", "")
        if isinstance(name, str) and name.strip():
            return name.strip()
        return os.path.basename(src.path)

    def refresh_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        for source in self.project.sources:
            # Display contract: show filename-ish label (not full path). Prefer source.name when present.
            label = self._source_label(source)
            s_id = self.tree.insert("", "end", text=label)
            self.tree.item(s_id, open=True)
            for recipe in source.recipes:
                r_id = self.tree.insert(s_id, "end", text=recipe.name)
                self.tree.item(r_id, open=True)
                for sheet in recipe.sheets:
                    self.tree.insert(r_id, "end", text=sheet.name)

        # Maintain editor visibility parity with selection.
        self._sync_right_panel_visibility()

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
            self.selection_name_var.set(self.current_sheet.name)
            self._sync_right_panel_visibility(is_sheet=True)
            self._load_sheet_into_editor(self.current_sheet)
            return

        # Source selection
        if len(path) == 1:
            src = self.project.sources[path[0]]
            self.selection_name_var.set(self._source_label(src))
            self.current_sheet = None
            self.current_source_path = None
            self.current_recipe_name = None
            self._sync_right_panel_visibility(is_sheet=False)
            self._clear_editor()
            return

        # Recipe selection
        if len(path) == 2:
            recipe = self.project.sources[path[0]].recipes[path[1]]
            self.selection_name_var.set(recipe.name)
            self.current_sheet = None
            self.current_source_path = None
            self.current_recipe_name = None
            self._sync_right_panel_visibility(is_sheet=False)
            self._clear_editor()
            return

        self.current_sheet = None
        self.current_source_path = None
        self.current_recipe_name = None
        self._sync_right_panel_visibility(is_sheet=False)
        self._clear_editor()

    def _sync_right_panel_visibility(self, is_sheet: Optional[bool] = None) -> None:
        """Show full editor only for Sheet selection; otherwise show name-only."""
        if is_sheet is None:
            sel = self.tree.selection()
            if not sel:
                is_sheet = False
            else:
                is_sheet = (len(self._get_tree_path(sel[0])) == 3)
        # selection_box is always visible (never removed).
        # Ensure the top-level window is mapped so winfo_ismapped() returns 1.
        try:
            self.deiconify()
        except Exception:
            pass
        if is_sheet:
            try:
                self.sheet_box.grid(row=1, column=0, sticky="ew")
                self.rules_box.grid(row=2, column=0, sticky="nsew")
                self.dest_box.grid(row=3, column=0, sticky="ew")
            except Exception:
                pass
        else:
            try:
                self.sheet_box.grid_remove()
                self.rules_box.grid_remove()
                self.dest_box.grid_remove()
            except Exception:
                pass
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
            # Ensure stable display name for tree/tests
            try:
                if not getattr(src, "name", ""):
                    src.name = os.path.basename(p)
            except Exception:
                pass
            if default_template:
                tpl.apply_template_to_source(src, default_template)
            else:
                sheet = self._make_default_sheet(name="sheet1")
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


    def move_selected_up(self) -> None:
        # Move selected node up within its parent (Source / Recipe / Sheet).
        # Use tree.focus() as fallback: btn.invoke() can clear selection in headless Tk.
        sel = self.tree.selection()
        if not sel:
            focused = self.tree.focus()
            if not focused:
                return
            sel = (focused,)

        path = self._get_tree_path(sel[0])
        if len(path) == 1:
            idx = path[0]
            if idx <= 0:
                return
            self.project.sources[idx - 1], self.project.sources[idx] = self.project.sources[idx], self.project.sources[idx - 1]
            moved_key = ("source", self.project.sources[idx - 1].path)

        elif len(path) == 2:
            s, r = path
            if r <= 0:
                return
            recipes = self.project.sources[s].recipes
            recipes[r - 1], recipes[r] = recipes[r], recipes[r - 1]
            moved_key = ("recipe", s, recipes[r - 1].name)

        elif len(path) == 3:
            s, r, sh = path
            if sh <= 0:
                return
            sheets = self.project.sources[s].recipes[r].sheets
            sheets[sh - 1], sheets[sh] = sheets[sh], sheets[sh - 1]
            moved_key = ("sheet", s, r, sheets[sh - 1].name)

        else:
            return

        self.refresh_tree()
        self._reselect_after_move(moved_key)
        self._mark_dirty()

    def move_selected_down(self) -> None:
        # Move selected node down within its parent (Source / Recipe / Sheet).
        # Use tree.focus() as fallback: btn.invoke() can clear selection in headless Tk.
        sel = self.tree.selection()
        if not sel:
            focused = self.tree.focus()
            if not focused:
                return
            sel = (focused,)

        path = self._get_tree_path(sel[0])
        if len(path) == 1:
            idx = path[0]
            if idx >= len(self.project.sources) - 1:
                return
            self.project.sources[idx + 1], self.project.sources[idx] = self.project.sources[idx], self.project.sources[idx + 1]
            moved_key = ("source", self.project.sources[idx + 1].path)

        elif len(path) == 2:
            s, r = path
            recipes = self.project.sources[s].recipes
            if r >= len(recipes) - 1:
                return
            recipes[r + 1], recipes[r] = recipes[r], recipes[r + 1]
            moved_key = ("recipe", s, recipes[r + 1].name)

        elif len(path) == 3:
            s, r, sh = path
            sheets = self.project.sources[s].recipes[r].sheets
            if sh >= len(sheets) - 1:
                return
            sheets[sh + 1], sheets[sh] = sheets[sh], sheets[sh + 1]
            moved_key = ("sheet", s, r, sheets[sh + 1].name)

        else:
            return

        self.refresh_tree()
        self._reselect_after_move(moved_key)
        self._mark_dirty()

    def _reselect_after_move(self, key) -> None:
        # Reselect the moved item after tree refresh.
        if not key:
            return

        kind = key[0]
        if kind == "source":
            _, src_path = key
            want = os.path.basename(src_path)
            for s_id in self.tree.get_children(""):
                txt = self.tree.item(s_id, "text")
                if txt == src_path or txt == want:
                    self.tree.selection_set(s_id)
                    self.tree.see(s_id)
                    self._on_tree_select()
                    return

        if kind == "recipe":
            _, s, recipe_name = key
            s_children = self.tree.get_children("")
            if s >= len(s_children):
                return
            s_id = s_children[s]
            for r_id in self.tree.get_children(s_id):
                if self.tree.item(r_id, "text") == recipe_name:
                    self.tree.selection_set(r_id)
                    self.tree.see(r_id)
                    self._on_tree_select()
                    return

        if kind == "sheet":
            _, s, r, sheet_name = key
            s_children = self.tree.get_children("")
            if s >= len(s_children):
                return
            s_id = s_children[s]
            r_children = self.tree.get_children(s_id)
            if r >= len(r_children):
                return
            r_id = r_children[r]
            for sh_id in self.tree.get_children(r_id):
                if self.tree.item(sh_id, "text") == sheet_name:
                    self.tree.selection_set(sh_id)
                    self.tree.see(sh_id)
                    self._on_tree_select()
                    return

    def _select_source_by_path(self, source_path: str) -> None:
        want = os.path.basename(source_path)
        for item_id in self.tree.get_children(""):
            txt = self.tree.item(item_id, "text")
            if txt == source_path or txt == want:
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
        if len(path) not in (1, 2, 3):
            messagebox.showwarning("Select Source", "Select a Source to add a Recipe.")
            return

        source = self.project.sources[path[0]]
        new_recipe = RecipeConfig(name=f"Recipe{len(source.recipes) + 1}", sheets=[])
        source.recipes.append(new_recipe)
        # Insert into tree incrementally so pre-captured item IDs remain valid.
        src_children = self.tree.get_children("")
        s_id = src_children[path[0]]
        r_id = self.tree.insert(s_id, "end", text=new_recipe.name)
        self.tree.item(s_id, open=True)
        self._mark_dirty()

    def add_sheet(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Select Recipe", "Select a Recipe to add a Sheet.")
            return

        path = self._get_tree_path(sel[0])
        if len(path) not in (1, 2, 3):
            messagebox.showwarning("Select Recipe", "Select a Source/Recipe/Sheet to add a Sheet.")
            return

        source = self.project.sources[path[0]]

        # Source selected: add under first recipe (create if missing)
        if len(path) == 1:
            if not source.recipes:
                source.recipes.append(RecipeConfig(name="Recipe1", sheets=[]))
            recipe = source.recipes[0]
        else:
            # Recipe selected or Sheet selected -> parent recipe
            recipe = source.recipes[path[1]]

        # Name contract: all new sheets are "sheet1" and duplicates are allowed.
        new_sheet = self._make_default_sheet(name="sheet1")
        recipe.sheets.append(new_sheet)
        # Insert into tree incrementally so pre-captured item IDs remain valid.
        src_children = self.tree.get_children("")
        s_id = src_children[path[0]]
        recipe_idx = path[1] if len(path) >= 2 else 0
        r_children = self.tree.get_children(s_id)
        r_id = r_children[recipe_idx]
        self.tree.insert(r_id, "end", text=new_sheet.name)
        self.tree.item(r_id, open=True)
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

    # ---------------- Editor binding ----------------

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
        # Suppressed while _load_sheet_into_editor is running.
        if self._loading:
            return
        # Always write into the currently selected sheet (guards against stale current_sheet pointers
        # during headless tests / event-order differences).
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
                # Backwards compatibility (older UI stored raw values)
                sheet.paste_mode = val
        if self.combine_var.get():
            sheet.rules_combine = self.combine_var.get()

        sheet.destination.file_path = self.dest_file_var.get()
        sheet.destination.sheet_name = self.dest_sheet_var.get()
        sheet.destination.start_col = self.start_col_var.get()
        sheet.destination.start_row = self.start_row_var.get()

        self._mark_dirty()

    # ---------------- Rules UI ----------------

    def add_rule(self) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        path = self._get_tree_path(sel[0])
        if len(path) != 3:
            return
        sheet = self.project.sources[path[0]].recipes[path[1]].sheets[path[2]]
        sheet.rules.append(Rule(mode="include", column="A", operator="equals", value=""))
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
    # Ensure autosave path is set so _try_load_autosave activates in real usage.
    import os
    from core.autosave import ENV_AUTOSAVE_PATH, resolve_autosave_path
    if not os.environ.get(ENV_AUTOSAVE_PATH):
        os.environ[ENV_AUTOSAVE_PATH] = resolve_autosave_path()
    app = TurboExtractorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
