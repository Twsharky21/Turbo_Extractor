from __future__ import annotations

import os
import tkinter as tk
from tkinter import messagebox
from typing import Optional

from core.project import SourceConfig


class TreeMixin:
    """
    Mixin for TurboExtractorApp: tree widget operations.

    Covers: refresh_tree, selection/path helpers, right-click context menus,
    inline rename, move up/down for sources/recipes/sheets, reselect helpers.
    """

    # ── Tree display ──────────────────────────────────────────────────────────

    def refresh_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        for source in self.project.sources:
            label = self._source_label(source)
            s_id = self.tree.insert("", "end", text=label)
            self.tree.item(s_id, open=True)
            for recipe in source.recipes:
                r_id = self.tree.insert(s_id, "end", text=recipe.name)
                self.tree.item(r_id, open=True)
                for sheet in recipe.sheets:
                    self.tree.insert(r_id, "end", text=sheet.name)

        self._sync_right_panel_visibility()

    # ── Path helpers ──────────────────────────────────────────────────────────

    def _get_tree_path(self, item_id):
        path = []
        current = item_id
        while current:
            parent = self.tree.parent(current)
            siblings = self.tree.get_children(parent)
            path.insert(0, list(siblings).index(current))
            current = parent
        return path

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

    def _select_source_by_path(self, source_path: str) -> None:
        want = os.path.basename(source_path)
        for item_id in self.tree.get_children(""):
            txt = self.tree.item(item_id, "text")
            if txt == source_path or txt == want:
                self.tree.selection_set(item_id)
                self.tree.see(item_id)
                self._on_tree_select()
                return

    # ── Selection / panel sync ────────────────────────────────────────────────

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

        if len(path) == 1:
            src = self.project.sources[path[0]]
            self.selection_name_var.set(self._source_label(src))
            self.current_sheet = None
            self.current_source_path = None
            self.current_recipe_name = None
            self._sync_right_panel_visibility(is_sheet=False)
            self._clear_editor()
            return

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
        if is_sheet is None:
            sel = self.tree.selection()
            if not sel:
                is_sheet = False
            else:
                is_sheet = (len(self._get_tree_path(sel[0])) == 3)
        try:
            self.deiconify()
        except Exception:
            pass
        if is_sheet:
            try:
                self.selection_box.grid_remove()
                self.sheet_box.grid(row=1, column=0, sticky="ew")
                self.rules_box.grid(row=2, column=0, sticky="nsew")
                self.dest_box.grid(row=3, column=0, sticky="ew")
            except Exception:
                pass
        else:
            try:
                self.selection_box.grid(row=0, column=0, sticky="ew")
                self.sheet_box.grid_remove()
                self.rules_box.grid_remove()
                self.dest_box.grid_remove()
            except Exception:
                pass

    # ── Right-click context menus ─────────────────────────────────────────────

    def _on_tree_right_click(self, event) -> None:
        item = self.tree.identify_row(event.y)
        if not item:
            return
        self.tree.selection_set(item)
        path = self._get_tree_path(item)

        if len(path) == 1:
            self._ctx_source_index = path[0]
            try:
                self._source_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self._source_menu.grab_release()
            return

        if len(path) == 2:
            self._ctx_recipe_path = path
            try:
                self._recipe_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self._recipe_menu.grab_release()
            return

        if len(path) == 3:
            self._ctx_sheet_path = path
            try:
                self._sheet_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self._sheet_menu.grab_release()
            return

    def _get_ctx_source(self) -> Optional[SourceConfig]:
        if self._ctx_source_index is None:
            return None
        if 0 <= self._ctx_source_index < len(self.project.sources):
            return self.project.sources[self._ctx_source_index]
        return None

    # ── Inline rename ─────────────────────────────────────────────────────────

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
        entry = tk.ttk.Entry(self.tree, textvariable=var)
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

    def _apply_recipe_rename(self, path: list[int], new_name: str) -> None:
        s, r = path[0], path[1]
        self.project.sources[s].recipes[r].name = new_name

    def _apply_sheet_rename(self, path: list[int], new_name: str) -> None:
        s, r, sh = path[0], path[1], path[2]
        sheet = self.project.sources[s].recipes[r].sheets[sh]
        sheet.name = new_name
        sheet.workbook_sheet = new_name

    # ── Move up/down ──────────────────────────────────────────────────────────

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

        self.project.sources[idx - 1], self.project.sources[idx] = \
            self.project.sources[idx], self.project.sources[idx - 1]
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

        self.project.sources[idx + 1], self.project.sources[idx] = \
            self.project.sources[idx], self.project.sources[idx + 1]
        moved_path = self.project.sources[idx + 1].path
        self.refresh_tree()
        self._select_source_by_path(moved_path)
        self._mark_dirty()

    def move_selected_up(self) -> None:
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
            self.project.sources[idx - 1], self.project.sources[idx] = \
                self.project.sources[idx], self.project.sources[idx - 1]
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
            self.project.sources[idx + 1], self.project.sources[idx] = \
                self.project.sources[idx], self.project.sources[idx + 1]
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

    # ── Reselect helpers ──────────────────────────────────────────────────────

    def _reselect_after_move(self, key) -> None:
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

    def _reselect_after_remove(self, removed_path: list) -> None:
        depth = len(removed_path)
        removed_idx = removed_path[-1]

        if depth == 1:
            roots = self.tree.get_children("")
            if not roots:
                return
            target_idx = min(removed_idx, len(roots) - 1)
            item = roots[target_idx]

        elif depth == 2:
            roots = self.tree.get_children("")
            if removed_path[0] >= len(roots):
                return
            s_id = roots[removed_path[0]]
            recipes = self.tree.get_children(s_id)
            if recipes:
                target_idx = min(removed_idx, len(recipes) - 1)
                item = recipes[target_idx]
            else:
                item = s_id

        elif depth == 3:
            roots = self.tree.get_children("")
            if removed_path[0] >= len(roots):
                return
            s_id = roots[removed_path[0]]
            recipes = self.tree.get_children(s_id)
            if removed_path[1] >= len(recipes):
                item = s_id
            else:
                r_id = recipes[removed_path[1]]
                sheets = self.tree.get_children(r_id)
                if sheets:
                    target_idx = min(removed_idx, len(sheets) - 1)
                    item = sheets[target_idx]
                else:
                    item = r_id
        else:
            return

        self.tree.selection_set(item)
        self.tree.see(item)
        self._on_tree_select()
