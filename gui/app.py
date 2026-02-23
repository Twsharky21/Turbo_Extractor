from __future__ import annotations

import os
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional

from gui.ui_build import build_ui
from gui.mixins import ReportMixin, TreeMixin, EditorMixin, ThrobberMixin

from core.project import ProjectConfig, SourceConfig, RecipeConfig
from core.models import SheetConfig, Destination, Rule
from core import templates as tpl
from core.engine import run_all as engine_run_all, run_sheet as engine_run_sheet
from core.errors import AppError, friendly_message
from core.autosave import resolve_autosave_path, save_project_atomic, load_project_if_exists


class TurboExtractorApp(ReportMixin, TreeMixin, EditorMixin, ThrobberMixin, tk.Tk):
    """
    V3 GUI (merged): Tree structure + minimal sheet editor + rules UI.

    Logic is split across focused mixin modules:
      - gui/mixins/report_mixin.py   – report formatting & dialog
      - gui/mixins/tree_mixin.py     – tree widget operations
      - gui/mixins/editor_mixin.py   – sheet editor & rules UI
      - gui/mixins/throbber_mixin.py – animated spinner widget

    This class contains only the orchestration core: init, autosave,
    project-level add/remove, run, context-menu template actions,
    destination browse, feedback, and the main() entry point.
    """

    def __init__(self) -> None:
        super().__init__()
        self.title("Excel Turbo Extractor V3")
        self.minsize(1100, 700)

        self.project: ProjectConfig = ProjectConfig()
        self.current_sheet: Optional[SheetConfig] = None
        self.current_source_path: Optional[str] = None
        self.current_recipe_name: Optional[str] = None

        self._rename_entry: Optional[ttk.Entry] = None
        self._rename_item_id: Optional[str] = None
        self._rename_path: Optional[list[int]] = None
        self._rename_kind: Optional[str] = None

        self._loading: bool = False

        self._autosave_dirty: bool = False
        self._autosave_after_id: Optional[str] = None
        self._autosave_periodic_id: Optional[str] = None
        self._autosave_path: str = resolve_autosave_path()

        self._build_ui()
        self._try_load_autosave()
        self._schedule_periodic_autosave()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        build_ui(self)

    # ── Autosave ──────────────────────────────────────────────────────────────

    def _mark_dirty(self) -> None:
        self._autosave_dirty = True
        self._schedule_debounced_autosave()

    def _schedule_debounced_autosave(self) -> None:
        if self._autosave_after_id is not None:
            try:
                self.after_cancel(self._autosave_after_id)
            except Exception:
                pass
        self._autosave_after_id = self.after(1200, self._autosave_now)

    def _schedule_periodic_autosave(self) -> None:
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
            pass

    def _try_load_autosave(self) -> None:
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

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _source_label(self, src: SourceConfig) -> str:
        name = getattr(src, "name", "")
        if isinstance(name, str) and name.strip():
            return name.strip()
        return os.path.basename(src.path)

    # ── Context-menu template actions ─────────────────────────────────────────

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

    # ── Destination browse ────────────────────────────────────────────────────

    def browse_destination(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Select destination XLSX",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not path:
            return
        self.dest_file_var.set(path)

    # ── Project-level add / remove ────────────────────────────────────────────

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
        src_children = self.tree.get_children("")
        s_id = src_children[path[0]]
        self.tree.insert(s_id, "end", text=new_recipe.name)
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

        auto_created_recipe = False
        if len(path) == 1:
            if not source.recipes:
                source.recipes.append(RecipeConfig(name="Recipe1", sheets=[]))
                auto_created_recipe = True
            recipe = source.recipes[0]
        else:
            recipe = source.recipes[path[1]]

        new_sheet = self._make_default_sheet(name="sheet1")
        recipe.sheets.append(new_sheet)

        if auto_created_recipe:
            self.refresh_tree()
        else:
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
            if not recipe.sheets:
                del source.recipes[path[1]]

        self.current_sheet = None
        self.current_source_path = None
        self.current_recipe_name = None
        self.refresh_tree()
        self._clear_editor()
        self._mark_dirty()
        self._reselect_after_remove(path)

    # ── Feedback / progress ───────────────────────────────────────────────────

    def _feedback_clear(self) -> None:
        self.throbber_start()

    def _feedback_key(self, source_path: str, recipe_name: str, sheet_name: str) -> str:
        base = os.path.basename(source_path)
        return f"{base} | {recipe_name} / {sheet_name}"

    def _feedback_set_row(self, key: str, status: str, rows: str, message: str) -> None:
        tree = getattr(self, "feedback_tree", None)
        if tree is None:
            return
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
        try:
            progress_item = event if payload is None and hasattr(event, "source_path") else payload
            if progress_item is None:
                return
            key    = self._feedback_key(progress_item.source_path, progress_item.recipe_name, progress_item.sheet_name)
            status = getattr(progress_item, "status", None) or getattr(progress_item, "message", "") or ""
            rows   = "" if getattr(progress_item, "rows_written", None) is None else str(getattr(progress_item, "rows_written"))
            msg    = getattr(progress_item, "message", "") or ""
            self._feedback_set_row(key, str(status), rows, msg)
        except Exception:
            return

    # ── Run (threaded) ────────────────────────────────────────────────────────

    def _safe_after(self, ms, func, *args) -> None:
        """Schedule func on the main thread, silently ignoring destroyed Tk."""
        try:
            self.after(ms, func, *args)
        except RuntimeError:
            pass  # Tk root already destroyed (e.g. during tests)

    def _run_finished(self, report) -> None:
        """Called on the main thread after a background run completes."""
        try:
            self.throbber_stop()
            title = "Run complete" if report.ok else "Run complete (with errors)"
            self._show_scrollable_report_dialog(title, self._format_run_report(report))
        except Exception:
            pass  # Tk root may be destroyed during tests

    def run_all(self) -> None:
        items = self.project.build_run_items()
        self._feedback_clear()

        def _work():
            try:
                report = engine_run_all(items, on_progress=self._feedback_progress_callback)
            except TypeError:
                report = engine_run_all(items)
                try:
                    for r in getattr(report, "results", []) or []:
                        self._feedback_progress_callback("result", r)
                except Exception:
                    pass
            self._safe_after(0, self._run_finished, report)

        threading.Thread(target=_work, daemon=True).start()

    def run_selected_sheet(self) -> None:
        if not self.current_sheet or not self.current_source_path or not self.current_recipe_name:
            messagebox.showwarning("Select Sheet", "Select a Sheet to run.")
            return

        source_path = self.current_source_path
        sheet_cfg = self.current_sheet
        recipe_name = self.current_recipe_name

        self._feedback_clear()

        def _work():
            try:
                res = engine_run_sheet(
                    source_path,
                    sheet_cfg,
                    recipe_name=recipe_name,
                )
                self._feedback_progress_callback("result", res)
                from core.models import RunReport as _RunReport
                _mini = _RunReport(ok=True, results=[res])
                self._safe_after(0, self._run_finished, _mini)
            except AppError as e:
                from core.models import SheetResult, RunReport as _RunReport
                err_res = SheetResult(
                    source_path=source_path,
                    recipe_name=recipe_name,
                    sheet_name=sheet_cfg.name,
                    dest_file=sheet_cfg.destination.file_path,
                    dest_sheet=sheet_cfg.destination.sheet_name,
                    rows_written=0,
                    message="ERROR",
                    error_code=e.code,
                    error_message=e.message,
                    error_details=e.details,
                )
                self._feedback_progress_callback("error", err_res)
                _mini = _RunReport(ok=False, results=[err_res])
                self._safe_after(0, self._run_finished, _mini)

        threading.Thread(target=_work, daemon=True).start()


def main() -> None:
    from core.autosave import ENV_AUTOSAVE_PATH, resolve_autosave_path
    if not os.environ.get(ENV_AUTOSAVE_PATH):
        os.environ[ENV_AUTOSAVE_PATH] = resolve_autosave_path()
    app = TurboExtractorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
