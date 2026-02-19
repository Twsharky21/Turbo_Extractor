\
from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from dataclasses import asdict
from typing import List, Tuple, Optional

from core.engine import run_all, RunItem
from core.errors import AppError


class TurboExtractorApp(tk.Tk):
    """
    Minimal GUI shell (V3):
    - Left: placeholder tree (Sources/Recipes/Sheets)
    - Right: placeholder editor area
    - Bottom-right: Run / Run All buttons
    - Uses core.engine.run_all() for execution

    This file is intentionally minimal and import-safe for pytest.
    """

    def __init__(self) -> None:
        super().__init__()
        self.title("Excel Turbo Extractor V3 (Minimal GUI)")
        self.minsize(1000, 650)

        self._build_ui()

        # For now, we keep an in-memory run list (RunItem tuples).
        # Later, this will be derived from the ProjectConfig tree.
        self._run_items: List[RunItem] = []

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        root = ttk.Frame(self, padding=10)
        root.grid(row=0, column=0, sticky="nsew")
        root.columnconfigure(0, weight=1)
        root.columnconfigure(1, weight=3)
        root.rowconfigure(0, weight=1)

        # Left panel: tree placeholder
        left = ttk.LabelFrame(root, text="Sources / Recipes / Sheets", padding=8)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(left, show="tree", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=yscroll.set)

        btns = ttk.Frame(left)
        btns.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(btns, text="Add Source(s)...", command=self._add_sources_placeholder).pack(side="left")
        ttk.Button(btns, text="Clear", command=self._clear_placeholder).pack(side="left", padx=(8, 0))

        # Right panel: editor placeholder
        right = ttk.LabelFrame(root, text="Editor (Selected Sheet)", padding=8)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=1)

        self.editor_text = tk.Text(right, height=10, wrap="word")
        self.editor_text.grid(row=0, column=0, sticky="nsew")
        self.editor_text.insert("1.0", "Minimal V3 GUI shell.\n\nNext steps:\n- Load/save ProjectConfig\n- Populate tree\n- Bind selection to editor\n- Build Sheet editor fields\n")

        bottom = ttk.Frame(right)
        bottom.grid(row=1, column=0, sticky="e", pady=(10, 0))

        ttk.Button(bottom, text="Run Selected (TODO)", command=self._run_selected_placeholder).pack(side="right")
        ttk.Button(bottom, text="Run All", command=self._run_all_now).pack(side="right", padx=(0, 8))

    # ----- Placeholder actions (will be replaced by real ProjectConfig editing) -----

    def _add_sources_placeholder(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Add source file(s)",
            filetypes=[("Excel/CSV", "*.xlsx *.xlsm *.csv"), ("All files", "*.*")]
        )
        if not paths:
            return

        # For minimal shell: add sources as tree nodes only; no recipes/sheets yet.
        for p in paths:
            self.tree.insert("", "end", text=p)
        self.editor_text.insert("end", f"\nAdded {len(paths)} source(s) to tree (placeholder).\n")

    def _clear_placeholder(self) -> None:
        for item in self.tree.get_children(""):
            self.tree.delete(item)
        self._run_items.clear()
        self.editor_text.insert("end", "\nCleared tree and run list.\n")

    def _run_selected_placeholder(self) -> None:
        messagebox.showinfo("Not implemented", "Run Selected is not implemented in the minimal shell yet.")

    # ----- Execution -----

    def set_run_items(self, items: List[RunItem]) -> None:
        """
        Allows tests or future config loader to inject a run list.
        """
        self._run_items = list(items)

    def _run_all_now(self) -> None:
        if not self._run_items:
            messagebox.showwarning(
                "No run items",
                "No run items are configured yet.\n\n"
                "This is the minimal shell.\n"
                "Next: tree->config wiring will build run items.",
            )
            return

        report = run_all(self._run_items)

        # Build a human-readable report
        lines = []
        for r in report.results:
            if r.error_code:
                lines.append(f"❌ {r.source_path} | {r.recipe_name}/{r.sheet_name}: {r.error_code} - {r.error_message}")
            else:
                lines.append(f"✅ {r.source_path} | {r.recipe_name}/{r.sheet_name}: {r.rows_written} rows")

        title = "Run complete" if report.ok else "Run complete (with errors)"
        messagebox.showinfo(title, "\n".join(lines) if lines else "(no results)")


def main() -> None:
    app = TurboExtractorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
