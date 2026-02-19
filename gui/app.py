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
    Minimal GUI wired to ProjectConfig.
    Still intentionally thin — no full editor yet.
    """

    def __init__(self) -> None:
        super().__init__()
        self.title("Excel Turbo Extractor V3")
        self.minsize(1000, 650)

        self.project: ProjectConfig = ProjectConfig()

        self._build_ui()

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

        yscroll = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=yscroll.set)

        btns = ttk.Frame(left)
        btns.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))

        ttk.Button(btns, text="New Project", command=self.new_project).pack(side="left")
        ttk.Button(btns, text="Load Project", command=self.load_project).pack(side="left", padx=(8, 0))
        ttk.Button(btns, text="Save Project", command=self.save_project).pack(side="left", padx=(8, 0))

        # ----- RIGHT PANEL -----

        right = ttk.LabelFrame(root, text="Project Info", padding=8)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=1)

        self.info_text = tk.Text(right, wrap="word")
        self.info_text.grid(row=0, column=0, sticky="nsew")
        self._refresh_info()

        bottom = ttk.Frame(right)
        bottom.grid(row=1, column=0, sticky="e", pady=(10, 0))

        ttk.Button(bottom, text="Run All", command=self.run_all_now).pack(side="right")

    # ----- Project State -----

    def new_project(self) -> None:
        self.project = ProjectConfig()
        self._refresh_tree()
        self._refresh_info()

    def load_project(self) -> None:
        path = filedialog.askopenfilename(
            title="Load Project",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not path:
            return

        self.project = ProjectConfig.load_json(path)
        self._refresh_tree()
        self._refresh_info()

    def save_project(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save Project",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not path:
            return

        self.project.save_json(path)
        messagebox.showinfo("Saved", f"Project saved to:\n{path}")

    # ----- Execution -----

    def run_all_now(self) -> None:
        items = self.project.build_run_items()
        if not items:
            messagebox.showwarning("Nothing to run", "Project has no sheets configured.")
            return

        report = run_all(items)

        lines = []
        for r in report.results:
            if r.error_code:
                lines.append(f"❌ {r.source_path} | {r.recipe_name}/{r.sheet_name}: {r.error_code}")
            else:
                lines.append(f"✅ {r.source_path} | {r.recipe_name}/{r.sheet_name}: {r.rows_written} rows")

        title = "Run complete" if report.ok else "Run complete (with errors)"
        messagebox.showinfo(title, "\n".join(lines))

    # ----- UI Refresh -----

    def _refresh_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        for source in self.project.sources:
            s_id = self.tree.insert("", "end", text=source.path)
            for recipe in source.recipes:
                r_id = self.tree.insert(s_id, "end", text=recipe.name)
                for sheet in recipe.sheets:
                    self.tree.insert(r_id, "end", text=sheet.name)

    def _refresh_info(self) -> None:
        self.info_text.delete("1.0", "end")
        self.info_text.insert(
            "1.0",
            f"Sources: {len(self.project.sources)}\n\n"
            "This is the minimal ProjectConfig-wired GUI.\n"
            "Next: full sheet editor wiring.\n"
        )


def main() -> None:
    app = TurboExtractorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
