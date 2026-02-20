from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Optional

# NOTE: GUI-only module. No business logic here.
# This module exists to keep gui/app.py smaller and make UI issues easier to isolate.

def build_ui(app) -> None:
    # Overall layout: top toolbar, then left tree + right editor (monolith-like)
    app.columnconfigure(0, weight=1)
    app.rowconfigure(0, weight=1)

    root = ttk.Frame(app, padding=8)
    root.grid(row=0, column=0, sticky="nsew")
    # Fixed minsize on the left column prevents the tree from expanding when
    # the right panel's content collapses (source/recipe selection hides editor).
    root.columnconfigure(0, weight=1, minsize=220, uniform="cols")
    root.columnconfigure(1, weight=3, minsize=660, uniform="cols")
    root.rowconfigure(1, weight=1)

    # ----- TOP TOOLBAR -----
    topbar = ttk.Frame(root)
    topbar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
    for i in range(8):
        topbar.columnconfigure(i, weight=0)
    topbar.columnconfigure(7, weight=1)

    # Buttons: Add Source, Move (up/down), Add Recipe, Add Sheet, Remove Selected
    ttk.Button(topbar, text="Add Source (XLSX/CSV)", style="RunAccent.TButton", command=app.add_sources).grid(row=0, column=0, padx=(0, 6))
    move_frame = ttk.Frame(topbar)
    move_frame.grid(row=0, column=1, padx=(0, 6))
    # Two small stacked move buttons apply to Sources, Recipes, or Sheets (based on selection)
    ttk.Button(move_frame, text="MOVE ▲", width=8, command=app.move_selected_up).grid(row=0, column=0, sticky="ew")
    ttk.Button(move_frame, text="MOVE ▼", width=8, command=app.move_selected_down).grid(row=1, column=0, sticky="ew", pady=(2, 0))
    ttk.Button(topbar, text="Add Recipe", command=app.add_recipe).grid(row=0, column=2, padx=(0, 6))
    ttk.Button(topbar, text="Add Sheet", command=app.add_sheet).grid(row=0, column=3, padx=(0, 6))
    ttk.Button(topbar, text="Remove Selected", command=app.remove_selected).grid(row=0, column=4, padx=(0, 6))

    # ----- LEFT: TREE -----
    left = ttk.Frame(root)
    left.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
    left.columnconfigure(0, weight=1)
    left.rowconfigure(0, weight=1)

    app.tree = ttk.Treeview(left, show="tree", selectmode="browse")
    app.tree.grid(row=0, column=0, sticky="nsew")
    app.tree.bind("<<TreeviewSelect>>", app._on_tree_select)
    app.tree.bind("<Button-3>", app._on_tree_right_click)

    yscroll = ttk.Scrollbar(left, orient="vertical", command=app.tree.yview)
    yscroll.grid(row=0, column=1, sticky="ns")
    app.tree.configure(yscrollcommand=yscroll.set)

    # ----- RIGHT: EDITOR -----
    right = ttk.Frame(root)
    right.grid(row=1, column=1, sticky="nsew")
    right.columnconfigure(0, weight=1)
    right.rowconfigure(2, weight=1)

    # ----- RIGHT: SELECTION HEADER (name-only view) -----
    # NOTE: stored as app.selection_box so tests can check winfo_ismapped()
    app.selection_name_var = tk.StringVar(value="")
    app.selection_box = ttk.LabelFrame(right, text="Selected", padding=10)
    app.selection_box.grid(row=0, column=0, sticky="ew")
    app.selection_box.columnconfigure(0, weight=1)
    ttk.Label(app.selection_box, textvariable=app.selection_name_var).grid(row=0, column=0, sticky="w")

    # ----- Selected Sheet (within Recipe) -----
    app.sheet_box = ttk.LabelFrame(right, text="Selected Sheet (within Recipe)", padding=10)
    app.sheet_box.grid(row=1, column=0, sticky="ew")
    app.sheet_box.columnconfigure(1, weight=1)

    ttk.Label(app.sheet_box, text="Columns (e.g., A,C,AC-ZZ):").grid(row=0, column=0, sticky="w")
    app.columns_var = tk.StringVar()
    ttk.Entry(app.sheet_box, textvariable=app.columns_var).grid(row=0, column=1, sticky="ew", padx=(10, 0))
    app.columns_var.trace_add("write", app._push_editor_to_sheet)

    ttk.Label(app.sheet_box, text="Rows (e.g., 1-3,9-80,117):").grid(row=1, column=0, sticky="w", pady=(6, 0))
    app.rows_var = tk.StringVar()
    ttk.Entry(app.sheet_box, textvariable=app.rows_var).grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=(6, 0))
    app.rows_var.trace_add("write", app._push_editor_to_sheet)

    ttk.Label(app.sheet_box, text="Source Start Row:").grid(row=2, column=0, sticky="w", pady=(6, 0))
    app.source_start_row_var = tk.StringVar()
    ttk.Entry(app.sheet_box, textvariable=app.source_start_row_var, width=10).grid(row=2, column=1, sticky="w", padx=(10, 0), pady=(6, 0))
    app.source_start_row_var.trace_add("write", app._push_editor_to_sheet)

    ttk.Label(app.sheet_box, text="Column paste mode:").grid(row=3, column=0, sticky="w", pady=(6, 0))
    app.paste_var = tk.StringVar()
    app.paste_combo = ttk.Combobox(
        app.sheet_box,
        textvariable=app.paste_var,
        values=["Pack Together", "Keep Format"],
        state="readonly",
        width=18,
    )
    app.paste_combo.grid(row=3, column=1, sticky="w", padx=(10, 0), pady=(6, 0))
    app.paste_combo.bind("<<ComboboxSelected>>", app._push_editor_to_sheet)

    # ----- RULES -----
    app.rules_box = ttk.LabelFrame(right, text="Rules", padding=10)
    app.rules_box.grid(row=2, column=0, sticky="nsew", pady=(10, 0))
    app.rules_box.columnconfigure(0, weight=1)
    app.rules_box.rowconfigure(2, weight=1)
    right.rowconfigure(2, weight=1)

    top_rules = ttk.Frame(app.rules_box)
    top_rules.grid(row=0, column=0, sticky="ew")
    top_rules.columnconfigure(1, weight=1)

    ttk.Label(top_rules, text="Rules combine:").grid(row=0, column=0, sticky="w")
    app.combine_var = tk.StringVar()
    app.combine_combo = ttk.Combobox(top_rules, textvariable=app.combine_var, values=["AND", "OR"], state="readonly", width=8)
    app.combine_combo.grid(row=0, column=1, sticky="w", padx=(10, 0))
    app.combine_combo.bind("<<ComboboxSelected>>", app._push_editor_to_sheet)

    ttk.Button(top_rules, text="+ Add rule", command=app.add_rule).grid(row=0, column=2, sticky="w", padx=(20, 0))

    # Header row
    hdr = ttk.Frame(app.rules_box)
    hdr.grid(row=1, column=0, sticky="ew", pady=(8, 4))
    ttk.Label(hdr, text="Include/Exclude").grid(row=0, column=0, sticky="w", padx=(0, 10))
    ttk.Label(hdr, text="Column").grid(row=0, column=1, sticky="w", padx=(0, 10))
    ttk.Label(hdr, text="Operator").grid(row=0, column=2, sticky="w", padx=(0, 10))
    ttk.Label(hdr, text="Value").grid(row=0, column=3, sticky="w", padx=(0, 10))

    # Scrollable rules area
    rules_area = ttk.Frame(app.rules_box)
    rules_area.grid(row=2, column=0, sticky="nsew")
    rules_area.columnconfigure(0, weight=1)
    rules_area.rowconfigure(0, weight=1)

    app.rules_canvas = tk.Canvas(rules_area, height=260)
    app.rules_canvas.grid(row=0, column=0, sticky="nsew")

    rules_scroll = ttk.Scrollbar(rules_area, orient="vertical", command=app.rules_canvas.yview)
    rules_scroll.grid(row=0, column=1, sticky="ns")
    app.rules_canvas.configure(yscrollcommand=rules_scroll.set)

    app.rules_frame = ttk.Frame(app.rules_canvas)
    app.rules_canvas.create_window((0, 0), window=app.rules_frame, anchor="nw")

    app.rules_frame.bind(
        "<Configure>",
        lambda e: app.rules_canvas.configure(scrollregion=app.rules_canvas.bbox("all")),
    )

    # ----- DESTINATION -----
    app.dest_box = ttk.LabelFrame(right, text="Destination", padding=10)
    app.dest_box.grid(row=3, column=0, sticky="ew", pady=(10, 0))
    app.dest_box.columnconfigure(1, weight=1)

    ttk.Label(app.dest_box, text="File:").grid(row=0, column=0, sticky="w")
    app.dest_file_var = tk.StringVar()
    ttk.Entry(app.dest_box, textvariable=app.dest_file_var).grid(row=0, column=1, sticky="ew", padx=(10, 10))
    ttk.Button(app.dest_box, text="Browse", command=app.browse_destination).grid(row=0, column=2, sticky="ew")
    app.dest_file_var.trace_add("write", app._push_editor_to_sheet)

    ttk.Label(app.dest_box, text="Sheet name:").grid(row=1, column=0, sticky="w", pady=(6, 0))
    app.dest_sheet_var = tk.StringVar()
    ttk.Entry(app.dest_box, textvariable=app.dest_sheet_var).grid(row=1, column=1, columnspan=2, sticky="ew", padx=(10, 0), pady=(6, 0))
    app.dest_sheet_var.trace_add("write", app._push_editor_to_sheet)

    ttk.Label(app.dest_box, text="Start column (e.g., A, D, AA):").grid(row=2, column=0, sticky="w", pady=(6, 0))
    start_frame = ttk.Frame(app.dest_box)
    start_frame.grid(row=2, column=1, columnspan=2, sticky="w", padx=(10, 0), pady=(6, 0))

    app.start_col_var = tk.StringVar()
    ttk.Entry(start_frame, textvariable=app.start_col_var, width=8).grid(row=0, column=0, sticky="w")
    app.start_col_var.trace_add("write", app._push_editor_to_sheet)

    ttk.Label(start_frame, text="Start row:").grid(row=0, column=1, sticky="w", padx=(15, 6))
    app.start_row_var = tk.StringVar()
    ttk.Entry(start_frame, textvariable=app.start_row_var, width=10).grid(row=0, column=2, sticky="w")
    app.start_row_var.trace_add("write", app._push_editor_to_sheet)

    # ----- BOTTOM: STATUS + RUN BUTTONS -----
    bottom = ttk.Frame(right)
    bottom.grid(row=4, column=0, sticky="ew", pady=(10, 0))
    bottom.columnconfigure(0, weight=1)

    app.status_var = tk.StringVar(value="Idle")
    ttk.Label(bottom, textvariable=app.status_var).grid(row=0, column=0, sticky="w")

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

    ttk.Button(run_btns, text="RUN", style="RunAccent.TButton", command=app.run_selected_sheet).pack(side="right", padx=(6, 0))
    ttk.Button(run_btns, text="RUN ALL", style="RunAccent.TButton", command=app.run_all).pack(side="right")

    # Context menu (Source only)
    app._source_menu = tk.Menu(app, tearoff=0)
    app._source_menu.add_command(label="Save Template...", command=app._ctx_save_template)
    app._source_menu.add_command(label="Load Template...", command=app._ctx_load_template)
    app._source_menu.add_separator()
    app._source_menu.add_command(label="Set Default", command=app._ctx_set_default)
    app._source_menu.add_command(label="Reset Default", command=app._ctx_reset_default)
    app._ctx_source_index = None

    # Context menu (Recipe)
    app._recipe_menu = tk.Menu(app, tearoff=0)
    app._recipe_menu.add_command(label="Rename Recipe", command=app._ctx_rename_recipe)
    app._ctx_recipe_path = None

    # Context menu (Sheet)
    app._sheet_menu = tk.Menu(app, tearoff=0)
    app._sheet_menu.add_command(label="Rename Sheet", command=app._ctx_rename_sheet)
    app._ctx_sheet_path = None

    # Hide sheet/rules/dest initially; selection_box stays always visible.
    app.sheet_box.grid_remove()
    app.rules_box.grid_remove()
    app.dest_box.grid_remove()
