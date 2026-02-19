\
from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional, Literal, Dict, Any


# ---- Core project tree model ----

@dataclass
class Destination:
    """
    Destination placement for a shaped output table.
    Start row blank ("") means append mode.
    """
    file_path: str
    sheet_name: str = "Sheet1"
    start_col: str = "A"          # anchor column letters, required
    start_row: str = ""           # explicit numeric string, or "" for append mode


@dataclass
class Rule:
    """
    Rules filter rows after row selection and before column selection/paste shaping.
    Column is absolute source column letters (A, D, AA).
    """
    mode: Literal["include", "exclude"] = "include"
    column: str = "A"
    operator: Literal["contains", "equals", "<", ">"] = "equals"
    value: str = ""


@dataclass
class SheetConfig:
    """
    Editable extraction settings (only live on Sheet nodes).
    """
    name: str = "Sheet1"                  # UI label + workbook sheet name (Option A contract)
    workbook_sheet: str = "Sheet1"        # actual workbook sheet to read (same as name under Option A)
    source_start_row: str = ""            # blank means start at row 1; otherwise 1-based row offset into source
    columns_spec: str = ""                # blank means ALL used columns
    rows_spec: str = ""                   # blank means ALL used rows
    paste_mode: Literal["pack", "keep"] = "pack"
    rules_combine: Literal["AND", "OR"] = "AND"
    rules: List[Rule] = field(default_factory=list)
    destination: Destination = field(default_factory=lambda: Destination(file_path=""))


@dataclass
class RecipeConfig:
    name: str = "Recipe1"
    sheets: List[SheetConfig] = field(default_factory=list)


@dataclass
class SourceConfig:
    path: str = ""                        # file path (XLSX/CSV)
    name: str = ""                        # filename (display only)
    recipes: List[RecipeConfig] = field(default_factory=list)


@dataclass
class ProjectConfig:
    sources: List[SourceConfig] = field(default_factory=list)
    last_selected_node_id: Optional[str] = None


# ---- Run reporting ----

@dataclass
class SheetResult:
    source_path: str
    recipe_name: str
    sheet_name: str
    dest_file: str
    dest_sheet: str
    rows_written: int
    message: str = ""
    error_code: Optional[str] = None
    error_message: Optional[str] = None
    error_details: Optional[Dict[str, Any]] = None


@dataclass
class RunReport:
    """
    Returned by engine.run_sheet/run_all. GUI renders this; tests can assert it.
    """
    ok: bool
    results: List[SheetResult] = field(default_factory=list)

    @property
    def has_errors(self) -> bool:
        return any(r.error_code for r in self.results)
