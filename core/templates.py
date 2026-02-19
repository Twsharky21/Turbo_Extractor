from __future__ import annotations

import json
import os
from copy import deepcopy
from dataclasses import asdict
from pathlib import Path
from typing import Any, Dict, Optional

from .models import Destination, Rule, SheetConfig
from .project import RecipeConfig, SourceConfig


ENV_DEFAULT_TEMPLATE_PATH = "TURBO_DEFAULT_TEMPLATE_PATH"


def resolve_default_template_path(project_root: Optional[str] = None) -> str:
    """Resolve the default Source template path.

    Priority:
    1) TURBO_DEFAULT_TEMPLATE_PATH env var (absolute or relative)
    2) User-home scoped default: ~/.turbo_extractor_v3/default_source_template.json
    """
    env = os.getenv(ENV_DEFAULT_TEMPLATE_PATH)
    if env:
        p = Path(env)
        if not p.is_absolute():
            base = Path(project_root) if project_root else Path.cwd()
            p = base / p
        return str(p)

    base = Path.home() / ".turbo_extractor_v3"
    return str(base / "default_source_template.json")


def source_to_template(source: SourceConfig) -> Dict[str, Any]:
    """Serialize a SourceConfig into a portable template dict.

    Template intentionally excludes the Source path.
    """
    recipes = []
    for r in source.recipes:
        sheets = [asdict(sh) for sh in r.sheets]
        recipes.append({"name": r.name, "sheets": sheets})
    return {"version": 1, "recipes": recipes}


def apply_template_to_source(source: SourceConfig, template: Dict[str, Any]) -> None:
    """Overwrite source.recipes using the given template.

    Does NOT modify source.path.
    """
    recipes = []
    for r in template.get("recipes", []):
        sheets = []
        for sh in r.get("sheets", []):
            rules = [Rule(**rd) for rd in sh.get("rules", [])]
            dest_dict = sh.get("destination", {})
            dest = Destination(**dest_dict)
            sheet = SheetConfig(
                name=sh.get("name", "Sheet1"),
                workbook_sheet=sh.get("workbook_sheet", sh.get("name", "Sheet1")),
                source_start_row=sh.get("source_start_row", ""),
                columns_spec=sh.get("columns_spec", ""),
                rows_spec=sh.get("rows_spec", ""),
                paste_mode=sh.get("paste_mode", "pack"),
                rules_combine=sh.get("rules_combine", "AND"),
                rules=rules,
                destination=dest,
            )
            sheets.append(sheet)
        recipes.append(RecipeConfig(name=r.get("name", "Recipe1"), sheets=sheets))
    source.recipes = recipes


def save_template_json(template: Dict[str, Any], path: str) -> None:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(template, indent=2), encoding="utf-8")


def load_template_json(path: str) -> Dict[str, Any]:
    p = Path(path)
    return json.loads(p.read_text(encoding="utf-8"))


def set_default_template(template: Dict[str, Any], path: Optional[str] = None) -> str:
    dest = path or resolve_default_template_path()
    save_template_json(template, dest)
    return dest


def load_default_template(path: Optional[str] = None) -> Optional[Dict[str, Any]]:
    p = Path(path or resolve_default_template_path())
    if not p.exists():
        return None
    return load_template_json(str(p))


def reset_default_template(path: Optional[str] = None) -> None:
    p = Path(path or resolve_default_template_path())
    if p.exists():
        p.unlink()


def clone_template(template: Dict[str, Any]) -> Dict[str, Any]:
    return deepcopy(template)
