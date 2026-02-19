\
from __future__ import annotations

from dataclasses import dataclass, asdict, field
from typing import List, Dict, Any
import json

from .models import SheetConfig, Destination, Rule


@dataclass
class RecipeConfig:
    name: str
    sheets: List[SheetConfig] = field(default_factory=list)


@dataclass
class SourceConfig:
    path: str
    recipes: List[RecipeConfig] = field(default_factory=list)


@dataclass
class ProjectConfig:
    sources: List[SourceConfig] = field(default_factory=list)

    # ---------- Serialization ----------

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "ProjectConfig":
        sources: List[SourceConfig] = []

        for s in data.get("sources", []):
            recipes: List[RecipeConfig] = []
            for r in s.get("recipes", []):
                sheets: List[SheetConfig] = []
                for sh in r.get("sheets", []):
                    rules = [
                        Rule(**rule_dict)
                        for rule_dict in sh.get("rules", [])
                    ]
                    dest = Destination(**sh["destination"])
                    sheet = SheetConfig(
                        name=sh["name"],
                        workbook_sheet=sh["workbook_sheet"],
                        source_start_row=sh.get("source_start_row", ""),
                        columns_spec=sh["columns_spec"],
                        rows_spec=sh["rows_spec"],
                        paste_mode=sh["paste_mode"],
                        rules_combine=sh["rules_combine"],
                        rules=rules,
                        destination=dest,
                    )
                    sheets.append(sheet)
                recipes.append(RecipeConfig(name=r["name"], sheets=sheets))
            sources.append(SourceConfig(path=s["path"], recipes=recipes))

        return cls(sources=sources)

    # ---------- File IO ----------

    def save_json(self, path: str) -> None:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, indent=2)

    @classmethod
    def load_json(cls, path: str) -> "ProjectConfig":
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return cls.from_dict(data)

    # ---------- Execution Flattening ----------

    def build_run_items(self):
        """
        Returns list of (source_path, recipe_name, sheet_cfg)
        in tree order.
        """
        items = []
        for source in self.sources:
            for recipe in source.recipes:
                for sheet in recipe.sheets:
                    items.append((source.path, recipe.name, sheet))
        return items
