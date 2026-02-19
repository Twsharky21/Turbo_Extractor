from __future__ import annotations

import os
from pathlib import Path

from core.models import Destination, Rule, SheetConfig
from core.project import ProjectConfig, SourceConfig, RecipeConfig
from core import templates as tpl


def _make_source(path: str) -> SourceConfig:
    sh = SheetConfig(
        name="SheetB",
        workbook_sheet="SheetB",
        columns_spec="A,C",
        rows_spec="1-3",
        paste_mode="keep",
        rules_combine="OR",
        rules=[Rule(mode="include", column="B", operator="contains", value="beta")],
        destination=Destination(file_path="out.xlsx", sheet_name="Dest", start_col="D", start_row=""),
    )
    r = RecipeConfig(name="RecipeX", sheets=[sh])
    return SourceConfig(path=path, recipes=[r])


def test_source_template_roundtrip_apply_does_not_change_path(tmp_path: Path):
    src1 = _make_source("/tmp/source1.xlsx")
    template = tpl.source_to_template(src1)

    # Save/load JSON
    p = tmp_path / "t.json"
    tpl.save_template_json(template, str(p))
    loaded = tpl.load_template_json(str(p))

    src2 = _make_source("/tmp/DIFFERENT.xlsx")
    src2.recipes = []
    tpl.apply_template_to_source(src2, loaded)

    assert src2.path == "/tmp/DIFFERENT.xlsx"  # path preserved
    assert len(src2.recipes) == 1
    assert src2.recipes[0].name == "RecipeX"
    assert src2.recipes[0].sheets[0].name == "SheetB"
    assert src2.recipes[0].sheets[0].destination.start_col == "D"
    assert src2.recipes[0].sheets[0].rules[0].operator == "contains"


def test_default_template_set_load_reset(tmp_path: Path, monkeypatch):
    src = _make_source("/tmp/source.xlsx")
    template = tpl.source_to_template(src)

    default_path = tmp_path / "default.json"
    monkeypatch.setenv(tpl.ENV_DEFAULT_TEMPLATE_PATH, str(default_path))

    assert tpl.load_default_template() is None
    saved_to = tpl.set_default_template(template)
    assert saved_to == str(default_path)
    assert tpl.load_default_template() is not None

    tpl.reset_default_template()
    assert tpl.load_default_template() is None
