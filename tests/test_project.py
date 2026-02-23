"""
test_project.py — Consolidated project config and templates tests.

Covers:
  - core.project: ProjectConfig serialization, load/save JSON, build_run_items
  - core.templates: source_to_template, apply_template_to_source,
                    save/load template JSON, default template management
"""
from __future__ import annotations

import os
from pathlib import Path

from core.models import Destination, Rule, SheetConfig
from core.project import ProjectConfig, RecipeConfig, SourceConfig
from core import templates as tpl


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _make_source(path: str) -> SourceConfig:
    sh = SheetConfig(
        name="SheetB", workbook_sheet="SheetB",
        source_start_row="4", columns_spec="A,C", rows_spec="1-3",
        paste_mode="keep", rules_combine="OR",
        rules=[Rule(mode="include", column="B", operator="contains", value="beta")],
        destination=Destination(file_path="out.xlsx", sheet_name="Dest",
                                start_col="D", start_row=""),
    )
    r = RecipeConfig(name="RecipeX", sheets=[sh])
    return SourceConfig(path=path, recipes=[r])


# ══════════════════════════════════════════════════════════════════════════════
# PROJECTCONFIG — SERIALIZATION
# ══════════════════════════════════════════════════════════════════════════════

def test_project_config_save_load_roundtrip(tmp_path):
    proj = ProjectConfig(sources=[
        SourceConfig(path="a.xlsx", recipes=[
            RecipeConfig(name="R1", sheets=[
                SheetConfig(name="S1", workbook_sheet="S1",
                            destination=Destination(file_path="o.xlsx")),
                SheetConfig(name="S2", workbook_sheet="S2",
                            destination=Destination(file_path="o.xlsx")),
            ]),
            RecipeConfig(name="R2", sheets=[
                SheetConfig(name="S3", workbook_sheet="S3",
                            destination=Destination(file_path="o.xlsx")),
            ]),
        ]),
        SourceConfig(path="b.csv", recipes=[
            RecipeConfig(name="R3", sheets=[
                SheetConfig(name="S4", workbook_sheet="S4",
                            destination=Destination(file_path="o2.xlsx")),
            ]),
        ]),
    ])

    p = str(tmp_path / "proj.json")
    proj.save_json(p)
    loaded = ProjectConfig.load_json(p)

    assert len(loaded.sources) == 2
    assert loaded.sources[0].recipes[0].sheets[1].name == "S2"
    assert loaded.sources[0].recipes[1].name == "R2"
    assert loaded.sources[1].path == "b.csv"


def test_project_build_run_items_order(tmp_path):
    proj = ProjectConfig(sources=[
        SourceConfig(path="a.xlsx", recipes=[
            RecipeConfig(name="R1", sheets=[
                SheetConfig(name="S1", workbook_sheet="S1",
                            destination=Destination(file_path="o.xlsx")),
                SheetConfig(name="S2", workbook_sheet="S2",
                            destination=Destination(file_path="o.xlsx")),
            ]),
            RecipeConfig(name="R2", sheets=[
                SheetConfig(name="S3", workbook_sheet="S3",
                            destination=Destination(file_path="o.xlsx")),
            ]),
        ]),
        SourceConfig(path="b.csv", recipes=[
            RecipeConfig(name="R3", sheets=[
                SheetConfig(name="S4", workbook_sheet="S4",
                            destination=Destination(file_path="o2.xlsx")),
            ]),
        ]),
    ])

    p = str(tmp_path / "proj.json")
    proj.save_json(p)
    loaded = ProjectConfig.load_json(p)

    items = loaded.build_run_items()
    assert len(items) == 4
    assert [i[1] for i in items] == ["R1", "R1", "R2", "R3"]


def test_project_config_empty_project_roundtrip(tmp_path):
    proj = ProjectConfig(sources=[])
    p    = str(tmp_path / "empty.json")
    proj.save_json(p)
    loaded = ProjectConfig.load_json(p)
    assert loaded.sources == []
    assert loaded.build_run_items() == []


def test_project_config_preserves_all_sheet_fields(tmp_path):
    sh = SheetConfig(
        name="Full", workbook_sheet="FullWB",
        source_start_row="3", columns_spec="A-E", rows_spec="2-8",
        paste_mode="keep", rules_combine="OR",
        rules=[Rule(mode="include", column="C", operator="contains", value="test")],
        destination=Destination(file_path="dest.xlsx", sheet_name="Out",
                                start_col="B", start_row="5"),
    )
    proj = ProjectConfig(sources=[
        SourceConfig(path="src.xlsx", recipes=[
            RecipeConfig(name="R1", sheets=[sh])
        ])
    ])
    p = str(tmp_path / "p.json")
    proj.save_json(p)
    loaded = ProjectConfig.load_json(p)

    sh2 = loaded.sources[0].recipes[0].sheets[0]
    assert sh2.source_start_row == "3"
    assert sh2.columns_spec     == "A-E"
    assert sh2.rows_spec        == "2-8"
    assert sh2.paste_mode       == "keep"
    assert sh2.rules_combine    == "OR"
    assert sh2.rules[0].column  == "C"
    assert sh2.destination.start_col == "B"
    assert sh2.destination.start_row == "5"


# ══════════════════════════════════════════════════════════════════════════════
# TEMPLATES — SAVE / LOAD / APPLY
# ══════════════════════════════════════════════════════════════════════════════

def test_source_template_roundtrip_preserves_path(tmp_path):
    src1     = _make_source("/tmp/source1.xlsx")
    template = tpl.source_to_template(src1)

    p = tmp_path / "t.json"
    tpl.save_template_json(template, str(p))
    loaded = tpl.load_template_json(str(p))

    src2 = _make_source("/tmp/DIFFERENT.xlsx")
    src2.recipes = []
    tpl.apply_template_to_source(src2, loaded)

    assert src2.path == "/tmp/DIFFERENT.xlsx"
    assert len(src2.recipes) == 1
    assert src2.recipes[0].name == "RecipeX"
    assert src2.recipes[0].sheets[0].name == "SheetB"
    assert src2.recipes[0].sheets[0].source_start_row == "4"
    assert src2.recipes[0].sheets[0].destination.start_col == "D"
    assert src2.recipes[0].sheets[0].rules[0].operator == "contains"


def test_template_all_sheet_fields_roundtrip(tmp_path):
    sh = SheetConfig(
        name="Full", workbook_sheet="Full",
        source_start_row="2", columns_spec="A-D", rows_spec="5-10",
        paste_mode="keep", rules_combine="OR",
        rules=[Rule(mode="exclude", column="B", operator=">", value="99")],
        destination=Destination(file_path="x.xlsx", sheet_name="S",
                                start_col="C", start_row="3"),
    )
    src  = SourceConfig(path="s.xlsx", recipes=[
        RecipeConfig(name="R1", sheets=[sh])
    ])
    tmpl = tpl.source_to_template(src)
    p    = str(tmp_path / "t.json")
    tpl.save_template_json(tmpl, p)
    loaded = tpl.load_template_json(p)

    tgt = _make_source("/other.xlsx")
    tgt.recipes = []
    tpl.apply_template_to_source(tgt, loaded)

    sh2 = tgt.recipes[0].sheets[0]
    assert sh2.source_start_row == "2"
    assert sh2.columns_spec     == "A-D"
    assert sh2.rows_spec        == "5-10"
    assert sh2.paste_mode       == "keep"
    assert sh2.rules_combine    == "OR"
    assert sh2.rules[0].mode    == "exclude"
    assert sh2.rules[0].value   == "99"
    assert sh2.destination.start_col == "C"
    assert sh2.destination.start_row == "3"


def test_template_does_not_include_source_path(tmp_path):
    src  = _make_source("/private/path/source.xlsx")
    tmpl = tpl.source_to_template(src)
    assert "path" not in tmpl or tmpl.get("path") != "/private/path/source.xlsx"


def test_default_template_set_load_reset(tmp_path, monkeypatch):
    src      = _make_source("/tmp/source.xlsx")
    template = tpl.source_to_template(src)

    default_path = tmp_path / "default.json"
    monkeypatch.setenv(tpl.ENV_DEFAULT_TEMPLATE_PATH, str(default_path))

    assert tpl.load_default_template() is None

    saved_to = tpl.set_default_template(template)
    assert saved_to == str(default_path)
    assert tpl.load_default_template() is not None

    tpl.reset_default_template()
    assert tpl.load_default_template() is None


def test_template_apply_replaces_all_recipes(tmp_path):
    """Applying a template with 2 recipes replaces all existing recipes."""
    sh1 = SheetConfig(name="S1", workbook_sheet="S1",
                      destination=Destination(file_path="o.xlsx"))
    sh2 = SheetConfig(name="S2", workbook_sheet="S2",
                      destination=Destination(file_path="o.xlsx"))
    src = SourceConfig(path="s.xlsx", recipes=[
        RecipeConfig(name="Recipe1", sheets=[sh1]),
        RecipeConfig(name="Recipe2", sheets=[sh2]),
    ])
    tmpl = tpl.source_to_template(src)
    p    = str(tmp_path / "t.json")
    tpl.save_template_json(tmpl, p)
    loaded = tpl.load_template_json(p)

    tgt = _make_source("/other.xlsx")
    tpl.apply_template_to_source(tgt, loaded)
    assert len(tgt.recipes) == 2
    assert tgt.recipes[0].name == "Recipe1"
    assert tgt.recipes[1].name == "Recipe2"
