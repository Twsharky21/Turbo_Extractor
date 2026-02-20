"""Tests for core.templates — source template save/load/apply and default template management."""
import os
from pathlib import Path

from core.models import Destination, Rule, SheetConfig
from core.project import ProjectConfig, SourceConfig, RecipeConfig
from core import templates as tpl


def _make_source(path: str) -> SourceConfig:
    sh = SheetConfig(
        name="SheetB", workbook_sheet="SheetB",
        source_start_row="4", columns_spec="A,C", rows_spec="1-3",
        paste_mode="keep", rules_combine="OR",
        rules=[Rule(mode="include", column="B", operator="contains", value="beta")],
        destination=Destination(file_path="out.xlsx", sheet_name="Dest", start_col="D", start_row=""),
    )
    r = RecipeConfig(name="RecipeX", sheets=[sh])
    return SourceConfig(path=path, recipes=[r])


def test_source_template_roundtrip_preserves_path(tmp_path: Path):
    src1 = _make_source("/tmp/source1.xlsx")
    template = tpl.source_to_template(src1)

    p = tmp_path / "t.json"
    tpl.save_template_json(template, str(p))
    loaded = tpl.load_template_json(str(p))

    src2 = _make_source("/tmp/DIFFERENT.xlsx")
    src2.recipes = []
    tpl.apply_template_to_source(src2, loaded)

    assert src2.path == "/tmp/DIFFERENT.xlsx"  # path must not change
    assert len(src2.recipes) == 1
    assert src2.recipes[0].name == "RecipeX"
    assert src2.recipes[0].sheets[0].name == "SheetB"
    assert src2.recipes[0].sheets[0].source_start_row == "4"
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


def test_template_all_sheet_fields_roundtrip(tmp_path: Path):
    sh = SheetConfig(
        name="Full", workbook_sheet="Full",
        source_start_row="2", columns_spec="A-D", rows_spec="5-10",
        paste_mode="keep", rules_combine="OR",
        rules=[
            Rule(mode="include", column="C", operator=">", value="100"),
            Rule(mode="exclude", column="A", operator="equals", value="skip"),
        ],
        destination=Destination(file_path="x.xlsx", sheet_name="Out", start_col="E", start_row="3"),
    )
    src = SourceConfig(path="src.xlsx", recipes=[RecipeConfig(name="R", sheets=[sh])])
    template = tpl.source_to_template(src)

    p = tmp_path / "full.json"
    tpl.save_template_json(template, str(p))
    loaded = tpl.load_template_json(str(p))

    target = SourceConfig(path="other.xlsx", recipes=[])
    tpl.apply_template_to_source(target, loaded)

    out_sh = target.recipes[0].sheets[0]
    assert out_sh.source_start_row == "2"
    assert out_sh.columns_spec == "A-D"
    assert out_sh.rows_spec == "5-10"
    assert out_sh.paste_mode == "keep"
    assert out_sh.rules_combine == "OR"
    assert len(out_sh.rules) == 2
    assert out_sh.rules[0].operator == ">"
    assert out_sh.rules[1].mode == "exclude"
    assert out_sh.destination.start_col == "E"
    assert out_sh.destination.start_row == "3"
