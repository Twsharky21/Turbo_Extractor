"""Tests for core.project — serialization, JSON roundtrip, build_run_items, source_start_row persistence."""
import os
from tempfile import TemporaryDirectory

from core.project import ProjectConfig, SourceConfig, RecipeConfig
from core.models import SheetConfig, Destination, Rule


def _make_project(dest_path: str) -> ProjectConfig:
    sheet = SheetConfig(
        name="Sheet1", workbook_sheet="Sheet1",
        columns_spec="A,C", rows_spec="1-1",
        paste_mode="pack", rules_combine="AND",
        rules=[Rule(mode="include", column="A", operator="equals", value="alpha")],
        destination=Destination(file_path=dest_path, sheet_name="Out", start_col="B", start_row=""),
    )
    recipe = RecipeConfig(name="R1", sheets=[sheet])
    source = SourceConfig(path="dummy.xlsx", recipes=[recipe])
    return ProjectConfig(sources=[source])


def test_project_serialization_roundtrip():
    with TemporaryDirectory() as td:
        json_path = os.path.join(td, "proj.json")
        proj = _make_project("out.xlsx")
        proj.save_json(json_path)
        loaded = ProjectConfig.load_json(json_path)

        assert len(loaded.sources) == 1
        assert loaded.sources[0].recipes[0].name == "R1"
        assert loaded.sources[0].recipes[0].sheets[0].columns_spec == "A,C"
        assert loaded.sources[0].recipes[0].sheets[0].rules[0].value == "alpha"


def test_project_build_run_items_order():
    proj = _make_project("out.xlsx")
    items = proj.build_run_items()

    assert len(items) == 1
    src_path, recipe_name, sheet = items[0]
    assert src_path == "dummy.xlsx"
    assert recipe_name == "R1"
    assert sheet.name == "Sheet1"


def test_project_json_roundtrip_preserves_source_start_row():
    with TemporaryDirectory() as td:
        sh = SheetConfig(
            name="Sheet1", workbook_sheet="Sheet1",
            source_start_row="7", columns_spec="A", rows_spec="",
            paste_mode="pack", rules_combine="AND", rules=[],
            destination=Destination(file_path="out.xlsx", sheet_name="Dest", start_col="A", start_row=""),
        )
        proj = ProjectConfig(sources=[
            SourceConfig(path="/tmp/source.xlsx", recipes=[RecipeConfig(name="Recipe1", sheets=[sh])])
        ])

        p = os.path.join(td, "proj.json")
        proj.save_json(p)
        loaded = ProjectConfig.load_json(p)
        assert loaded.sources[0].recipes[0].sheets[0].source_start_row == "7"


def test_project_multiple_sources_recipes_sheets_roundtrip():
    with TemporaryDirectory() as td:
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

        p = os.path.join(td, "proj.json")
        proj.save_json(p)
        loaded = ProjectConfig.load_json(p)

        assert len(loaded.sources) == 2
        assert loaded.sources[0].recipes[0].sheets[1].name == "S2"
        assert loaded.sources[0].recipes[1].name == "R2"
        assert loaded.sources[1].path == "b.csv"

        items = loaded.build_run_items()
        assert len(items) == 4
        assert [i[1] for i in items] == ["R1", "R1", "R2", "R3"]
