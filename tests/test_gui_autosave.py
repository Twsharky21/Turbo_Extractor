import json
import os


def test_gui_autosave_writes_project(tmp_path, monkeypatch):
    autosave_path = tmp_path / "autosave.json"
    monkeypatch.setenv("TURBO_AUTOSAVE_PATH", str(autosave_path))

    from gui.app import TurboExtractorApp

    app = TurboExtractorApp()
    try:
        # Mutate project (no dialogs) and force-save
        app.project.sources.append(
            __import__("core.project", fromlist=["SourceConfig"]).SourceConfig(path="C:/x.csv", recipes=[])
        )
        app._mark_dirty()
        app._autosave_now()

        assert autosave_path.exists()
        data = json.loads(autosave_path.read_text(encoding="utf-8"))
        assert data["sources"][0]["path"] == "C:/x.csv"
    finally:
        app.destroy()


def test_gui_autoload_on_start(tmp_path, monkeypatch):
    autosave_path = tmp_path / "autosave.json"
    monkeypatch.setenv("TURBO_AUTOSAVE_PATH", str(autosave_path))

    from core.project import ProjectConfig, SourceConfig
    from core.autosave import save_project_atomic

    proj = ProjectConfig(sources=[SourceConfig(path="C:/a.xlsx", recipes=[])])
    save_project_atomic(proj, str(autosave_path))

    from gui.app import TurboExtractorApp

    app = TurboExtractorApp()
    try:
        assert len(app.project.sources) == 1
        assert app.project.sources[0].path == "C:/a.xlsx"
    finally:
        app.destroy()
