from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Optional

from .project import ProjectConfig


ENV_AUTOSAVE_PATH = "TURBO_AUTOSAVE_PATH"


def resolve_autosave_path(project_root: Optional[str] = None) -> str:
    """Resolve autosave path.

    Priority:
    1) TURBO_AUTOSAVE_PATH env var (absolute or relative)
    2) User-home scoped default: ~/.turbo_extractor_v3/autosave.json
    """
    env = os.getenv(ENV_AUTOSAVE_PATH)
    if env:
        p = Path(env)
        if not p.is_absolute():
            base = Path(project_root) if project_root else Path.cwd()
            p = base / p
        return str(p)

    base = Path.home() / ".turbo_extractor_v3"
    return str(base / "autosave.json")


def atomic_write_text(path: str, text: str, encoding: str = "utf-8") -> None:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_suffix(p.suffix + ".tmp")
    tmp.write_text(text, encoding=encoding)
    os.replace(str(tmp), str(p))


def atomic_write_json(path: str, data) -> None:
    payload = json.dumps(data, indent=2)
    atomic_write_text(path, payload)


def save_project_atomic(project: ProjectConfig, path: str) -> None:
    atomic_write_json(path, project.to_dict())


def load_project_if_exists(path: str) -> Optional[ProjectConfig]:
    p = Path(path)
    if not p.exists():
        return None
    return ProjectConfig.load_json(str(p))
