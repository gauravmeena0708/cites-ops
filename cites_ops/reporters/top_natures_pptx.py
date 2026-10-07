"""Offline template-based slide export through the local Artifact Tool runtime."""

import json
import os
import shutil
import subprocess
from pathlib import Path


def write_presentations(data, output_dir, runtime=None):
    node = shutil.which("node")
    if not node:
        raise RuntimeError("PPT export needs Node.js and a local @oai/artifact-tool runtime. Use --no-pptx for HTML/Excel only.")
    runtime = Path(runtime or os.environ.get("CITES_ARTIFACT_WORKSPACE") or ".cache/presentations").resolve()
    output_dir = Path(output_dir).resolve()
    templates = Path(__file__).parents[1] / "templates"
    # This temporary payload includes only masked examples, never full ticket descriptions.
    payload = output_dir / "_internal" / "pptx_input.json"
    payload.parent.mkdir(parents=True, exist_ok=True)
    payload.write_text(json.dumps({k: data[k] for k in ("date", "scope", "week", "rankings", "count_source")}, ensure_ascii=False), encoding="utf-8")
    try:
        result = subprocess.run(
            [node, str(templates / "render_top_natures.mjs"), str(runtime), str(payload), str(output_dir)],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=600,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if result.returncode:
            raise RuntimeError("PPT export failed. Set --pptx-runtime to a local workspace containing @oai/artifact-tool, or use --no-pptx. " + (result.stderr or result.stdout)[-2000:])
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError("PPT export exceeded 10 minutes; existing reports were preserved.") from exc
    finally:
        payload.unlink(missing_ok=True)
