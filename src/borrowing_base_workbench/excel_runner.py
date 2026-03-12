from __future__ import annotations

import json
import subprocess
import tempfile
from pathlib import Path


def _run_json_script(script_path: str | Path, arguments: list[str]) -> dict:
    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(Path(script_path)),
    ] + arguments
    completed = subprocess.run(command, capture_output=True, text=True, check=False)
    stdout = completed.stdout.strip()
    if not stdout:
        return {"status": "error", "message": completed.stderr.strip() or "Excel probe returned no output."}
    try:
        return json.loads(stdout)
    except json.JSONDecodeError:
        start = stdout.find("{")
        end = stdout.rfind("}")
        if start != -1 and end != -1 and end > start:
            try:
                return json.loads(stdout[start : end + 1])
            except json.JSONDecodeError:
                pass
        return {"status": "error", "message": "Excel probe returned non-JSON output.", "stdout": stdout}


def probe_excel_workbook(workbook_path: str | Path, script_path: str | Path) -> dict:
    return _run_json_script(script_path, ["-WorkbookPath", str(Path(workbook_path))])


def run_pro_forma_workbook(workbook_path: str | Path, script_path: str | Path, scenario: dict) -> dict:
    with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as handle:
        json.dump(scenario, handle)
        scenario_path = handle.name
    return _run_json_script(
        script_path,
        [
            "-WorkbookPath",
            str(Path(workbook_path)),
            "-ScenarioPath",
            scenario_path,
        ],
    )
