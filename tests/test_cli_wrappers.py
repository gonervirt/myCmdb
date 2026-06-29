from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pandas as pd

from sample_data_generator import build_sample_workbooks


SCRIPT_TABLES = [
    "load_user.py",
    "load_localisation.py",
    "load_os.py",
    "load_vlan.py",
    "load_ip_address.py",
    "load_application.py",
    "load_team.py",
    "load_server.py",
]


def test_cli_wrapper_scripts(tmp_path: Path) -> None:
    repo_root = Path(__file__).resolve().parent.parent
    sample_dir = tmp_path / "sample_data"
    sample_dir.mkdir(parents=True, exist_ok=True)
    build_sample_workbooks(out_dir=sample_dir)

    config_path = tmp_path / "cmdb_config.json"
    config = {
        "sql_schema_path": str(repo_root / "db" / "cmdb_2026-06-28T15_49_08.585Z.sql"),
        "cmdb_model": str(sample_dir / "cmdb_model.xlsx"),
        "backup_dir": str(tmp_path / "backup"),
        "data_dir": str(sample_dir),
        "map_dir": str(sample_dir),
        "normalization_dir": str(sample_dir),
    }
    config_path.write_text(json.dumps(config), encoding="utf-8")

    expected_normalization_files = [
        sample_dir / "user_normalization.xlsx",
        sample_dir / "server_normalization.xlsx",
        sample_dir / "localisation_normalization.xlsx",
        sample_dir / "os_normalization.xlsx",
        sample_dir / "vlan_normalization.xlsx",
        sample_dir / "ip_address_normalization.xlsx",
        sample_dir / "application_normalization.xlsx",
        sample_dir / "team_normalization.xlsx",
    ]
    for normalization_file in expected_normalization_files:
        assert normalization_file.exists(), f"Missing normalization file: {normalization_file}"

    output_file = tmp_path / "cmdb_output.xlsx"
    output_file.write_text("", encoding="utf-8")

    for script_name in SCRIPT_TABLES:
        script_path = Path(__file__).resolve().parent.parent / "scripts" / script_name
        result = subprocess.run(
            [
                Path(".venv") / "Scripts" / "python.exe",
                script_path,
                "--config",
                str(config_path),
                "--output-file",
                str(output_file),
            ],
            cwd=Path(__file__).resolve().parent.parent,
            capture_output=True,
            text=True,
            check=False,
        )
        assert result.returncode == 0, f"{script_name} failed: {result.stderr}"

    workbook = pd.ExcelFile(output_file, engine="openpyxl")
    expected_sheets = {"Application", "IP address", "Localisation", "OS", "Server", "Team", "User", "VLAN"}
    assert set(workbook.sheet_names) == expected_sheets


def test_cli_wrapper_uses_default_config_file_when_not_provided(tmp_path: Path) -> None:
    repo_root = Path(__file__).resolve().parent.parent
    sample_dir = tmp_path / "sample_data"
    sample_dir.mkdir(parents=True, exist_ok=True)
    build_sample_workbooks(out_dir=sample_dir)

    config_path = tmp_path / "cmdb_config.json"
    config = {
        "sql_schema_path": str(repo_root / "db" / "cmdb_2026-06-28T15_49_08.585Z.sql"),
        "cmdb_model": str(tmp_path / "generated_model.xlsx"),
        "backup_dir": str(tmp_path / "backup"),
        "data_dir": str(sample_dir),
        "map_dir": str(sample_dir),
        "normalization_dir": str(sample_dir),
    }
    config_path.write_text(json.dumps(config), encoding="utf-8")

    assert (sample_dir / "server_normalization.xlsx").exists()

    script_path = repo_root / "scripts" / "load_server.py"
    result = subprocess.run(
        [sys.executable, str(script_path)],
        cwd=tmp_path,
        capture_output=True,
        text=True,
        check=False,
    )

    assert result.returncode == 0, result.stderr
    assert (tmp_path / "generated_model.xlsx").exists()
