from pathlib import Path

import pandas as pd

from cmdb_cli import CMDBCLI, CMDBSchema


def test_schema_and_export_round_trip(tmp_path: Path) -> None:
    schema_path = Path("db/cmdb_2026-06-28T15_49_08.585Z.sql")
    cli = CMDBCLI(schema_path)
    conn = cli.schema.create_database()
    assert set(cli.schema.table_names) == {
        "Server",
        "Localisation",
        "Application",
        "User",
        "IP address",
        "VLAN",
        "OS",
        "Team",
    }

    output_file = tmp_path / "cmdb_output.xlsx"
    cli.schema.export_db_to_excel(conn, output_file)
    assert output_file.exists()

    workbook = pd.ExcelFile(output_file, engine="openpyxl")
    assert set(workbook.sheet_names) == {
        "Application",
        "IP address",
        "Localisation",
        "OS",
        "Server",
        "Team",
        "User",
        "VLAN",
    }
