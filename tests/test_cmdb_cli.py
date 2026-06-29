import sqlite3
from pathlib import Path

import pandas as pd

from cmdb_cli import CMDBCLI, CMDBConfig, CMDBSchema, NormalizationConfig


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


def test_normalization_config_supports_one_sheet_per_field(tmp_path: Path) -> None:
    workbook_path = tmp_path / "server_normalization.xlsx"
    with pd.ExcelWriter(workbook_path, engine="openpyxl") as writer:
        pd.DataFrame([{"source_value": "user-001", "target_value": "user-001"}]).to_excel(
            writer, sheet_name="Owner_id", index=False
        )
        pd.DataFrame([{"source_value": "os-001", "target_value": "os-001"}]).to_excel(
            writer, sheet_name="OS_id", index=False
        )

    config = NormalizationConfig.from_excel(workbook_path, "Server")
    frame = pd.DataFrame([{"Owner_id": "user-001", "OS_id": "os-001"}])
    normalized = config.apply("Server", frame)

    assert normalized.loc[0, "Owner_id"] == "user-001"
    assert normalized.loc[0, "OS_id"] == "os-001"


def test_database_backup_dump_is_created_on_save(tmp_path: Path) -> None:
    cli = CMDBCLI(Path("db/cmdb_2026-06-28T15_49_08.585Z.sql"))
    connection = cli.schema.create_database()
    connection.execute('CREATE TABLE "TestTable" (id INTEGER PRIMARY KEY, name TEXT)')
    connection.execute('INSERT INTO "TestTable" (name) VALUES (?)', ("demo",))
    connection.commit()

    backup_dir = tmp_path / "backup"
    output_path = tmp_path / "cmdb_model.xlsx"
    backup_path = cli._backup_database_dump(connection, output_path, backup_dir)

    assert backup_path.exists()
    assert backup_path.parent == backup_dir
    assert backup_path.name.startswith("cmdb_model_")
    assert backup_path.suffix == ".sqlite"

    backup_connection = sqlite3.connect(backup_path)
    try:
        rows = backup_connection.execute('SELECT name FROM "TestTable"').fetchall()
    finally:
        backup_connection.close()

    assert rows == [("demo",)]


def test_server_inventory_export_includes_related_foreign_key_data(tmp_path: Path) -> None:
    from sample_data_generator import build_sample_workbooks

    build_sample_workbooks(out_dir=tmp_path)

    cli = CMDBCLI(Path("db/cmdb_2026-06-28T15_49_08.585Z.sql"))
    output_path = tmp_path / "server_inventory.xlsx"
    cli.export_server_inventory(tmp_path / "cmdb_model.xlsx", output_path)

    assert output_path.exists()
    frame = pd.read_excel(output_path, engine="openpyxl")
    assert "Name" in frame.columns
    assert "Owner_Name" in frame.columns
    assert "IP_Address_IPv4" in frame.columns
    assert "IP_Address_VLAN_Name" in frame.columns
    assert "IP_Address_VLAN_Router_Name" in frame.columns


def test_generated_normalization_workbooks_include_all_field_sheets(tmp_path: Path) -> None:
    from sample_data_generator import build_sample_workbooks

    build_sample_workbooks(out_dir=tmp_path)

    expected_sheets = {
        "user_normalization.xlsx": ["id", "Name", "email"],
        "server_normalization.xlsx": ["id", "Name", "Owner_id", "Localisation_id", "OS_id", "IP_Address_id"],
        "localisation_normalization.xlsx": ["id", "Country", "City", "Room"],
        "os_normalization.xlsx": ["id", "Name", "Version", "EDR", "type"],
        "vlan_normalization.xlsx": ["router", "Gateway", "id", "Name", "CIDR", "Advertised"],
        "ip_address_normalization.xlsx": ["id", "IPv4", "VLAN_id", "IP public"],
        "application_normalization.xlsx": ["id", "Owner", "support", "Name", "Description", "Criticality", "Hosted"],
        "team_normalization.xlsx": ["id", "user"],
    }

    for workbook_name, expected in expected_sheets.items():
        with pd.ExcelFile(tmp_path / workbook_name, engine="openpyxl") as workbook:
            assert set(workbook.sheet_names) == set(expected)
