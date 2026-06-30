import sqlite3
from pathlib import Path

import pandas as pd

from cmdb_cli import (
    CMDBCLI,
    CMDBConfig,
    CMDBSchema,
    MappingConfig,
    MappingRule,
    NormalizationConfig,
    TableImporter,
)


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


def test_mapping_config_supports_static_source_column_marker(tmp_path: Path) -> None:
    workbook_path = tmp_path / "localisation_mapping.xlsx"
    frame = pd.DataFrame(
        [
            {"source_column": "id", "target_column": "id"},
            {"source_column": "source_country", "target_column": "Country"},
            {"source_column": "*toto", "target_column": "City"},
            {"source_column": "source_room", "target_column": "Room"},
        ]
    )
    frame.to_excel(workbook_path, sheet_name="Localisation", index=False)

    config = MappingConfig.from_excel(workbook_path, "Localisation")
    city_rule = next(rule for rule in config.rules if rule.target_column == "City")

    assert city_rule.source_column == ""
    assert city_rule.static_value == "toto"


def test_table_importer_uses_schema_defaults_and_star_prefix_values() -> None:
    connection = sqlite3.connect(":memory:")
    connection.execute('CREATE TABLE "TestTable" (id INTEGER PRIMARY KEY, name TEXT DEFAULT "fallback", flag TEXT)')

    schema = CMDBSchema(Path("db/cmdb_2026-06-28T15_49_08.585Z.sql"))
    mapping_config = MappingConfig(
        rules=[
            MappingRule(source_column="source_name", target_column="name"),
            MappingRule(source_column="source_flag", target_column="flag"),
        ]
    )
    normalization_config = NormalizationConfig(rules=[])
    importer = TableImporter(connection, schema, "TestTable", mapping_config, normalization_config)

    source_frame = pd.DataFrame([{"source_name": "", "source_flag": "*override"}])
    importer.import_rows(source_frame)

    row = connection.execute('SELECT name, flag FROM "TestTable"').fetchone()
    assert row[0] == "fallback"
    assert row[1] == "override"


def test_table_importer_uses_static_mapping_values_from_source_marker() -> None:
    connection = sqlite3.connect(":memory:")
    connection.execute('CREATE TABLE "Localisation" (id TEXT PRIMARY KEY, Country TEXT, City TEXT, Room TEXT)')

    schema = CMDBSchema(Path("db/cmdb_2026-06-28T15_49_08.585Z.sql"))
    mapping_config = MappingConfig(
        rules=[
            MappingRule(source_column="source_id", target_column="id"),
            MappingRule(source_column="source_country", target_column="Country"),
            MappingRule(source_column="", target_column="City", static_value="toto"),
            MappingRule(source_column="source_room", target_column="Room"),
        ]
    )
    normalization_config = NormalizationConfig(rules=[])
    importer = TableImporter(connection, schema, "Localisation", mapping_config, normalization_config)

    source_frame = pd.DataFrame([{"source_id": "loc-001", "source_country": "FR", "source_room": "HQ"}])
    importer.import_rows(source_frame)

    row = connection.execute('SELECT id, Country, City, Room FROM "Localisation"').fetchone()
    assert row == ("loc-001", "FR", "toto", "HQ")


def test_table_importer_ignores_mappings_to_unknown_target_columns() -> None:
    connection = sqlite3.connect(":memory:")
    connection.execute('CREATE TABLE "Localisation" (id TEXT PRIMARY KEY, Country TEXT, City TEXT, Room TEXT)')

    schema = CMDBSchema(Path("db/cmdb_2026-06-28T15_49_08.585Z.sql"))
    mapping_config = MappingConfig(
        rules=[
            MappingRule(source_column="source_id", target_column="id"),
            MappingRule(source_column="source_city", target_column="source_city"),
            MappingRule(source_column="source_room", target_column="source_room"),
        ]
    )
    normalization_config = NormalizationConfig(rules=[])
    importer = TableImporter(connection, schema, "Localisation", mapping_config, normalization_config)

    source_frame = pd.DataFrame([{"source_id": "loc-001", "source_city": "Paris", "source_room": "HQ"}])
    importer.import_rows(source_frame)

    row = connection.execute('SELECT id FROM "Localisation"').fetchone()
    assert row[0] == "loc-001"


def test_table_importer_updates_existing_primary_key_rows_with_merge_rules() -> None:
    connection = sqlite3.connect(":memory:")
    connection.execute('CREATE TABLE "TestTable" (id INTEGER PRIMARY KEY, name TEXT, email TEXT)')
    connection.execute('INSERT INTO "TestTable" (id, name, email) VALUES (?, ?, ?)', (1, "db-name", ""))
    connection.execute('INSERT INTO "TestTable" (id, name, email) VALUES (?, ?, ?)', (2, "", ""))
    connection.commit()

    schema = CMDBSchema(Path("db/cmdb_2026-06-28T15_49_08.585Z.sql"))
    mapping_config = MappingConfig(
        rules=[
            MappingRule(source_column="source_id", target_column="id"),
            MappingRule(source_column="source_name", target_column="name"),
            MappingRule(source_column="source_email", target_column="email"),
        ]
    )
    normalization_config = NormalizationConfig(rules=[])
    importer = TableImporter(connection, schema, "TestTable", mapping_config, normalization_config)

    source_frame = pd.DataFrame(
        [
            {"source_id": 1, "source_name": "import-name", "source_email": ""},
            {"source_id": 2, "source_name": "", "source_email": "import-email"},
        ]
    )
    importer.import_rows(source_frame)

    row_one = connection.execute('SELECT name, email FROM "TestTable" WHERE id = 1').fetchone()
    row_two = connection.execute('SELECT name, email FROM "TestTable" WHERE id = 2').fetchone()

    assert row_one == ("import-name", "")
    assert row_two == ("", "import-email")


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
