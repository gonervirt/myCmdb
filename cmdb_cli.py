from __future__ import annotations

import argparse
import json
import re
import shutil
import sqlite3
from dataclasses import dataclass, field
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

import pandas as pd


SUPPORTED_TABLES = {
    "server": "Server",
    "load-server": "Server",
    "localisation": "Localisation",
    "load-localisation": "Localisation",
    "application": "Application",
    "load-application": "Application",
    "user": "User",
    "load-user": "User",
    "ip-address": "IP address",
    "load-ip-address": "IP address",
    "vlan": "VLAN",
    "load-vlan": "VLAN",
    "os": "OS",
    "load-os": "OS",
    "team": "Team",
    "load-team": "Team",
}


@dataclass(slots=True)
class CMDBConfig:
    sql_schema_path: Path
    cmdb_model: Path
    backup_dir: Path | None = None
    data_dir: Path | None = None
    map_dir: Path | None = None
    normalization_dir: Path | None = None
    tables: dict[str, dict[str, str]] = field(default_factory=dict)
    base_dir: Path | None = None

    @classmethod
    def default(cls) -> "CMDBConfig":
        base_dir = Path(__file__).resolve().parent
        return cls(
            sql_schema_path=base_dir / "db" / "cmdb_2026-06-28T15_49_08.585Z.sql",
            cmdb_model=base_dir / "sample_data" / "cmdb_model.xlsx",
            backup_dir=base_dir / "backup",
            data_dir=base_dir / "sample_data",
            map_dir=base_dir / "sample_data",
            normalization_dir=base_dir / "sample_data",
            tables={},
            base_dir=base_dir,
        )

    @classmethod
    def from_json(cls, config_path: Path) -> "CMDBConfig":
        raw_config = json.loads(config_path.read_text(encoding="utf-8"))
        base_dir = config_path.parent
        return cls(
            sql_schema_path=cls._resolve_path(raw_config.get("sql_schema_path"), base_dir)
            or base_dir / "db" / "cmdb_2026-06-28T15_49_08.585Z.sql",
            cmdb_model=cls._resolve_path(raw_config.get("cmdb_model"), base_dir)
            or base_dir / "sample_data" / "cmdb_model.xlsx",
            backup_dir=cls._resolve_path(raw_config.get("backup_dir"), base_dir),
            data_dir=cls._resolve_path(raw_config.get("data_dir"), base_dir),
            map_dir=cls._resolve_path(raw_config.get("map_dir"), base_dir),
            normalization_dir=cls._resolve_path(raw_config.get("normalization_dir"), base_dir),
            tables=raw_config.get("tables", {}),
            base_dir=base_dir,
        )

    @staticmethod
    def _resolve_path(value: Any, base_dir: Path) -> Path | None:
        if value is None:
            return None
        path = Path(str(value))
        return path if path.is_absolute() else base_dir / path

    def resolve_path(self, value: str | None) -> Path | None:
        if value is None:
            return None
        return self._resolve_path(value, self.base_dir or Path.cwd())

    def _table_config_for(self, command: str) -> dict[str, str]:
        if command in self.tables:
            return self.tables[command]
        alias = command.removeprefix("load-")
        return self.tables.get(alias, {})

    def resolve_data_file(self, command: str, explicit_path: str | None) -> Path:
        if explicit_path:
            return self.resolve_path(explicit_path)
        table_config = self._table_config_for(command)
        if "data_file" in table_config:
            return self._resolve_path(table_config["data_file"], self.base_dir or Path.cwd())
        if self.data_dir:
            command_name = command.removeprefix("load-").replace("-", "_")
            return self.data_dir / f"{command_name}_data.xlsx"
        raise ValueError(
            f"Data file for command '{command}' is missing. Provide --data-file or configure data_dir/table data_file."
        )

    def resolve_map_file(self, command: str, explicit_path: str | None) -> Path:
        if explicit_path:
            return self.resolve_path(explicit_path)
        table_config = self._table_config_for(command)
        if "map_file" in table_config:
            return self._resolve_path(table_config["map_file"], self.base_dir or Path.cwd())
        if self.map_dir:
            command_name = command.removeprefix("load-").replace("-", "_")
            return self.map_dir / f"{command_name}_mapping.xlsx"
        raise ValueError(
            f"Mapping file for command '{command}' is missing. Provide --map-file or configure map_dir/table map_file."
        )

    def resolve_normalization_file(self, command: str, explicit_path: str | None) -> Path:
        if explicit_path:
            return self.resolve_path(explicit_path)
        table_config = self._table_config_for(command)
        if "normalization_file" in table_config:
            return self._resolve_path(table_config["normalization_file"], self.base_dir or Path.cwd())
        if self.normalization_dir:
            command_name = command.removeprefix("load-").replace("-", "_")
            return self.normalization_dir / f"{command_name}_normalization.xlsx"
        raise ValueError(
            f"Normalization file for command '{command}' is missing. Provide --normalization-file or configure normalization_dir/table normalization_file."
        )

    def resolve_input_model(self, explicit_path: str | None) -> Path:
        if explicit_path:
            return self.resolve_path(explicit_path)
        return self.cmdb_model

    def resolve_output_file(self, explicit_path: str | None) -> Path:
        if explicit_path:
            return self.resolve_path(explicit_path)
        return self.cmdb_model


class CMDBSchema:
    """Loads the SQL schema and provides schema metadata for the in-memory database."""

    def __init__(self, sql_schema_path: Path) -> None:
        self.sql_schema_path = sql_schema_path
        self.table_names: set[str] = set()

    def create_database(self) -> sqlite3.Connection:
        connection = sqlite3.connect(":memory:")
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = OFF")
        schema_sql = self.sql_schema_path.read_text(encoding="utf-8")
        connection.executescript(schema_sql)
        self.table_names = self._discover_tables(connection)
        return connection

    @staticmethod
    def _discover_tables(connection: sqlite3.Connection) -> set[str]:
        cursor = connection.execute("SELECT name FROM sqlite_master WHERE type='table'")
        return {row[0] for row in cursor.fetchall()}

    def load_excel_data_model(self, connection: sqlite3.Connection, workbook_path: Path) -> None:
        workbook = pd.ExcelFile(workbook_path, engine="openpyxl")
        for sheet_name in workbook.sheet_names:
            if sheet_name not in self.table_names:
                raise ValueError(
                    f"Excel model sheet '{sheet_name}' does not match any known CMDB table."
                )
            frame = pd.read_excel(workbook_path, sheet_name=sheet_name, engine="openpyxl")
            self._append_dataframe_to_table(connection, sheet_name, frame)

    def _append_dataframe_to_table(
        self, connection: sqlite3.Connection, table_name: str, df: pd.DataFrame
    ) -> None:
        if df.empty:
            return
        df = df.rename(columns=str)
        df.to_sql(table_name, connection, if_exists="append", index=False)

    def export_db_to_excel(self, connection: sqlite3.Connection, output_path: Path) -> None:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        tables = sorted(self.table_names)
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for table_name in tables:
                frame = pd.read_sql_query(f'SELECT * FROM "{table_name}"', connection)
                frame.to_excel(writer, sheet_name=table_name, index=False)

    @staticmethod
    def get_table_columns(connection: sqlite3.Connection, table_name: str) -> list[str]:
        cursor = connection.execute(f'PRAGMA table_info("{table_name}")')
        return [row[1] for row in cursor.fetchall()]

    @staticmethod
    def get_table_primary_keys(connection: sqlite3.Connection, table_name: str) -> list[str]:
        cursor = connection.execute(f'PRAGMA table_info("{table_name}")')
        return [row[1] for row in cursor.fetchall() if row[5] > 0]

    @staticmethod
    def get_required_columns(connection: sqlite3.Connection, table_name: str) -> list[str]:
        cursor = connection.execute(f'PRAGMA table_info("{table_name}")')
        required: list[str] = []
        for row in cursor.fetchall():
            name = row[1]
            not_null = row[3] == 1
            default_value = row[4]
            if not_null and default_value is None:
                required.append(name)
        return required


@dataclass(slots=True)
class MappingRule:
    source_column: str
    target_column: str
    static_value: Any | None = None
    table: str | None = None


@dataclass(slots=True)
class MappingConfig:
    rules: list[MappingRule] = field(default_factory=list)

    @classmethod
    def from_excel(cls, workbook_path: Path, table_name: str) -> "MappingConfig":
        workbook = pd.ExcelFile(workbook_path, engine="openpyxl")
        if table_name in workbook.sheet_names:
            frame = pd.read_excel(workbook_path, sheet_name=table_name, engine="openpyxl")
        elif len(workbook.sheet_names) == 1:
            frame = pd.read_excel(workbook_path, sheet_name=workbook.sheet_names[0], engine="openpyxl")
        else:
            raise ValueError(
                f"Mapping workbook '{workbook_path}' must contain a sheet named '{table_name}' or exactly one sheet."
            )

        frame = frame.rename(columns=str)
        missing_columns = {"source_column", "target_column"} - set(frame.columns)
        if missing_columns:
            raise ValueError(
                f"Mapping file '{workbook_path}' is missing required columns: {', '.join(sorted(missing_columns))}."
            )

        rules: list[MappingRule] = []
        for _, row in frame.iterrows():
            table = row.get("table") if "table" in frame.columns else None
            source_value = row["source_column"]
            target_value = row["target_column"]
            if pd.isna(source_value) or pd.isna(target_value):
                continue
            source_column = str(source_value).strip()
            target_column = str(target_value).strip()
            if not source_column or not target_column:
                continue
            if source_column.lower() == "nan":
                continue

            static_value: Any | None = None
            if source_column.startswith("*"):
                static_value = source_column[1:].strip()
                source_column = ""

            if table is None or pd.isna(table) or str(table).strip() == table_name:
                rules.append(
                    MappingRule(
                        source_column=source_column,
                        target_column=target_column,
                        static_value=static_value,
                        table=str(table).strip() if pd.notna(table) else None,
                    )
                )
        return cls(rules=rules)

    def get_column_mapping(self) -> dict[str, str]:
        return {rule.source_column: rule.target_column for rule in self.rules}


@dataclass(slots=True)
class NormalizationRule:
    field: str
    source_value: Any
    target_value: Any
    table: str | None = None


@dataclass(slots=True)
class NormalizationConfig:
    rules: list[NormalizationRule] = field(default_factory=list)

    @classmethod
    def from_excel(cls, workbook_path: Path, table_name: str) -> "NormalizationConfig":
        workbook = pd.ExcelFile(workbook_path, engine="openpyxl")
        rules: list[NormalizationRule] = []

        if table_name in workbook.sheet_names:
            sheet_names = [table_name]
        elif len(workbook.sheet_names) == 1:
            sheet_names = [workbook.sheet_names[0]]
        else:
            sheet_names = workbook.sheet_names

        for sheet_name in sheet_names:
            frame = pd.read_excel(workbook_path, sheet_name=sheet_name, engine="openpyxl")
            frame = frame.rename(columns=str)
            if {"field", "source_value", "target_value"}.issubset(frame.columns):
                for _, row in frame.iterrows():
                    field = str(row["field"]).strip()
                    if not field:
                        continue
                    rules.append(
                        NormalizationRule(
                            field=field,
                            source_value=row["source_value"],
                            target_value=row["target_value"],
                            table=None,
                        )
                    )
                continue

            if {"source_value", "target_value"}.issubset(frame.columns):
                field_name = str(sheet_name).strip()
                if not field_name:
                    continue
                for _, row in frame.iterrows():
                    source_value = row["source_value"]
                    target_value = row["target_value"]
                    if pd.isna(source_value) and pd.isna(target_value):
                        continue
                    rules.append(
                        NormalizationRule(
                            field=field_name,
                            source_value=source_value,
                            target_value=target_value,
                            table=None,
                        )
                    )
                continue

            raise ValueError(
                f"Normalization sheet '{sheet_name}' in '{workbook_path}' must contain either 'field/source_value/target_value' columns or 'source_value/target_value' columns."
            )

        return cls(rules=rules)

    def apply(self, table_name: str, frame: pd.DataFrame) -> pd.DataFrame:
        normalized = frame.copy()
        for rule in self.rules:
            if rule.table and rule.table != table_name:
                continue
            if rule.field not in normalized.columns:
                raise ValueError(
                    f"Normalization rule references field '{rule.field}' for table '{table_name}', but that column is not present after mapping."
                )
            mask = normalized[rule.field] == rule.source_value
            if mask.any():
                normalized.loc[mask, rule.field] = rule.target_value
        return normalized


class TableImporter:
    """Applies mapping and normalization rules and writes rows to the target table."""

    def __init__(
        self,
        connection: sqlite3.Connection,
        schema: CMDBSchema,
        table_name: str,
        mapping_config: MappingConfig,
        normalization_config: NormalizationConfig,
    ) -> None:
        self.connection = connection
        self.schema = schema
        self.table_name = table_name
        self.mapping_config = mapping_config
        self.normalization_config = normalization_config
        self.table_columns = self.schema.get_table_columns(connection, table_name)
        self.required_columns = self.schema.get_required_columns(connection, table_name)
        self.primary_keys = self.schema.get_table_primary_keys(connection, table_name)

    def import_rows(self, source_frame: pd.DataFrame) -> None:
        framed = self._prepare_dataframe(source_frame)
        self._write_rows(framed)

    def _prepare_dataframe(self, source_frame: pd.DataFrame) -> pd.DataFrame:
        frame = source_frame.rename(columns=str).copy()
        if frame.empty:
            return pd.DataFrame(columns=self.table_columns)

        valid_rules = [rule for rule in self.mapping_config.rules if rule.target_column in self.table_columns]
        source_rules = [rule for rule in valid_rules if rule.source_column]
        missing_source_columns = [rule.source_column for rule in source_rules if rule.source_column not in frame.columns]
        if missing_source_columns:
            raise ValueError(
                "Mapping file references source columns that are missing from the imported data: "
                f"{', '.join(missing_source_columns)}"
            )

        mapped_frame = pd.DataFrame(index=frame.index)
        for rule in valid_rules:
            if rule.static_value is not None:
                mapped_frame[rule.target_column] = rule.static_value
            else:
                mapped_frame[rule.target_column] = frame[rule.source_column]

        unknown_target_columns = [column for column in mapped_frame.columns if column not in self.table_columns]
        if unknown_target_columns:
            raise ValueError(
                f"Mapped data contains columns not present in target table '{self.table_name}': "
                f"{', '.join(unknown_target_columns)}"
            )

        for column in mapped_frame.columns:
            if column not in self.table_columns:
                continue
            mapped_frame[column] = mapped_frame[column].apply(lambda value: self._normalize_mapped_value(column, value))

        normalized_frame = self.normalization_config.apply(self.table_name, mapped_frame)
        self._validate_required_columns(normalized_frame)
        return normalized_frame.reindex(columns=self.table_columns).copy()

    def _normalize_mapped_value(self, target_column: str, value: Any) -> Any:
        if pd.isna(value):
            return None

        if isinstance(value, str):
            stripped = value.strip()
            if stripped == "":
                return self._resolve_default_value(target_column)
            if stripped.startswith("*"):
                return stripped[1:].strip()
            return stripped

        return value

    def _resolve_default_value(self, target_column: str) -> Any:
        cursor = self.connection.execute(f'PRAGMA table_info("{self.table_name}")')
        for row in cursor.fetchall():
            if row[1] == target_column:
                default_value = row[4]
                if default_value is None:
                    return None
                if isinstance(default_value, str):
                    value = default_value.strip()
                    if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
                        return value[1:-1]
                    return value
                return default_value
        return None

    def _validate_required_columns(self, frame: pd.DataFrame) -> None:
        missing_required_columns = [column for column in self.required_columns if column not in frame.columns]
        generated_columns = [column for column in missing_required_columns if column in self.primary_keys]
        for column in generated_columns:
            frame[column] = frame.index.map(lambda row_index: self._generate_value_for_column(frame, column, row_index))

        for column in self.primary_keys:
            if column not in frame.columns:
                continue
            missing_values = frame[column].isna() | frame[column].eq("")
            if missing_values.any():
                frame.loc[missing_values, column] = frame.index[missing_values].map(
                    lambda row_index: self._generate_value_for_column(frame, column, row_index)
                )

        missing_required_columns = [column for column in self.required_columns if column not in frame.columns]
        if missing_required_columns:
            raise ValueError(
                f"Target table '{self.table_name}' requires columns that are not provided by mapping: "
                f"{', '.join(missing_required_columns)}"
            )

        for column in self.required_columns:
            if column not in frame.columns:
                continue
            missing_values = frame[column].isna() | frame[column].eq("")
            if missing_values.any():
                offending_rows = [int(index) + 2 for index, missing in enumerate(missing_values) if missing]
                raise ValueError(
                    f"Required column '{column}' has empty values in rows {', '.join(str(row_number) for row_number in offending_rows)}."
                )

    def _generate_value_for_column(self, frame: pd.DataFrame, column: str, row_index: int) -> Any:
        if column != "id":
            return None
        if self.table_name == "Application":
            cursor = self.connection.execute(f'SELECT MAX(CAST("id" AS INTEGER)) FROM "{self.table_name}"')
            max_id = cursor.fetchone()[0]
            next_id = (int(max_id) if max_id is not None else 0) + 1 + row_index
            return next_id

        candidate_columns = [candidate for candidate in frame.columns if candidate != "id"]
        for candidate in candidate_columns:
            value = frame.at[row_index, candidate]
            if pd.notna(value) and str(value).strip():
                slug = re.sub(r"[^A-Za-z0-9]+", "-", str(value)).strip("-").lower()
                return f"{self.table_name.lower().replace(' ', '-')}-{slug}"
        return f"{self.table_name.lower().replace(' ', '-')}-{row_index + 1}"

    def _write_rows(self, frame: pd.DataFrame) -> None:
        if frame.empty:
            return

        frame = frame.copy()
        for column in frame.columns:
            if column in frame.columns and frame[column].dtype == "object":
                frame[column] = frame[column].replace({pd.NA: None})

        if not self.primary_keys:
            frame.to_sql(self.table_name, self.connection, if_exists="append", index=False)
            self.connection.commit()
            return

        columns = [column for column in frame.columns if column in self.table_columns]
        for _, row in frame.iterrows():
            self._upsert_row(columns, row)
        self.connection.commit()

    def _upsert_row(self, columns: list[str], row: pd.Series) -> None:
        primary_key_values = [self._normalize_value(row[column]) for column in self.primary_keys if column in row.index]
        if not primary_key_values:
            self._insert_row(columns, row)
            return

        select_sql = (
            f'SELECT * FROM "{self.table_name}" WHERE '
            + " AND ".join(f'"{column}" = ?' for column in self.primary_keys if column in row.index)
        )
        cursor = self.connection.execute(select_sql, primary_key_values)
        existing_row = cursor.fetchone()
        if existing_row is None:
            self._insert_row(columns, row)
            return

        existing_values = dict(zip([col[0] for col in cursor.description], existing_row)) if cursor.description else {}
        update_assignments: list[str] = []
        update_values: list[Any] = []
        for column in columns:
            if column in self.primary_keys:
                continue
            incoming_value = self._normalize_value(row[column])
            existing_value = existing_values.get(column)
            merged_value = self._merge_column_value(existing_value, incoming_value)
            update_assignments.append(f'"{column}" = ?')
            update_values.append(merged_value)

        if not update_assignments:
            return

        where_clause = " AND ".join(f'"{column}" = ?' for column in self.primary_keys if column in row.index)
        update_sql = f'UPDATE "{self.table_name}" SET {", ".join(update_assignments)} WHERE {where_clause}'
        self.connection.execute(update_sql, [*update_values, *primary_key_values])

    def _insert_row(self, columns: list[str], row: pd.Series) -> None:
        row_columns = [column for column in columns if column in row.index]
        values = [self._normalize_value(row[column]) for column in row_columns]
        quoted_columns = ", ".join(f'"{column}"' for column in row_columns)
        placeholders = ", ".join("?" for _ in row_columns)
        self.connection.execute(
            f'INSERT INTO "{self.table_name}" ({quoted_columns}) VALUES ({placeholders})',
            values,
        )

    def _normalize_value(self, value: Any) -> Any:
        if pd.isna(value):
            return None
        if isinstance(value, str):
            stripped = value.strip()
            if stripped == "":
                return None
            return stripped
        return value

    def _merge_column_value(self, existing_value: Any, incoming_value: Any) -> Any:
        if self._is_empty_value(incoming_value):
            return existing_value
        if self._is_empty_value(existing_value):
            return incoming_value
        return incoming_value

    @staticmethod
    def _is_empty_value(value: Any) -> bool:
        if value is None:
            return True
        if pd.isna(value):
            return True
        if isinstance(value, str):
            return value.strip() == ""
        return False


class CMDBCLI:
    def __init__(self, sql_schema_path: Path) -> None:
        self.sql_schema_path = sql_schema_path
        self.schema = CMDBSchema(sql_schema_path)

    def build_parser(self) -> argparse.ArgumentParser:
        parser = argparse.ArgumentParser(
            description="CMDB ETL CLI: import Excel data into an in-memory CMDB and export a reusable Excel workbook."
        )
        parser.add_argument(
            "--config",
            required=False,
            help="Optional JSON config file defining default paths and directories",
        )
        subparsers = parser.add_subparsers(dest="command", required=True)

        export_parser = subparsers.add_parser("export-server-inventory", help="Export a flattened server inventory workbook")
        export_parser.add_argument("--input-model", required=True, help="Excel workbook representing the current CMDB data model")
        export_parser.add_argument("--output-file", required=True, help="Path to the Excel workbook that will receive the inventory export")

        for subcommand, table_name in SUPPORTED_TABLES.items():
            command_parser = subparsers.add_parser(subcommand, help=f"Load data into the {table_name} table")
            command_parser.add_argument("--data-file", required=False, help="Excel file containing rows to import")
            command_parser.add_argument("--map-file", required=False, help="Excel file defining source-to-target column mapping")
            command_parser.add_argument(
                "--normalization-file",
                required=False,
                help="Excel file defining normalization rules for target values",
            )
            command_parser.add_argument(
                "--input-model",
                required=False,
                help="Excel workbook representing the current CMDB data model",
            )
            command_parser.add_argument(
                "--output-file",
                required=False,
                help="Path to the Excel workbook that will receive the updated CMDB state",
            )

        return parser

    def run(self, argv: list[str] | None = None) -> None:
        parser = self.build_parser()
        args = parser.parse_args(argv)
        config = CMDBConfig.from_json(Path(args.config)) if args.config else CMDBConfig.default()
        if args.command == "export-server-inventory":
            self.export_server_inventory(Path(args.input_model), Path(args.output_file), config)
            return
        if args.command not in SUPPORTED_TABLES:
            raise ValueError(f"Unknown command: {args.command}")
        table_name = SUPPORTED_TABLES[args.command]
        self._execute_command(table_name, args, config)

    def _backup_existing_output(self, output_path: Path, backup_dir: Path | None) -> None:
        if not output_path.exists():
            return

        backup_root = backup_dir or output_path.parent
        backup_root.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now(UTC).strftime("%Y%m%dT%H%M%SZ")
        backup_path = backup_root / f"{output_path.stem}_{timestamp}{output_path.suffix}"
        shutil.copy2(output_path, backup_path)
        print(f"Backed up existing output file to {backup_path}")

    def _backup_database_dump(self, connection: sqlite3.Connection, output_path: Path, backup_dir: Path | None) -> Path:
        backup_root = backup_dir or output_path.parent
        backup_root.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now(UTC).strftime("%Y%m%dT%H%M%SZ")
        backup_path = backup_root / f"{output_path.stem}_{timestamp}.sqlite"
        dump_connection = sqlite3.connect(backup_path)
        try:
            connection.backup(dump_connection)
        finally:
            dump_connection.close()
        print(f"Backed up SQLite database to {backup_path}")
        return backup_path

    @staticmethod
    def _flatten_field_prefix(column_name: str) -> str:
        cleaned = str(column_name)
        for suffix in ("_id", "_ID"):
            if cleaned.endswith(suffix):
                cleaned = cleaned[: -len(suffix)]
                break
        return f"{cleaned.replace(' ', '_')}_"

    @staticmethod
    def _derive_table_name(column_name: str) -> str | None:
        cleaned = str(column_name).strip()
        for suffix in ("_id", "_ID"):
            if cleaned.endswith(suffix):
                cleaned = cleaned[: -len(suffix)]
                break
        if not cleaned:
            return None
        lowered = cleaned.lower()
        if lowered.startswith("ip") or lowered.endswith("address"):
            return "IP address"
        if lowered.startswith("local") or lowered.endswith("location"):
            return "Localisation"
        if lowered.startswith("owner") or lowered.startswith("referant") or lowered.startswith("support"):
            return "User"
        if lowered.startswith("os"):
            return "OS"
        if lowered.startswith("vlan"):
            return "VLAN"
        return None

    def _collect_related_rows(
        self,
        connection: sqlite3.Connection,
        table_name: str,
        row: sqlite3.Row,
        prefix: str = "",
        visited: set[str] | None = None,
    ) -> dict[str, Any]:
        visited = visited or set()
        table_key = table_name
        if table_key in visited:
            return {}
        visited = set(visited)
        visited.add(table_key)

        row_data: dict[str, Any] = {}
        for key in row.keys():
            row_data[f"{prefix}{key}"] = row[key]

        handled_columns: set[str] = set()
        cursor = connection.execute(f'PRAGMA foreign_key_list("{table_name}")')
        for foreign_key in cursor.fetchall():
            fk_table = foreign_key[2]
            fk_column = foreign_key[3]
            fk_to_column = foreign_key[4]
            if not fk_column or not fk_to_column:
                continue
            handled_columns.add(fk_column)
            fk_value = row[fk_column] if fk_column in row.keys() else None
            if fk_value is None or fk_value == "":
                continue
            related_row = connection.execute(
                f'SELECT * FROM "{fk_table}" WHERE "{fk_to_column}" = ?', (fk_value,)
            ).fetchone()
            if related_row is None:
                continue
            child_prefix = f"{prefix}{self._flatten_field_prefix(fk_column)}"
            related_data = self._collect_related_rows(connection, fk_table, related_row, prefix=child_prefix, visited=visited)
            for key, value in related_data.items():
                row_data[key] = value

        for column_name in row.keys():
            if column_name in handled_columns or not column_name.endswith(("_id", "_ID")):
                continue
            candidate_table = self._derive_table_name(column_name)
            if candidate_table is None or candidate_table not in self.schema.table_names or candidate_table == table_name:
                continue
            fk_value = row[column_name]
            if fk_value is None or fk_value == "":
                continue
            related_row = connection.execute(
                f'SELECT * FROM "{candidate_table}" WHERE "id" = ?', (fk_value,)
            ).fetchone()
            if related_row is None:
                continue
            child_prefix = f"{prefix}{self._flatten_field_prefix(column_name)}"
            related_data = self._collect_related_rows(connection, candidate_table, related_row, prefix=child_prefix, visited=visited)
            for key, value in related_data.items():
                row_data[key] = value

        if table_name == "VLAN" and "router" in row.keys():
            router_value = row["router"]
            if router_value not in (None, ""):
                router_row = connection.execute('SELECT * FROM "Server" WHERE "id" = ?', (router_value,)).fetchone()
                if router_row is not None:
                    child_prefix = f"{prefix}Router_"
                    for key in router_row.keys():
                        row_data[f"{child_prefix}{key}"] = router_row[key]

        if table_name == "VLAN" and "id" in row.keys() and row["id"] not in (None, ""):
            vlan_id = row["id"]
            related_row = connection.execute('SELECT * FROM "IP address" WHERE "VLAN_id" = ?', (vlan_id,)).fetchone()
            if related_row is not None:
                child_prefix = f"{prefix}IP_Address_"
                for key in related_row.keys():
                    row_data[f"{child_prefix}{key}"] = related_row[key]

        return row_data

    def export_server_inventory(self, input_model_path: Path, output_path: Path, config: CMDBConfig | None = None) -> None:
        config = config or CMDBConfig.default()
        connection = self.schema.create_database()
        if input_model_path.exists():
            self.schema.load_excel_data_model(connection, input_model_path)
        else:
            raise FileNotFoundError(f"Input model file not found: {input_model_path}")

        server_rows = connection.execute('SELECT * FROM "Server"').fetchall()
        expanded_rows: list[dict[str, Any]] = []
        for row in server_rows:
            expanded = self._collect_related_rows(connection, "Server", row)
            expanded_rows.append(expanded)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        frame = pd.DataFrame(expanded_rows)
        frame.to_excel(output_path, index=False)
        print(f"Wrote server inventory workbook to {output_path}")

    def _execute_command(self, table_name: str, args: argparse.Namespace, config: CMDBConfig) -> None:
        connection = self.schema.create_database()
        input_model_path = config.resolve_input_model(args.input_model)
        if input_model_path and input_model_path.exists():
            self.schema.load_excel_data_model(connection, input_model_path)
        elif args.input_model:
            raise FileNotFoundError(f"Input model file not found: {input_model_path}")

        mapping_config = MappingConfig.from_excel(config.resolve_map_file(args.command, args.map_file), table_name)
        normalization_config = NormalizationConfig.from_excel(
            config.resolve_normalization_file(args.command, args.normalization_file), table_name
        )
        source_frame = pd.read_excel(config.resolve_data_file(args.command, args.data_file), engine="openpyxl")

        importer = TableImporter(
            connection=connection,
            schema=self.schema,
            table_name=table_name,
            mapping_config=mapping_config,
            normalization_config=normalization_config,
        )
        importer.import_rows(source_frame)

        output_path = config.resolve_output_file(args.output_file)
        self._backup_existing_output(output_path, config.backup_dir)
        self._backup_database_dump(connection, output_path, config.backup_dir)
        self.schema.export_db_to_excel(connection, output_path)
        print(f"Wrote updated CMDB workbook to {output_path}")


def main() -> None:
    sql_schema_path = Path(__file__).resolve().parent / "db" / "cmdb_2026-06-28T15_49_08.585Z.sql"
    cli = CMDBCLI(sql_schema_path)
    cli.run()


if __name__ == "__main__":
    main()
