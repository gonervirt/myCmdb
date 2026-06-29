from __future__ import annotations

import argparse
import re
import sqlite3
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd


SUPPORTED_TABLES = {
    "load-server": "Server",
    "load-localisation": "Localisation",
    "load-application": "Application",
    "load-user": "User",
    "load-ip-address": "IP address",
    "load-vlan": "VLAN",
    "load-os": "OS",
    "load-team": "Team",
}


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
            source_column = str(row["source_column"]).strip()
            target_column = str(row["target_column"]).strip()
            if not source_column or not target_column:
                continue
            if table is None or pd.isna(table) or str(table).strip() == table_name:
                rules.append(
                    MappingRule(
                        source_column=source_column,
                        target_column=target_column,
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
        if table_name in workbook.sheet_names:
            frame = pd.read_excel(workbook_path, sheet_name=table_name, engine="openpyxl")
        elif len(workbook.sheet_names) == 1:
            frame = pd.read_excel(workbook_path, sheet_name=workbook.sheet_names[0], engine="openpyxl")
        else:
            raise ValueError(
                f"Normalization workbook '{workbook_path}' must contain a sheet named '{table_name}' or exactly one sheet."
            )

        frame = frame.rename(columns=str)
        missing_columns = {"field", "source_value", "target_value"} - set(frame.columns)
        if missing_columns:
            raise ValueError(
                f"Normalization file '{workbook_path}' is missing required columns: {', '.join(sorted(missing_columns))}."
            )

        rules: list[NormalizationRule] = []
        for _, row in frame.iterrows():
            table = row.get("table") if "table" in frame.columns else None
            field = str(row["field"]).strip()
            if not field:
                continue
            if table is None or pd.isna(table) or str(table).strip() == table_name:
                rules.append(
                    NormalizationRule(
                        field=field,
                        source_value=row["source_value"],
                        target_value=row["target_value"],
                        table=str(table).strip() if pd.notna(table) else None,
                    )
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

        column_map = self.mapping_config.get_column_mapping()
        missing_source_columns = [source_column for source_column in column_map if source_column not in frame.columns]
        if missing_source_columns:
            raise ValueError(
                "Mapping file references source columns that are missing from the imported data: "
                f"{', '.join(missing_source_columns)}"
            )

        mapped_frame = frame.rename(columns=column_map)
        unknown_target_columns = [column for column in mapped_frame.columns if column not in self.table_columns]
        if unknown_target_columns:
            raise ValueError(
                f"Mapped data contains columns not present in target table '{self.table_name}': "
                f"{', '.join(unknown_target_columns)}"
            )

        normalized_frame = self.normalization_config.apply(self.table_name, mapped_frame)
        self._validate_required_columns(normalized_frame)
        return normalized_frame.reindex(columns=self.table_columns).copy()

    def _validate_required_columns(self, frame: pd.DataFrame) -> None:
        missing_required_columns = [column for column in self.required_columns if column not in frame.columns]
        generated_columns = [column for column in missing_required_columns if column in self.primary_keys]
        for column in generated_columns:
            frame[column] = frame.index.map(lambda row_index: self._generate_value_for_column(frame, column, row_index))

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
        placeholders = ", ".join("?" for _ in columns)
        quoted_columns = ", ".join(f'"{column}"' for column in columns)
        sql = (
            f'INSERT OR REPLACE INTO "{self.table_name}" ({quoted_columns}) VALUES ({placeholders})'
        )
        rows = [tuple(row[column] for column in columns) for _, row in frame.iterrows()]
        self.connection.executemany(sql, rows)
        self.connection.commit()


class CMDBCLI:
    def __init__(self, sql_schema_path: Path) -> None:
        self.sql_schema_path = sql_schema_path
        self.schema = CMDBSchema(sql_schema_path)

    def build_parser(self) -> argparse.ArgumentParser:
        parser = argparse.ArgumentParser(
            description="CMDB ETL CLI: import Excel data into an in-memory CMDB and export a reusable Excel workbook."
        )
        subparsers = parser.add_subparsers(dest="command", required=True)

        for subcommand, table_name in SUPPORTED_TABLES.items():
            command_parser = subparsers.add_parser(subcommand, help=f"Load data into the {table_name} table")
            command_parser.add_argument("--data-file", required=True, help="Excel file containing rows to import")
            command_parser.add_argument("--map-file", required=True, help="Excel file defining source-to-target column mapping")
            command_parser.add_argument(
                "--normalization-file",
                required=True,
                help="Excel file defining normalization rules for target values",
            )
            command_parser.add_argument(
                "--input-model",
                required=False,
                help="Optional existing Excel workbook representing the current CMDB data model",
            )
            command_parser.add_argument(
                "--output-file",
                required=True,
                help="Path to the Excel workbook that will receive the updated CMDB state",
            )

        return parser

    def run(self, argv: list[str] | None = None) -> None:
        parser = self.build_parser()
        args = parser.parse_args(argv)
        if args.command not in SUPPORTED_TABLES:
            raise ValueError(f"Unknown command: {args.command}")
        table_name = SUPPORTED_TABLES[args.command]
        self._execute_command(table_name, args)

    def _execute_command(self, table_name: str, args: argparse.Namespace) -> None:
        connection = self.schema.create_database()
        if args.input_model:
            input_model_path = Path(args.input_model)
            if not input_model_path.exists():
                raise FileNotFoundError(f"Input model file not found: {input_model_path}")
            self.schema.load_excel_data_model(connection, input_model_path)

        mapping_config = MappingConfig.from_excel(Path(args.map_file), table_name)
        normalization_config = NormalizationConfig.from_excel(Path(args.normalization_file), table_name)
        source_frame = pd.read_excel(Path(args.data_file), engine="openpyxl")

        importer = TableImporter(
            connection=connection,
            schema=self.schema,
            table_name=table_name,
            mapping_config=mapping_config,
            normalization_config=normalization_config,
        )
        importer.import_rows(source_frame)

        output_path = Path(args.output_file)
        self.schema.export_db_to_excel(connection, output_path)
        print(f"Wrote updated CMDB workbook to {output_path}")


def main() -> None:
    sql_schema_path = Path(__file__).resolve().parent / "db" / "cmdb_2026-06-28T15_49_08.585Z.sql"
    cli = CMDBCLI(sql_schema_path)
    cli.run()


if __name__ == "__main__":
    main()
