from __future__ import annotations

import argparse
import sqlite3
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

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
    def __init__(self, sql_schema_path: Path) -> None:
        self.sql_schema_path = sql_schema_path
        self.table_names = set()

    def create_database(self) -> sqlite3.Connection:
        conn = sqlite3.connect(":memory:")
        conn.row_factory = sqlite3.Row
        schema_sql = self.sql_schema_path.read_text(encoding="utf-8")
        conn.executescript(schema_sql)
        self.table_names = self._discover_tables(conn)
        return conn

    @staticmethod
    def _discover_tables(conn: sqlite3.Connection) -> set[str]:
        cursor = conn.execute("SELECT name FROM sqlite_master WHERE type='table'")
        return {row[0] for row in cursor.fetchall()}

    def load_excel_data_model(self, conn: sqlite3.Connection, workbook_path: Path) -> None:
        workbook = pd.read_excel(workbook_path, sheet_name=None, engine="openpyxl")
        for sheet_name, df in workbook.items():
            if sheet_name not in self.table_names:
                raise ValueError(f"Excel model sheet '{sheet_name}' does not match any known CMDB table.")
            self._append_dataframe_to_table(conn, sheet_name, df)

    def _append_dataframe_to_table(self, conn: sqlite3.Connection, table_name: str, df: pd.DataFrame) -> None:
        if df.empty:
            return
        df = df.rename(columns=str)
        df.to_sql(table_name, conn, if_exists="append", index=False)

    def export_db_to_excel(self, conn: sqlite3.Connection, output_path: Path) -> None:
        tables = sorted(self.table_names)
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for table in tables:
                df = pd.read_sql_query(f'SELECT * FROM "{table}"', conn)
                df.to_excel(writer, sheet_name=table, index=False)

    @staticmethod
    def get_table_columns(conn: sqlite3.Connection, table_name: str) -> List[str]:
        cursor = conn.execute(f'PRAGMA table_info("{table_name}")')
        return [row[1] for row in cursor.fetchall()]

    @staticmethod
    def get_table_primary_keys(conn: sqlite3.Connection, table_name: str) -> List[str]:
        cursor = conn.execute(f'PRAGMA table_info("{table_name}")')
        return [row[1] for row in cursor.fetchall() if row[5] > 0]

    @staticmethod
    def get_required_columns(conn: sqlite3.Connection, table_name: str) -> List[str]:
        cursor = conn.execute(f'PRAGMA table_info("{table_name}")')
        required: List[str] = []
        for row in cursor.fetchall():
            name = row[1]
            notnull = row[3] == 1
            dflt_value = row[4]
            if notnull and dflt_value is None:
                required.append(name)
        return required


@dataclass
class MappingRule:
    source_column: str
    target_column: str
    table: Optional[str] = None


@dataclass
class MappingConfig:
    rules: List[MappingRule] = field(default_factory=list)

    @classmethod
    def from_excel(cls, workbook_path: Path, table_name: str) -> MappingConfig:
        workbook = pd.read_excel(workbook_path, sheet_name=None, engine="openpyxl")
        if table_name in workbook:
            df = workbook[table_name]
        else:
            df = next(iter(workbook.values()))
        df = df.rename(columns=str)
        missing = {"source_column", "target_column"} - set(df.columns)
        if missing:
            raise ValueError(
                f"Mapping file '{workbook_path}' is missing required columns: {', '.join(sorted(missing))}."
            )
        rules = []
        for _, row in df.iterrows():
            table = row.get("table") if "table" in df.columns else None
            rules.append(
                MappingRule(
                    source_column=str(row["source_column"]).strip(),
                    target_column=str(row["target_column"]).strip(),
                    table=str(table).strip() if pd.notna(table) else None,
                )
            )
        return cls([rule for rule in rules if rule.source_column and rule.target_column and (rule.table is None or rule.table == table_name)])

    def get_column_mapping(self) -> Dict[str, str]:
        return {rule.source_column: rule.target_column for rule in self.rules}


@dataclass
class NormalizationRule:
    field: str
    source_value: object
    target_value: object
    table: Optional[str] = None


@dataclass
class NormalizationConfig:
    rules: List[NormalizationRule] = field(default_factory=list)

    @classmethod
    def from_excel(cls, workbook_path: Path, table_name: str) -> NormalizationConfig:
        workbook = pd.read_excel(workbook_path, sheet_name=None, engine="openpyxl")
        if table_name in workbook:
            df = workbook[table_name]
        else:
            df = next(iter(workbook.values()))
        df = df.rename(columns=str)
        missing = {"field", "source_value", "target_value"} - set(df.columns)
        if missing:
            raise ValueError(
                f"Normalization file '{workbook_path}' is missing required columns: {', '.join(sorted(missing))}."
            )
        rules = []
        for _, row in df.iterrows():
            table = row.get("table") if "table" in df.columns else None
            rules.append(
                NormalizationRule(
                    field=str(row["field"]).strip(),
                    source_value=row["source_value"],
                    target_value=row["target_value"],
                    table=str(table).strip() if pd.notna(table) else None,
                )
            )
        return cls([rule for rule in rules if rule.field and (rule.table is None or rule.table == table_name)])

    def apply(self, table_name: str, data_frame: pd.DataFrame) -> pd.DataFrame:
        normalized = data_frame.copy()
        for rule in self.rules:
            if rule.table and rule.table != table_name:
                continue
            if rule.field not in normalized.columns:
                raise ValueError(
                    f"Normalization rule specifies field '{rule.field}' for table '{table_name}', but the imported data does not contain that column."
                )
            mask = normalized[rule.field] == rule.source_value
            if mask.any():
                normalized.loc[mask, rule.field] = rule.target_value
        return normalized


class TableImporter:
    def __init__(
        self,
        conn: sqlite3.Connection,
        schema: CMDBSchema,
        table_name: str,
        mapping: MappingConfig,
        normalization: NormalizationConfig,
    ) -> None:
        self.conn = conn
        self.schema = schema
        self.table_name = table_name
        self.mapping = mapping
        self.normalization = normalization
        self.table_columns = self.schema.get_table_columns(conn, table_name)
        self.required_columns = self.schema.get_required_columns(conn, table_name)
        self.primary_keys = self.schema.get_table_primary_keys(conn, table_name)

    def import_rows(self, source_df: pd.DataFrame) -> None:
        normalized_df = self._prepare_dataframe(source_df)
        self._write_rows(normalized_df)

    def _prepare_dataframe(self, source_df: pd.DataFrame) -> pd.DataFrame:
        df = source_df.rename(columns=str).copy()
        column_map = self.mapping.get_column_mapping()
        missing_sources = [src for src in column_map if src not in df.columns]
        if missing_sources:
            raise ValueError(
                f"Mapping file contains source columns that are missing from the import data: {', '.join(missing_sources)}"
            )
        df = df.rename(columns=column_map)
        missing_targets = [col for col in df.columns if col not in self.table_columns]
        if missing_targets:
            raise ValueError(
                f"Mapped data contains columns not present in target table '{self.table_name}': {', '.join(missing_targets)}"
            )
        normalized_df = self.normalization.apply(self.table_name, df)
        missing_required = [col for col in self.required_columns if col not in normalized_df.columns]
        if missing_required:
            raise ValueError(
                f"Target table '{self.table_name}' requires columns that are not provided by mapping: {', '.join(missing_required)}"
            )
        return normalized_df[self.table_columns].copy()

    def _write_rows(self, df: pd.DataFrame) -> None:
        if df.empty:
            return
        if not self.primary_keys:
            df.to_sql(self.table_name, self.conn, if_exists="append", index=False)
            return
        columns = [col for col in df.columns if col in self.table_columns]
        placeholders = ", ".join("?" for _ in columns)
        quoted_columns = ", ".join(f'"{col}"' for col in columns)
        sql = (
            f'INSERT OR REPLACE INTO "{self.table_name}" ({quoted_columns}) VALUES ({placeholders})'
        )
        rows = [tuple(row[col] for col in columns) for _, row in df.iterrows()]
        self.conn.executemany(sql, rows)
        self.conn.commit()


class CMDBCLI:
    def __init__(self, sql_schema_path: Path) -> None:
        self.sql_schema_path = sql_schema_path
        self.schema = CMDBSchema(sql_schema_path)

    def build_parser(self) -> argparse.ArgumentParser:
        parser = argparse.ArgumentParser(
            description="CMDB CLI ETL: import Excel data into an in-memory CMDB and export a reusable Excel workbook."
        )
        subparsers = parser.add_subparsers(dest="command", required=True)
        for subcommand, table_name in SUPPORTED_TABLES.items():
            command_parser = subparsers.add_parser(subcommand, help=f"Load data into the {table_name} table.")
            command_parser.add_argument("--data-file", required=True, help="Excel file containing import rows for the target table.")
            command_parser.add_argument("--map-file", required=True, help="Excel file defining source-to-target column mapping.")
            command_parser.add_argument(
                "--normalization-file",
                required=True,
                help="Excel file defining normalization rules for target values.",
            )
            command_parser.add_argument(
                "--input-model",
                required=False,
                help="Optional existing Excel workbook representing current CMDB data model.",
            )
            command_parser.add_argument(
                "--output-file",
                required=True,
                help="Output Excel workbook that will contain the updated CMDB state.",
            )
        return parser

    def run(self, argv: Optional[List[str]] = None) -> None:
        parser = self.build_parser()
        args = parser.parse_args(argv)
        if args.command not in SUPPORTED_TABLES:
            raise ValueError(f"Unknown command: {args.command}")
        table_name = SUPPORTED_TABLES[args.command]
        self._execute_command(table_name, args)

    def _execute_command(self, table_name: str, args: argparse.Namespace) -> None:
        conn = self.schema.create_database()
        if args.input_model:
            input_model_path = Path(args.input_model)
            if not input_model_path.exists():
                raise FileNotFoundError(f"Input model file not found: {input_model_path}")
            self.schema.load_excel_data_model(conn, input_model_path)
        mapping_config = MappingConfig.from_excel(Path(args.map_file), table_name)
        normalization_config = NormalizationConfig.from_excel(Path(args.normalization_file), table_name)
        source_df = pd.read_excel(Path(args.data_file), engine="openpyxl")
        importer = TableImporter(conn, self.schema, table_name, mapping_config, normalization_config)
        importer.import_rows(source_df)
        output_path = Path(args.output_file)
        self.schema.export_db_to_excel(conn, output_path)
        print(f"Wrote updated CMDB workbook to {output_path}")


def main() -> None:
    sql_schema_path = Path(__file__).resolve().parent / "db" / "cmdb_2026-06-28T15_49_08.585Z.sql"
    cli = CMDBCLI(sql_schema_path)
    cli.run()


if __name__ == "__main__":
    main()
