from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from cmdb_cli import CMDBCLI, CMDBConfig


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Load User table data into the CMDB model")
    parser.add_argument("--config", default="cmdb_config.json", help="JSON config file")
    parser.add_argument("--data-file", help="Excel file containing rows to import")
    parser.add_argument("--map-file", help="Excel file defining source-to-target column mapping")
    parser.add_argument("--normalization-file", help="Excel file defining normalization rules")
    parser.add_argument("--input-model", help="Existing CMDB model workbook to use as input")
    parser.add_argument("--output-file", help="Target CMDB workbook output file")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    config = CMDBConfig.from_json(Path(args.config)) if args.config else CMDBConfig.default()
    cli = CMDBCLI(config.sql_schema_path)
    argv = ["--config", str(Path(args.config))] if args.config else []
    argv.append("user")
    if args.data_file:
        argv.extend(["--data-file", args.data_file])
    if args.map_file:
        argv.extend(["--map-file", args.map_file])
    if args.normalization_file:
        argv.extend(["--normalization-file", args.normalization_file])
    if args.input_model:
        argv.extend(["--input-model", args.input_model])
    if args.output_file:
        argv.extend(["--output-file", args.output_file])
    else:
        argv.extend(["--output-file", str(config.cmdb_model)])
    cli.run(argv)


if __name__ == "__main__":
    main()
