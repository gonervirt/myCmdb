from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from cmdb_cli import CMDBCLI, CMDBConfig


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Export a flattened server inventory workbook from the CMDB")
    parser.add_argument("--config", default="cmdb_config.json", help="JSON config file")
    parser.add_argument("--input-model", help="Existing CMDB model workbook to use as input")
    parser.add_argument("--output-file", help="Path to the Excel workbook that will receive the inventory export")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    config = CMDBConfig.from_json(Path(args.config)) if args.config else CMDBConfig.default()
    cli = CMDBCLI(config.sql_schema_path)

    input_model = Path(args.input_model) if args.input_model else config.cmdb_model
    output_file = Path(args.output_file) if args.output_file else input_model.with_name("server_inventory.xlsx")

    cli.export_server_inventory(input_model, output_file, config)


if __name__ == "__main__":
    main()
