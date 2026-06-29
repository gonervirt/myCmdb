# myCmdb

## Overview

This repository provides a CLI-based CMDB ETL tool that loads Excel table data into an in-memory SQLite model, applies mapping and normalization rules, and writes the updated CMDB state back to an Excel workbook.

Each table has its own wrapper CLI script under `scripts/`, so you can import one table at a time.

## Supported table CLI scripts

- `scripts/load_server.py`
- `scripts/load_localisation.py`
- `scripts/load_application.py`
- `scripts/load_user.py`
- `scripts/load_ip_address.py`
- `scripts/load_vlan.py`
- `scripts/load_os.py`
- `scripts/load_team.py`

## Requirements

- Python 3.12+
- `pandas`
- `openpyxl`

Install dependencies with:

```bash
python -m pip install -r requirements.txt
```

## Config file

Use `cmdb_config.example.json` as a template. The config file defines common paths for:

- `sql_schema_path`: path to the CMDB SQL schema file
- `cmdb_model`: path to the input/output Excel workbook used as the CMDB model
- `backup_dir`: directory where existing output workbooks are backed up with a timestamp
- `data_dir`: directory containing source Excel data files
- `map_dir`: directory containing mapping Excel files
- `normalization_dir`: directory containing normalization Excel files

Example config:

```json
{
  "sql_schema_path": "db/cmdb_2026-06-28T15_49_08.585Z.sql",
  "cmdb_model": "sample_data/cmdb_model.xlsx",
  "backup_dir": "backup",
  "data_dir": "sample_data",
  "map_dir": "sample_data",
  "normalization_dir": "sample_data"
}
```

## How to run

Activate your environment and run one of the table-specific scripts.

Example: load users

```bash
python scripts/load_user.py \
  --config cmdb_config.json \
  --output-file sample_data/output.xlsx
```

Example: load servers

```bash
python scripts/load_server.py \
  --config cmdb_config.json \
  --output-file sample_data/output.xlsx

  python scripts/load_server.py config cmdb_config.json --output-file sample_data/output.xlsx
```

Each script accepts the same optional file arguments:

- `--data-file` (Excel file with rows to import)
- `--map-file` (Excel file mapping source columns to CMDB columns)
- `--normalization-file` (Excel file defining normalization rules)
- `--input-model` (Excel workbook used as the current CMDB state)
- `--output-file` (target workbook path)

If you omit `--data-file`, `--map-file`, or `--normalization-file`, the script will resolve defaults from the config file.

## Backup behavior

When writing the output workbook, an existing file is backed up automatically to the configured `backup_dir` with a UTC timestamp suffix.

## Sample data files

Built-in sample data files are stored under `sample_data/`.

### Example sample files included

- `sample_data/cmdb_model.xlsx` — current CMDB model workbook used as input/output
- `sample_data/server_data.xlsx` — source server rows to import
- `sample_data/server_map.xlsx` — server mapping rules
- `sample_data/server_normalization.xlsx` — server normalization rules
- `sample_data/user_mapping.xlsx` — user mapping rules
- `sample_data/user_normalization.xlsx` — user normalization rules
- `sample_data/output.xlsx` — example output workbook created by a load command

### Sample sheet structure

`server_data.xlsx` sheet `data` columns:

- `source_name`
- `source_owner`
- `source_location`
- `source_os`
- `source_ip`

`server_map.xlsx` sheet `Server` columns:

- `source_column`
- `target_column`

`server_normalization.xlsx` sheet `Server` columns:

- `field`
- `source_value`
- `target_value`

`user_mapping.xlsx` sheet `User` columns:

- `source_column`
- `target_column`

`user_normalization.xlsx` sheet `User` columns:

- `field`
- `source_value`
- `target_value`

## Notes

- `cmdb_model.xlsx` is intended to be reused as both input and output.
- The CLI workflows are table-specific, so each import command targets one CMDB table at a time.
- The output workbook contains one Excel sheet per CMDB table.

