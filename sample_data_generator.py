from __future__ import annotations

from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "sample_data"
OUT_DIR.mkdir(exist_ok=True)
DATA_DIR = ROOT / "data"
MAPPING_DIR = ROOT / "mapping"
NORMALIZATION_DIR = ROOT / "normalization"
MODEL_DIR = ROOT / "model"
for directory in (DATA_DIR, MAPPING_DIR, NORMALIZATION_DIR, MODEL_DIR):
    directory.mkdir(exist_ok=True)


def _write_workbook(filename: str, sheet_name: str, rows: list[dict[str, object]], out_dir: Path | None = None) -> None:
    target_dir = out_dir or OUT_DIR
    pd.DataFrame(rows).to_excel(target_dir / filename, sheet_name=sheet_name, index=False)


def _write_mapping(filename: str, sheet_name: str, mappings: list[dict[str, str]], out_dir: Path | None = None) -> None:
    target_dir = out_dir or OUT_DIR
    pd.DataFrame(mappings).to_excel(target_dir / filename, sheet_name=sheet_name, index=False)


def _write_normalization(filename: str, sheet_name: str, rules: list[dict[str, object]], out_dir: Path | None = None) -> None:
    target_dir = out_dir or OUT_DIR
    pd.DataFrame(rules).to_excel(target_dir / filename, sheet_name=sheet_name, index=False)


def _write_table_normalization(filename: str, rules_by_sheet: dict[str, list[dict[str, object]]], out_dir: Path | None = None) -> None:
    target_dir = out_dir or OUT_DIR
    with pd.ExcelWriter(target_dir / filename, engine="openpyxl") as writer:
        for sheet_name, rules in rules_by_sheet.items():
            pd.DataFrame(rules).to_excel(writer, sheet_name=sheet_name, index=False)


def build_sample_workbooks(out_dir: Path | None = None) -> None:
    target_dir = out_dir or OUT_DIR
    target_dir.mkdir(parents=True, exist_ok=True)

    if out_dir is None:
        data_dir = DATA_DIR
        mapping_dir = MAPPING_DIR
        normalization_dir = NORMALIZATION_DIR
        model_dir = MODEL_DIR
    else:
        data_dir = target_dir
        mapping_dir = target_dir
        normalization_dir = target_dir
        model_dir = target_dir

    for directory in (data_dir, mapping_dir, normalization_dir, model_dir):
        directory.mkdir(parents=True, exist_ok=True)

    model = {
        "Server": pd.DataFrame(
            [
                {"id": "srv-001", "Name": "srv-app-01", "Owner_id": "user-001", "Localisation_id": "loc-001", "OS_id": "os-001", "IP_Address_id": "ip-001"},
                {"id": "router-001", "Name": "router-001", "Owner_id": "user-001", "Localisation_id": "loc-001", "OS_id": "os-001", "IP_Address_id": "ip-001"},
            ]
        ),
        "Localisation": pd.DataFrame([{"id": "loc-001", "Country": "FR", "City": "Paris", "Room": "HQ"}]),
        "Application": pd.DataFrame([
            {
                "id": 1,
                "Owner": "user-001",
                "support": "user-002",
                "Name": "App A",
                "Description": "Primary app",
                "Criticality": "high",
                "Hosted": "router-001",
            }
        ]),
        "User": pd.DataFrame([{"id": "user-001", "Name": "Ada Lovelace", "email": "ada@example.com"}]),
        "IP address": pd.DataFrame([{"id": "ip-001", "IPv4": "10.0.0.10", "VLAN_id": "vlan-001", "IP public": False}]),
        "VLAN": pd.DataFrame([{"router": "router-001", "Gateway": "ip-001", "id": "vlan-001", "Name": "Prod", "CIDR": "10.0.0.0/24", "Advertised": True}]),
        "OS": pd.DataFrame([{"id": "os-001", "Name": "Ubuntu", "Version": "24.04", "EDR": True, "type": "linux"}]),
        "Team": pd.DataFrame([{"id": "team-001", "user": "user-001"}]),
    }

    def _write(filename: str, sheet_name: str, rows: list[dict[str, object]]) -> None:
        pd.DataFrame(rows).to_excel(data_dir / filename, sheet_name=sheet_name, index=False)

    def _write_map(filename: str, sheet_name: str, mappings: list[dict[str, str]]) -> None:
        pd.DataFrame(mappings).to_excel(mapping_dir / filename, sheet_name=sheet_name, index=False)

    def _write_norm(filename: str, sheet_name: str, rules: list[dict[str, object]]) -> None:
        pd.DataFrame(rules).to_excel(normalization_dir / filename, sheet_name=sheet_name, index=False)

    def _write_table_norm(filename: str, rules_by_sheet: dict[str, list[dict[str, object]]], field_names: list[str] | None = None) -> None:
        with pd.ExcelWriter(normalization_dir / filename, engine="openpyxl") as writer:
            for field_name in field_names or []:
                pd.DataFrame(columns=["source_value", "target_value"]).to_excel(
                    writer, sheet_name=field_name, index=False
                )
            for sheet_name, rules in rules_by_sheet.items():
                pd.DataFrame(rules).to_excel(writer, sheet_name=sheet_name, index=False)

    users = [
        {
            "id": f"user-{i:03d}",
            "source_name": f"User {i:03d}",
            "source_email": f"user{i:03d}@example.com",
        }
        for i in range(1, 51)
    ]
    _write("user_data.xlsx", "data", users)
    _write_map(
        "user_mapping.xlsx",
        "User",
        [
            {"source_column": "id", "target_column": "id"},
            {"source_column": "source_name", "target_column": "Name"},
            {"source_column": "source_email", "target_column": "email"},
        ],
    )
    _write_table_norm(
        "user_normalization.xlsx",
        {"Name": [{"source_value": "User 001", "target_value": "User 001"}]},
        list(model["User"].columns),
    )

    servers = [
        {
            "id": f"srv-{i:03d}",
            "source_name": f"srv-app-{i:03d}",
            "source_owner": f"user-{((i - 1) % 50) + 1:03d}",
            "source_location": f"loc-{((i - 1) % 50) + 1:03d}",
            "source_os": f"os-{((i - 1) % 50) + 1:03d}",
            "source_ip": f"ip-{((i - 1) % 50) + 1:03d}",
        }
        for i in range(1, 51)
    ]
    _write("server_data.xlsx", "data", servers)
    _write_map(
        "server_mapping.xlsx",
        "Server",
        [
            {"source_column": "id", "target_column": "id"},
            {"source_column": "source_name", "target_column": "Name"},
            {"source_column": "source_owner", "target_column": "Owner_id"},
            {"source_column": "source_location", "target_column": "Localisation_id"},
            {"source_column": "source_os", "target_column": "OS_id"},
            {"source_column": "source_ip", "target_column": "IP_Address_id"},
        ],
    )
    _write_table_norm(
        "server_normalization.xlsx",
        {
            "Owner_id": [{"source_value": "user-001", "target_value": "user-001"}],
            "Localisation_id": [{"source_value": "loc-001", "target_value": "loc-001"}],
            "OS_id": [{"source_value": "os-001", "target_value": "os-001"}],
            "IP_Address_id": [{"source_value": "ip-001", "target_value": "ip-001"}],
        },
        list(model["Server"].columns),
    )

    localisations = [
        {
            "id": f"loc-{i:03d}",
            "source_country": "FR",
            "source_city": f"City{i:03d}",
            "source_room": f"Room{i:03d}",
        }
        for i in range(1, 51)
    ]
    _write("localisation_data.xlsx", "data", localisations)
    _write_map(
        "localisation_mapping.xlsx",
        "Localisation",
        [
            {"source_column": "id", "target_column": "id"},
            {"source_column": "source_country", "target_column": "Country"},
            {"source_column": "source_city", "target_column": "City"},
            {"source_column": "source_room", "target_column": "Room"},
        ],
    )
    _write_table_norm(
        "localisation_normalization.xlsx",
        {"Country": [{"source_value": "FR", "target_value": "FR"}]},
        list(model["Localisation"].columns),
    )

    oss = [
        {
            "id": f"os-{i:03d}",
            "source_name": f"OS {i:03d}",
            "source_version": f"{24 + i % 5}.0{i % 10}",
            "source_edr": bool(i % 2),
            "source_type": "linux",
        }
        for i in range(1, 51)
    ]
    _write("os_data.xlsx", "data", oss)
    _write_map(
        "os_mapping.xlsx",
        "OS",
        [
            {"source_column": "id", "target_column": "id"},
            {"source_column": "source_name", "target_column": "Name"},
            {"source_column": "source_version", "target_column": "Version"},
            {"source_column": "source_edr", "target_column": "EDR"},
            {"source_column": "source_type", "target_column": "type"},
        ],
    )
    _write_table_norm(
        "os_normalization.xlsx",
        {"type": [{"source_value": "linux", "target_value": "linux"}]},
        list(model["OS"].columns),
    )

    vlans = [
        {
            "source_router": f"router-{i:03d}",
            "source_gateway": f"ip-{i:03d}",
            "id": f"vlan-{i:03d}",
            "source_name": f"VLAN {i:03d}",
            "source_cidr": f"10.{i % 10}.0.0/24",
            "source_advertised": bool(i % 2),
        }
        for i in range(1, 51)
    ]
    _write("vlan_data.xlsx", "data", vlans)
    _write_map(
        "vlan_mapping.xlsx",
        "VLAN",
        [
            {"source_column": "source_router", "target_column": "router"},
            {"source_column": "source_gateway", "target_column": "Gateway"},
            {"source_column": "id", "target_column": "id"},
            {"source_column": "source_name", "target_column": "Name"},
            {"source_column": "source_cidr", "target_column": "CIDR"},
            {"source_column": "source_advertised", "target_column": "Advertised"},
        ],
    )
    _write_table_norm(
        "vlan_normalization.xlsx",
        {"Advertised": [{"source_value": False, "target_value": False}]},
        list(model["VLAN"].columns),
    )

    ip_addresses = [
        {
            "id": f"ip-{i:03d}",
            "source_ipv4": f"192.168.{i // 256}.{i % 256}",
            "source_vlan_id": f"vlan-{((i - 1) % 50) + 1:03d}",
            "source_ip_public": bool(i % 2),
        }
        for i in range(1, 51)
    ]
    _write("ip_address_data.xlsx", "data", ip_addresses)
    _write_map(
        "ip_address_mapping.xlsx",
        "IP address",
        [
            {"source_column": "id", "target_column": "id"},
            {"source_column": "source_ipv4", "target_column": "IPv4"},
            {"source_column": "source_vlan_id", "target_column": "VLAN_id"},
            {"source_column": "source_ip_public", "target_column": "IP public"},
        ],
    )
    _write_table_norm(
        "ip_address_normalization.xlsx",
        {"IP public": [{"source_value": False, "target_value": False}]},
        list(model["IP address"].columns),
    )

    applications = [
        {
            "id": i,
            "source_owner": f"user-{(i % 50) + 1:03d}",
            "source_support": f"user-{((i + 1) % 50) + 1:03d}",
            "source_name": f"App {i:03d}",
            "source_description": f"Application {i:03d}",
            "source_criticality": "medium",
            "source_hosted": f"router-{((i - 1) % 50) + 1:03d}",
        }
        for i in range(1, 51)
    ]
    _write("application_data.xlsx", "data", applications)
    _write_map(
        "application_mapping.xlsx",
        "Application",
        [
            {"source_column": "id", "target_column": "id"},
            {"source_column": "source_owner", "target_column": "Owner"},
            {"source_column": "source_support", "target_column": "support"},
            {"source_column": "source_name", "target_column": "Name"},
            {"source_column": "source_description", "target_column": "Description"},
            {"source_column": "source_criticality", "target_column": "Criticality"},
            {"source_column": "source_hosted", "target_column": "Hosted"},
        ],
    )
    _write_table_norm(
        "application_normalization.xlsx",
        {"Criticality": [{"source_value": "medium", "target_value": "medium"}]},
        list(model["Application"].columns),
    )

    teams = [
        {"id": f"team-{i:03d}", "source_user": f"user-{((i - 1) % 50) + 1:03d}"}
        for i in range(1, 51)
    ]
    _write("team_data.xlsx", "data", teams)
    _write_map(
        "team_mapping.xlsx",
        "Team",
        [
            {"source_column": "id", "target_column": "id"},
            {"source_column": "source_user", "target_column": "user"},
        ],
    )
    _write_table_norm(
        "team_normalization.xlsx",
        {"user": [{"source_value": "user-001", "target_value": "user-001"}]},
        list(model["Team"].columns),
    )

    with pd.ExcelWriter(model_dir / "cmdb_model.xlsx", engine="openpyxl") as writer:
        for sheet_name, frame in model.items():
            frame.to_excel(writer, sheet_name=sheet_name, index=False)


if __name__ == "__main__":
    build_sample_workbooks()
