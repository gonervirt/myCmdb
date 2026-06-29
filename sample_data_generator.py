from __future__ import annotations

from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "sample_data"
OUT_DIR.mkdir(exist_ok=True)


def build_sample_workbooks() -> None:
    # One worksheet per import use case.
    sample_server = pd.DataFrame(
        [
            {
                "source_name": "srv-app-01",
                "source_owner": "user-001",
                "source_location": "loc-001",
                "source_os": "os-001",
                "source_ip": "ip-001",
            }
        ]
    )
    sample_server.to_excel(OUT_DIR / "server_data.xlsx", sheet_name="data", index=False)

    sample_mapping = pd.DataFrame(
        [
            {"source_column": "source_name", "target_column": "Name"},
            {"source_column": "source_owner", "target_column": "Owner_id"},
            {"source_column": "source_location", "target_column": "Localisation_id"},
            {"source_column": "source_os", "target_column": "OS_id"},
            {"source_column": "source_ip", "target_column": "IP_Address_id"},
        ]
    )
    sample_mapping.to_excel(OUT_DIR / "server_map.xlsx", sheet_name="Server", index=False)

    sample_normalization = pd.DataFrame(
        [
            {"field": "Owner_id", "source_value": "user-001", "target_value": "user-001"},
            {"field": "Localisation_id", "source_value": "loc-001", "target_value": "loc-001"},
            {"field": "OS_id", "source_value": "os-001", "target_value": "os-001"},
            {"field": "IP_Address_id", "source_value": "ip-001", "target_value": "ip-001"},
        ]
    )
    sample_normalization.to_excel(OUT_DIR / "server_normalization.xlsx", sheet_name="Server", index=False)

    # Create a simple reusable model workbook with one sheet per table.
    model = {
        "Server": pd.DataFrame(
            [{"id": "srv-001", "Name": "srv-app-01", "IP_Address_id": "ip-001", "OS_id": "os-001"}]
        ),
        "Localisation": pd.DataFrame([{"id": "loc-001", "Country": "FR", "City": "Paris"}]),
        "Application": pd.DataFrame([{"id": 1, "Name": "App A", "Owner": "user-001"}]),
        "User": pd.DataFrame([{"id": "user-001", "Name": "Ada Lovelace", "email": "ada@example.com"}]),
        "IP address": pd.DataFrame([{"id": "ip-001", "IPv4": "10.0.0.10", "VLAN_id": "vlan-001", "IP public": False}]),
        "VLAN": pd.DataFrame([{"id": "vlan-001", "Name": "Prod", "CIDR": "10.0.0.0/24", "Advertised": True}]),
        "OS": pd.DataFrame([{"id": "os-001", "Name": "Ubuntu", "Version": "24.04"}]),
        "Team": pd.DataFrame([{"id": "team-001", "user": "user-001"}]),
    }

    with pd.ExcelWriter(OUT_DIR / "cmdb_model.xlsx", engine="openpyxl") as writer:
        for sheet_name, frame in model.items():
            frame.to_excel(writer, sheet_name=sheet_name, index=False)


if __name__ == "__main__":
    build_sample_workbooks()
