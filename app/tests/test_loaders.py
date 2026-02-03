from pathlib import Path

import pandas as pd

from app.core.loaders.dump_loader import load_dump


def test_load_dump_parses_rows(tmp_path: Path) -> None:
    data = pd.DataFrame(
        [
            {
                "case_id": " CASE-1 ",
                "customer_name": "Customer",
                "created_on": 45200,
                "created_by": "user",
                "country": "US",
                "resolution_code": "parts shipped",
                "case_owner": "Owner",
                "otc_code": "OTC",
                "serial_number": "SN",
                "product_name": "Product",
                "email_status": "sent",
            },
            {
                "case_id": "",
                "customer_name": "Skip",
                "created_on": 45201,
                "created_by": "user",
                "country": "US",
                "resolution_code": "parts shipped",
                "case_owner": "Owner",
                "otc_code": "OTC",
                "serial_number": "SN",
                "product_name": "Product",
                "email_status": "sent",
            },
        ]
    )
    file_path = tmp_path / "dump.csv"
    data.to_csv(file_path, index=False)

    records = load_dump(str(file_path))

    assert len(records) == 1
    assert records[0].case_id == "CASE-1"
    assert records[0].created_on.year == 2023
