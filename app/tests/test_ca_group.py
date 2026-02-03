from datetime import datetime, timedelta

from app.engine.ca_group import get_ca_group


def test_get_ca_group_boundaries() -> None:
    now = datetime.now()

    assert get_ca_group(now - timedelta(days=1)) == "0–3 Days"
    assert get_ca_group(now - timedelta(days=4)) == "3–5 Days"
    assert get_ca_group(now - timedelta(days=7)) == "5–10 Days"
    assert get_ca_group(now - timedelta(days=12)) == "10–15 Days"
    assert get_ca_group(now - timedelta(days=20)) == "15–30 Days"
    assert get_ca_group(now - timedelta(days=40)) == "30–60 Days"
    assert get_ca_group(now - timedelta(days=70)) == "60–90 Days"
    assert get_ca_group(now - timedelta(days=120)) == "> 90 Days"
