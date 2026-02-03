from datetime import date, datetime, timedelta

from app.engine.sbd import calculate_sbd
from app.schemas.config import SBDConfig


def test_calculate_sbd_na_for_country() -> None:
    config = SBDConfig(
        period_start=date(2024, 1, 1),
        period_end=date(2024, 12, 31),
        cutoff_countries=["US"],
    )

    result = calculate_sbd(
        case_created_on=datetime(2024, 6, 1, 10, 0),
        country="CA",
        first_order_date=datetime(2024, 6, 1, 12, 0),
        config=config,
    )

    assert result == "NA"


def test_calculate_sbd_na_for_out_of_period() -> None:
    config = SBDConfig(
        period_start=date(2024, 1, 1),
        period_end=date(2024, 1, 31),
        cutoff_countries=["US"],
    )

    result = calculate_sbd(
        case_created_on=datetime(2024, 2, 1, 10, 0),
        country="US",
        first_order_date=datetime(2024, 2, 1, 12, 0),
        config=config,
    )

    assert result == "NA"


def test_calculate_sbd_met_with_rollover() -> None:
    config = SBDConfig(
        period_start=date(2024, 1, 1),
        period_end=date(2024, 12, 31),
        cutoff_countries=["US"],
    )

    created_on = datetime(2024, 6, 1, 23, 0)
    first_order_date = created_on + timedelta(days=1)

    result = calculate_sbd(
        case_created_on=created_on,
        country="US",
        first_order_date=first_order_date,
        config=config,
    )

    assert result == "Met"


def test_calculate_sbd_not_met_without_order() -> None:
    config = SBDConfig(
        period_start=date(2024, 1, 1),
        period_end=date(2024, 12, 31),
        cutoff_countries=["US"],
    )

    result = calculate_sbd(
        case_created_on=datetime(2024, 6, 1, 10, 0),
        country="US",
        first_order_date=None,
        config=config,
    )

    assert result == "Not Met"
