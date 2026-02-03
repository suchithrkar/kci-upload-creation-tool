"""Schemas for configuration mappings."""

from datetime import date

from pydantic import BaseModel


class TLMap(BaseModel):
    """Maps a team lead name to assigned agents."""

    name: str
    agents: list[str]


class MarketMap(BaseModel):
    """Maps a market name to included countries."""

    name: str
    countries: list[str]


class SBDConfig(BaseModel):
    """Configuration for SBD date periods and country cutoffs."""

    period_start: date
    period_end: date
    cutoff_countries: list[str]
