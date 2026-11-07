"""IOF XML to NormalizedCourse mappers."""

from __future__ import annotations

from ...models import NormalizedCourse


def normalize_plus_in_names(name: str) -> str:
    """Replace '+' character with ' Plus' in course/short names."""
    return name.replace("+", " Plus")


def map_course(raw: dict) -> list[NormalizedCourse]:
    """Map a raw course dict (possibly multiple variations) to NormalizedCourse rows."""
    name = normalize_plus_in_names(raw.get("name", ""))
    short_name = normalize_plus_in_names(raw.get("short_name") or name)

    rows: list[NormalizedCourse] = []
    for var in raw.get("variations", []):
        rows.append(
            NormalizedCourse(
                name=name,
                short_name=short_name,
                length=var.get("length", ""),
                climb=var.get("climb", "0"),
                controls=var.get("controls", []),
            )
        )
    return rows
