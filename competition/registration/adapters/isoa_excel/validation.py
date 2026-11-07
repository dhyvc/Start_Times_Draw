"""Validation of ISOA Excel startlist files.

Rules validated:
  1. Undefined registration: the class/category is missing.
  2. Invalid comment: the comment is not empty and is not "[מנוי ליגה]".
  3. Invalid SI card: the card number is 0 but no card rental was requested.
"""

from dataclasses import dataclass

from competition.registration.adapters.isoa_excel.mappers import map_row

# Comment that is considered valid / allowed and therefore not flagged as a problem.
ALLOWED_COMMENT = "[מנוי ליגה]"

_HEBREW_RANGE = range(0x0590, 0x05FF + 1)


def _contains_hebrew(text: str) -> bool:
    return any(ord(ch) in _HEBREW_RANGE for ch in str(text))


def _is_zero_card(value) -> bool:
    """Return True when the SI card value represents the number zero (e.g. "0", "0.0")."""
    if value is None:
        return False
    text = str(value).strip()
    if text == "":
        return False
    try:
        return float(text) == 0.0
    except ValueError:
        return False


@dataclass
class ValidationProblem:
    """A single validation problem found in an ISOA registration."""

    member_id: str
    name: str
    club: str
    category: str
    si_card_number: str
    note: str
    phone: str
    reason: str


def validate_competitors(competitors):
    """Validate a list of NormalizedCompetitor parsed from an ISOA Excel startlist.

    Returns a list of ValidationProblem, one per rule violation per competitor.
    """
    problems = []
    for competitor in competitors:
        category = str(competitor.category or "").strip()
        note = str(competitor.note or "").strip()
        si_card = str(competitor.si_card_number or "").strip()

        # Rule 1: undefined registration (missing class/category).
        if not category:
            problems.append(ValidationProblem(
                member_id=competitor.member_id,
                name=competitor.name,
                club=competitor.club,
                category=category,
                si_card_number=si_card,
                note=note,
                phone=competitor.phone,
                reason="Undefined registration: class/category is missing.",
            ))

        # Rule 2: comment is not empty and not "[מנוי ליגה]".
        if note and note != ALLOWED_COMMENT:
            problems.append(ValidationProblem(
                member_id=competitor.member_id,
                name=competitor.name,
                club=competitor.club,
                category=category,
                si_card_number=si_card,
                note=note,
                phone=competitor.phone,
                reason="Non standard comment",
            ))

        # Rule 3: card number is 0 but no card rental was requested.
        if _is_zero_card(si_card) and not competitor.rental_requested:
            problems.append(ValidationProblem(
                member_id=competitor.member_id,
                name=competitor.name,
                club=competitor.club,
                category=category,
                si_card_number=si_card,
                note=note,
                phone=competitor.phone,
                reason=(
                    "Invalid SI card: card number is 0 but no card rental was requested."
                ),
            ))

    return problems


def validate_rows(rows):
    """Validate raw ISOA Excel row dicts (as yielded by read_isoa_excel).

    Rows are mapped to NormalizedCompetitor before validation.
    """
    competitors = [map_row(row) for row in rows]
    return validate_competitors(competitors)
