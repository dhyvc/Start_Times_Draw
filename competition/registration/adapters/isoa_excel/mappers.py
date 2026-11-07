from competition.registration.models import NormalizedCompetitor

MEMBER_ID = "0"
NAME = "1"
CLUB = "2"
CATEGORY = "3"
SI_CARD = "6"
NOTE = "9"
RENTAL_FLAG = "10"
PHONE = "15"

_TRUE_VALUES = {"true", "1", "yes", "y"}


def _is_rental_requested(value) -> bool:
    return str(value).strip().lower() in _TRUE_VALUES


def map_row(row) -> NormalizedCompetitor:
    """Map a raw ISOA excel row dict to a NormalizedCompetitor."""
    return NormalizedCompetitor(
        member_id=str(row[MEMBER_ID]).strip(),
        name=str(row[NAME]).strip(),
        club=str(row[CLUB]).strip(),
        category=str(row[CATEGORY]).strip(),
        si_card_number=str(row[SI_CARD]).strip(),
        rental_requested=_is_rental_requested(row[RENTAL_FLAG]),
        note=str(row[NOTE]).strip(),
        phone=str(row[PHONE]).strip(),
    )
