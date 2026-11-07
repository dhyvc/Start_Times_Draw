from dataclasses import dataclass


@dataclass
class NormalizedCompetitor:
    """Shared competitor contract used across input adapters and the SI rental domain."""

    member_id: str
    name: str
    club: str
    category: str
    si_card_number: str
    rental_requested: bool
    note: str
    phone: str
