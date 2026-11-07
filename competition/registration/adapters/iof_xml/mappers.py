from competition.registration.models import NormalizedCompetitor

_NS = "{http://www.orienteering.org/datastandard/3.0}"


def _text(element, path):
    found = element.find(path)
    return found.text if found is not None and found.text else ""


def map_iof_competitor(raw) -> NormalizedCompetitor:
    """Map an IOF PersonResult raw dict to a NormalizedCompetitor."""
    element = raw["element"]
    person = element.find(f"{_NS}Person")
    name_el = person.find(f"{_NS}Name") if person is not None else None
    family = _text(name_el, f"{_NS}Family") if name_el is not None else ""
    given = _text(name_el, f"{_NS}Given") if name_el is not None else ""
    name = f"{given} {family}".strip()
    id_el = person.find(f"{_NS}Id") if person is not None else None
    member_id = id_el.text if id_el is not None and id_el.text else ""
    result = element.find(f"{_NS}Result")
    bib_el = result.find(f"{_NS}BibNumber") if result is not None else None
    bib = bib_el.text if bib_el is not None and bib_el.text else ""
    return NormalizedCompetitor(
        member_id=member_id or bib,
        name=name,
        club="",
        category=raw.get("class_name", ""),
        si_card_number="",
        rental_requested=False,
        note="",
        phone="",
    )
