from dataclasses import dataclass

from flask import abort


@dataclass
class SIRentalRequest:
    """Validated SI rental input."""
    box_ranges: list
    missing_cards: list
    available_si_numbers: list
    additional_cards: list


def validate_si_rental_input(box_starts, box_ends, missing, additional):
    """Validate SI box ranges, missing cards and additional cards.

    Returns a SIRentalRequest. Aborts 400 on invalid input.
    Ported from ISOA competition.validate_si_rent_data_user_input.
    """
    missing_cards = []
    box_ranges = []
    all_si_card_numbers = []
    for i in range(len(box_starts)):
        # an empty box row means no SI box was filled for this row
        if box_starts[i] == "" and box_ends[i] == "":
            continue
        if not box_starts[i].isdigit() or not box_ends[i].isdigit():
            abort(400, "SI box range start and end must be numbers.")
        # range start has to be smaller or equal to range end (one SI card is legal)
        if int(box_starts[i]) > int(box_ends[i]):
            abort(400, "SI box range start must be smaller or equal to range end.")
        box_ranges.append(range(int(box_starts[i]), int(box_ends[i]) + 1))
        # missing si card numbers has to be within SI box card number range
        if missing[i] != "":
            for si_card_number in missing[i].replace(" ", "").split(","):
                if not si_card_number.isdigit():
                    abort(400, "Missing SI card numbers must be numbers.")
                if int(si_card_number) not in [*range(int(box_starts[i]), (int(box_ends[i]) + 1))]:
                    abort(400, "Missing SI card number is outside the SI box range.")
                missing_cards.append(int(si_card_number))
    for si_range in box_ranges:
        all_si_card_numbers.extend([*si_range])
    # additional free text SI cards available for rental (supports ranges like 1-3, 5, 8-10)
    additional_cards_list = []
    if additional and additional.strip() != "":
        for token in additional.replace(" ", "").split(","):
            if token == "":
                continue
            if "-" in token:
                parts = token.split("-")
                if len(parts) != 2:
                    abort(400, "Invalid additional SI card range.")
                range_start_str, range_end_str = parts
                if not range_start_str.isdigit() or not range_end_str.isdigit():
                    abort(400, "Additional SI card range must contain numbers.")
                range_start = int(range_start_str)
                range_end = int(range_end_str)
                if range_start > range_end:
                    abort(400, "Additional SI card range start must be smaller or equal to range end.")
                si_numbers = range(range_start, range_end + 1)
            else:
                if not token.isdigit():
                    abort(400, "Additional SI cards must be numbers or ranges.")
                si_numbers = [int(token)]
            for si_card_number_int in si_numbers:
                if si_card_number_int not in all_si_card_numbers:
                    additional_cards_list.append(si_card_number_int)
    all_si_card_numbers.extend(additional_cards_list)
    available_si_numbers = [el for el in all_si_card_numbers if el not in missing_cards]
    return SIRentalRequest(
        box_ranges=box_ranges,
        missing_cards=missing_cards,
        available_si_numbers=available_si_numbers,
        additional_cards=additional_cards_list,
    )
