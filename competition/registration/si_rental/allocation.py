def allocate_si_cards(competitors, available_si_numbers):
    """Assign SI numbers to rental-requested competitors.

    Returns the list of competitors that were assigned an SI card.
    Allocation stops when available_si_numbers is exhausted (never raises).
    """
    si_number_index = 0
    renders = []
    for competitor in competitors:
        if not competitor.rental_requested:
            continue
        if si_number_index >= len(available_si_numbers):
            break
        competitor.si_card_number = str(available_si_numbers[si_number_index])
        renders.append(competitor)
        si_number_index += 1
    return renders
