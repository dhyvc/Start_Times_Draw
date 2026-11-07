import io

import xlsxwriter


def generate_si_rental_excel(box_ranges, missing_cards, renders, additional_cards):
    """Build the SI rental xlsx workbook (xlsxwriter) and return its bytes.

    One worksheet per box range ("מזוודה N") plus one for additional cards
    ("כרטיסים נוספים"). Hebrew headers, right-to-left worksheets.
    Ported from ISOA competition.populate_si_rental_worksheets.
    """
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    _populate_sheets(workbook, box_ranges, missing_cards, renders, additional_cards)
    workbook.close()
    output.seek(0)
    return output.read()


def _populate_sheets(workbook, box_ranges, missing_cards, renders, additional_cards):
    index = 1
    for number_range in box_ranges:
        _create_sheet(workbook, "מזוודה " + str(index), [*number_range], missing_cards, renders)
        index += 1
    if additional_cards:
        _create_sheet(workbook, "כרטיסים נוספים", list(additional_cards), [], renders)


def _create_sheet(workbook, sheet_name, si_numbers, missing_numbers, renders_list):
    worksheet = workbook.add_worksheet(sheet_name)
    merge_format = workbook.add_format({
        'bold': 1,
        'underline': 1,
        'font_name': 'Arial',
        'font_size': 14,
        'align': 'center',
        'valign': 'vcenter',
        'reading_order': 2})
    headers_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'font_name': 'Arial',
        'font_size': 12,
        'align': 'center',
        'valign': 'vcenter',
        'reading_order': 2})
    text_format = workbook.add_format({
        'border': 1,
        'font_name': 'Arial',
        'font_size': 12,
        'align': 'center',
        'valign': 'vcenter',
        'reading_order': 2})
    missing_format = workbook.add_format({
        'border': 1,
        'font_name': 'Arial',
        'font_size': 12,
        'fg_color': 'red'})
    worksheet.right_to_left()
    worksheet.set_column('A:A', 11)
    worksheet.set_column('B:B', 14)
    worksheet.set_column('C:C', 25)
    worksheet.set_column('D:D', 14)
    worksheet.set_column('E:E', 12)
    worksheet.set_column('F:F', 12)
    worksheet.set_column('G:G', 14)
    worksheet.set_column('H:H', 14)

    worksheet.merge_range('A1:H1', 'השאלת כרטיסי SI', merge_format)
    worksheet.write('A3', "מס' כרטיס", headers_format)
    worksheet.write('B3', "מס' חבר/ת.ז.", headers_format)
    worksheet.write('C3', "שם", headers_format)
    worksheet.write('D3', "טלפון", headers_format)
    worksheet.write('E3', "תעודת זהות", headers_format)
    worksheet.write('F3', "רישיון נהיגה", headers_format)
    worksheet.write('G3', "לקיחת כרטיס", headers_format)
    worksheet.write('H3', "החזרת כרטיס", headers_format)

    worksheet.write('A4', "", headers_format)
    worksheet.write('B4', "", headers_format)
    worksheet.write('C4', "", headers_format)
    worksheet.write('D4', "", headers_format)
    worksheet.write('E4', "", headers_format)
    worksheet.write('F4', "", headers_format)
    worksheet.write('G4', "", headers_format)
    worksheet.write('H4', "", headers_format)

    si_numbers_row_index = 5
    for si_number in si_numbers:
        worksheet.write('A' + str(si_numbers_row_index), str(si_number), text_format)
        if si_number in missing_numbers:
            worksheet.write('B' + str(si_numbers_row_index), "", missing_format)
            worksheet.write('C' + str(si_numbers_row_index), "", missing_format)
            worksheet.write('D' + str(si_numbers_row_index), "", missing_format)
            worksheet.write('E' + str(si_numbers_row_index), "", missing_format)
            worksheet.write('F' + str(si_numbers_row_index), "", missing_format)
            worksheet.write('H' + str(si_numbers_row_index), "", missing_format)
            worksheet.write('G' + str(si_numbers_row_index), "", missing_format)
        else:
            worksheet.write('B' + str(si_numbers_row_index), "", text_format)
            worksheet.write('C' + str(si_numbers_row_index), "", text_format)
            worksheet.write('D' + str(si_numbers_row_index), "", text_format)
            worksheet.write('E' + str(si_numbers_row_index), "", text_format)
            worksheet.write('F' + str(si_numbers_row_index), "", text_format)
            worksheet.write('H' + str(si_numbers_row_index), "", text_format)
            worksheet.write('G' + str(si_numbers_row_index), "", text_format)
            for render in renders_list:
                if render.si_card_number == str(si_number):
                    worksheet.write('B' + str(si_numbers_row_index), str(render.member_id), text_format)
                    worksheet.write('C' + str(si_numbers_row_index), str(render.name), text_format)
                    worksheet.write('D' + str(si_numbers_row_index), str(render.phone), text_format)
                    worksheet.write('E' + str(si_numbers_row_index), "", text_format)
                    worksheet.write('F' + str(si_numbers_row_index), "", text_format)
                    worksheet.write('H' + str(si_numbers_row_index), "", text_format)
        si_numbers_row_index += 1
