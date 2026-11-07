"""XLSX report generation for ISOA Excel validation results."""

import io

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

_HEADERS = [
    "Member ID",
    "Name",
    "Club",
    "Category",
    "SI Card",
    "Comment",
    "Phone",
    "Reason",
]

_COLUMN_WIDTHS = [14, 22, 16, 14, 12, 22, 16, 48]


def generate_validation_xlsx(problems, title=None):
    """Generate an XLSX report describing all ISOA validation problems.

    Args:
        problems: list of ValidationProblem (from validate_rows / validate_competitors).
        title: optional report title (defaults to English).

    Returns:
        The XLSX document as bytes.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Validation Report"

    header_font = Font(bold=True)
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(wrap_text=True, vertical="top")

    row = 1
    ws.cell(row=row, column=1, value=title or "ISOA Excel Validation Report")
    row += 1
    if not problems:
        ws.cell(row=row, column=1, value="No possible problems found in the file.")
        row += 1
    else:
        ws.cell(
            row=row,
            column=1,
            value=f"Found {len(problems)} possible problem(s) in the file:",
        )
        row += 1

    for col, header in enumerate(_HEADERS, start=1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    row += 1

    for problem in problems:
        values = [
            problem.member_id,
            problem.name,
            problem.club,
            problem.category,
            problem.si_card_number,
            problem.note,
            problem.phone,
            problem.reason,
        ]
        for col, value in enumerate(values, start=1):
            cell = ws.cell(row=row, column=col, value=value)
            cell.border = border
            cell.alignment = wrap
        row += 1

    for col, width in enumerate(_COLUMN_WIDTHS, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
