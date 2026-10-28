from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

HEADERS = [
    "report_name",
    "coverage",
    "article_impressions",
    "ad_value_dkk",
    "sentiment_positive",
    "sentiment_neutral",
    "sentiment_negative",
]

def format_thousand_sep(n):
    """Returner tal som tekst med punktummer for hver tusind."""
    if n is None:
        return ""
    return f"{n:,}".replace(",", ".")

def autosize(ws):
    for col in ws.columns:
        max_len = 0
        for cell in col:
            val = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(val))
        ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 2

def write_final_excel(path: str, rows: list):
    wb = Workbook()
    ws = wb.active
    ws.title = "Main"

    # Header
    ws.append(HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    # Data rows
    for row in rows:
        formatted_row = [
            row[0],  # report_name
            format_thousand_sep(row[1]),  # coverage
            format_thousand_sep(row[2]),  # impressions
            format_thousand_sep(row[3]),  # ad_value
            format_thousand_sep(row[4]),  # positive
            format_thousand_sep(row[5]),  # neutral
            format_thousand_sep(row[6]),  # negative
        ]
        ws.append(formatted_row)

    # Apply thousand-separator to numeric columns (B → G)
    for col in range(2, 8):
        letter = get_column_letter(col)
        for r in range(2, ws.max_row + 1):
            ws[f"{letter}{r}"].number_format = "# ##0"  # dansk tusindtalsformat

    autosize(ws)
    wb.save(path)
