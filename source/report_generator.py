from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (
    Font,
    Alignment,
    PatternFill,
    Border,
    Side
)
from openpyxl.utils import get_column_letter


def generate_report(data: pd.DataFrame, out_path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = datetime.now()

    f_bold = Font(bold=True)
    f_title = Font(size=15, bold=True)
    align_center = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )

    bg_gray = PatternFill("solid", fgColor="DDDDDD")
    bg_red = PatternFill("solid", fgColor="F4CCCC")

    thin_line = Side(style="thin")
    thin_border = Border(
        left=thin_line,
        right=thin_line,
        top=thin_line,
        bottom=thin_line
    )
