from openpyxl.styles import (
    Font,
    Alignment,
    PatternFill,
    Border,
    Side,
    NamedStyle,
)


def openpyxl_styles(wb):
    """Create and return a dictionary of reusable openpyxl style objects.

    Adds a named style `cell_style` to the workbook if missing.
    """
    f_title = Font(size=16, bold=True, name="Times New Roman")
    font_bold_14 = Font(size=14, bold=True, name="Times New Roman")
    font_bold_12 = Font(size=12, bold=True, name="Times New Roman")
    font_regular_12 = Font(size=12, name="Times New Roman")
    align_center = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )

    bg_gray = PatternFill("solid", fgColor="DDDDDD")
    bg_red = PatternFill("solid", fgColor="F4CCCC")

    thin_line = Side(style="thin")
    thick_line = Side(style="thick")
    thin_border = Border(
        left=thin_line, right=thin_line, top=thin_line, bottom=thin_line
    )

    # Add named style if not present
    try:
        existing = [s.name for s in wb.named_styles]
    except Exception:
        existing = []

    if "cell_style" not in existing:
        cell_style = NamedStyle(name="cell_style")
        cell_style.font = font_regular_12
        cell_style.border = thin_border
        try:
            wb.add_named_style(cell_style)
        except Exception:
            pass

    return {
        "f_title": f_title,
        "font_bold_14": font_bold_14,
        "font_bold_12": font_bold_12,
        "font_regular_12": font_regular_12,
        "align_center": align_center,
        "bg_gray": bg_gray,
        "bg_red": bg_red,
        "thin_line": thin_line,
        "thick_line": thick_line,
        "thin_border": thin_border,
    }


def xlsxwriter_formats(workbook):
    """Return a dict of xlsxwriter formats for reuse in writers."""
    # xlsxwriter formats removed — only openpyxl styles are required for the
    # authoritative report generation. If xlsxwriter support is added later,
    # recreate formats here.
    raise NotImplementedError("xlsxwriter formats have been removed")
