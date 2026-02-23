import pandas as pd
from source.excel_writers import generate_report_openpyxl
from source.camera_info_grabber import BASE_DIR


reports_dir = BASE_DIR / "reports"
reports_dir.mkdir(exist_ok=True)


def generate_report(data: pd.DataFrame, out_path: str) -> None:
    """Generate report using the original openpyxl writer.

    This is a thin wrapper kept for compatibility.
    """

    return generate_report_openpyxl(data, out_path)
