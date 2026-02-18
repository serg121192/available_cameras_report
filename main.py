from datetime import datetime

from source.camera_info_grabber import camera_info_grabber
from source.report_generator import generate_report


if __name__ == "__main__":
    data = camera_info_grabber()
    generate_report(
        data,
        f"reports/Непрацюючі_камери {datetime.now().strftime('%d.%m.%Y')}.xlsx",
    )
