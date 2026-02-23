from datetime import datetime
import time
import logging

from source.camera_info_grabber import camera_info_grabber
from source.report_generator import generate_report


def _configure_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )


if __name__ == "__main__":
    _configure_logging()
    out_path = f"reports/Непрацюючі_камери {datetime.now().strftime('%d.%m.%Y')}.xlsx"

    t0 = time.perf_counter()
    data = camera_info_grabber()
    t1 = time.perf_counter()
    logging.info("camera_info_grabber: %.3fs", t1 - t0)

    t2 = time.perf_counter()
    generate_report(data, out_path)
    t3 = time.perf_counter()
    logging.info("generate_report: %.3fs", t3 - t2)
    logging.info("Report saved to %s", out_path)
