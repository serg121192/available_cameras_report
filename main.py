from source.camera_info_grabber import camera_info_grabber
from source.report_generator import generate_report


if __name__ == "__main__":
    data = camera_info_grabber()
    generate_report(data, "Непрацюючі_камери.xlsx")
