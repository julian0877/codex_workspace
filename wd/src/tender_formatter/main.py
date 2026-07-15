import sys

from PySide6.QtWidgets import QApplication

from tender_formatter.service import FormatterService
from tender_formatter.ui.main_window import MainWindow


def main() -> int:
    application = QApplication.instance() or QApplication(sys.argv)
    window = MainWindow(FormatterService())
    window.show()
    return application.exec()


if __name__ == "__main__":
    raise SystemExit(main())
