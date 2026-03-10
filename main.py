"""
competency_wizard/main.py
職能說明書精靈 — 入口點
用法：python -m competency_wizard.main
  或：python competency_wizard/main.py
"""

import sys
from pathlib import Path

# 確保可從此目錄直接執行
_here = Path(__file__).parent
if str(_here) not in sys.path:
    sys.path.insert(0, str(_here))

from PyQt6.QtWidgets import QApplication

from wizard_ui import WizardMainWindow


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    win = WizardMainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
