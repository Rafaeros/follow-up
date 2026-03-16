import os
import sys
import asyncio
import logging
import qasync
import glob
from PySide6.QtWidgets import QApplication, QMessageBox, QDialog

from src.core.config import ConfigManager
from src.core.session_manager import SessionManager
from src.frontend.interface import Interface


def load_stylesheet(app: QApplication) -> None:
    """Loads the global QSS stylesheet for the application."""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    qss_path = os.path.join(base_dir, "src", "frontend", "theme.qss")
    if os.path.exists(qss_path):
        with open(qss_path, "r", encoding="utf-8") as f:
            app.setStyleSheet(f.read())
    else:
        logging.warning("Theme file not found at %s", qss_path)


def main() -> None:
    """Application entry point."""
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    load_stylesheet(app)

    loop = qasync.QEventLoop(app)
    asyncio.set_event_loop(loop)

    # Initialize Core Managers
    config_manager = ConfigManager()
    session_manager = SessionManager(config_manager)

    # Initialize Main Interface
    window = Interface(
        config_manager=config_manager,
        session_manager=session_manager,
    )
    window.show()

    app.setQuitOnLastWindowClosed(True)
    with loop:
        loop.run_forever()

async def test_report_scraper():
    """Test function for ReportScraper using correct await patterns."""
    from src.core.scraper import ReportScraper
    from src.core.session_manager import SessionManager
    from src.core.config import ConfigManager

    config_manager = ConfigManager()
    session_manager = SessionManager(config_manager)

    try:
        # CORREÇÃO 1: Use 'await' em vez de 'asyncio.run()' aqui dentro
        success = await session_manager.login()
        
        if success:
            scraper = ReportScraper(session_manager)
            # A data deve ser no formato que o site espera ou seu scraper trata
            await scraper.get_all_reports('01/03/2026', '31/03/2026')
        else:
            logging.error("Login failed, skipping report fetch.")

    finally:
        # CORREÇÃO 2: Sempre use await no close e garanta que rode no 'finally'
        await session_manager.close()
        logging.info("Session closed safely.")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
