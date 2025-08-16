import sys
import os
import logging
from pathlib import Path

from config_manager import ConfigManager
# from win_events import WindowsEventMonitor  # Removed for manual attendance
 
from word_handler import WordHandler
from web_server import AttendanceWebServer
from splash_screen import create_splash_screen

def setup_logging():
    """Setup logging configuration."""
    import os
    from pathlib import Path
    appdata = os.getenv('APPDATA') or os.path.expanduser('~')
    log_dir = Path(appdata) / "AttendanceTracker" / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    # Only create log_dir if/when a log file is actually written
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_dir / 'attendance.log'),
            logging.StreamHandler()
        ]
    )

def main():
    """Main application entry point."""
    setup_logging()
    logger = logging.getLogger(__name__)

    # Create QApplication first for splash screen
    from PyQt5.QtWidgets import QApplication
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    # Show splash screen
    splash_manager = create_splash_screen()
    splash = splash_manager.show_splash()

    logger.info("Starting Khaled's Attendance Tracker - PyQt Edition with Web Server")
    splash_manager.next_step("Starting Khaled's Attendance Tracker...")

    # Initialize components with progress updates
    splash_manager.next_step("Initializing configuration...")
    config = ConfigManager()

    # splash_manager.next_step("Setting up event monitoring...")
    event_monitor = None  # Windows event monitoring removed for manual attendance

    splash_manager.next_step("Loading document handler...")
    word_handler = WordHandler(config)

    splash_manager.next_step("Initializing web server for mobile access...")
    try:
        web_server = AttendanceWebServer(config, word_handler)
        logger.info("Web server initialized for mobile access")
    except Exception as e:
        logger.warning(f"Web server initialization failed: {e}. Running in desktop-only mode.")
        web_server = None


    splash_manager.next_step("Creating modern user interface...")
    # Use PyQt interface with web server
    from pyqt_ui import create_pyqt_app
    from pdf_handler import PDFHandler
    pdf_handler = PDFHandler(config)
    window = create_pyqt_app(config, event_monitor, pdf_handler, word_handler, web_server)
    logger.info("Using PyQt5 interface with mobile web server")

    splash_manager.next_step("Finalizing setup...")

    try:
        # Show main window
        window.show()
    finally:
        # Always finish splash screen, even if error occurs
        splash_manager.finish()

    # Run the application
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

