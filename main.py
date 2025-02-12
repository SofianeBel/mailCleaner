import sys
import logging
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt
from mail_cleaner import MailCleanerUI

# Configure logging with high precision timestamps
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s.%(msecs)03d %(levelname)s: %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.FileHandler('mail_cleaner.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def main():
    try:
        logger.debug("Starting Mail Cleaner application...")
        app = QApplication(sys.argv)
        logger.debug("QApplication initialized")
        
        app.setStyle('Fusion')
        logger.debug("Application style set to Fusion")
        
        window = MailCleanerUI()
        logger.debug("Main window created")
        
        # Set window to appear in front
        window.setWindowState(window.windowState() & ~Qt.WindowState.WindowMinimized | Qt.WindowState.WindowActive)
        window.show()
        # Activate the window
        window.activateWindow()
        window.raise_()
        
        logger.debug("Main window displayed and activated")
        
        logger.debug("Entering main event loop")
        sys.exit(app.exec())
    except Exception as e:
        logger.error(f"Critical error during startup: {str(e)}", exc_info=True)
        raise

if __name__ == '__main__':
    logger.debug("Script started")
    main() 