"""
Mail Cleaner Pro - A high-performance Outlook email management tool
"""

__version__ = "1.0.0"
__author__ = "Your Name"
__license__ = "MIT"

from .outlook_detector import check_outlook_version
from .outlook_worker import OutlookWorker
from .dialogs import OutlookVersionDialog, OutlookVersionSelectionDialog, PreviewDialog
from .main_window import MailCleanerUI 