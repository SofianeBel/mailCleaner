"""
Mail Cleaner Pro - A high-performance Outlook email management tool
"""

__version__ = "1.0.0"
__author__ = "Your Name"
__license__ = "MIT"

from .core.outlook_detector import check_outlook_version
from .core.outlook_worker import OutlookWorker
from .core.outlook_accounts import get_outlook_accounts
from .dialogs.outlook_version_dialog import OutlookVersionDialog
from .dialogs.outlook_version_selection_dialog import OutlookVersionSelectionDialog
from .dialogs.preview_dialog import PreviewDialog
from .dialogs.account_selection_dialog import AccountSelectionDialog
from .main_window import MailCleanerUI

__all__ = [
    'check_outlook_version',
    'OutlookWorker',
    'get_outlook_accounts',
    'OutlookVersionDialog',
    'OutlookVersionSelectionDialog',
    'PreviewDialog',
    'AccountSelectionDialog',
    'MailCleanerUI'
] 