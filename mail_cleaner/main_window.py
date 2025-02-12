import sys
import logging
from datetime import datetime
from typing import List, Dict, Tuple
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QSpinBox, QComboBox, QPushButton, QTextEdit,
    QProgressBar, QMessageBox, QListWidget, QFrame,
    QDialog, QApplication
)
from PyQt6.QtCore import Qt
from .dialogs.outlook_version_dialog import OutlookVersionDialog
from .dialogs.outlook_version_selection_dialog import OutlookVersionSelectionDialog
from .dialogs.preview_dialog import PreviewDialog
from .dialogs.account_selection_dialog import AccountSelectionDialog
from .core.outlook_detector import check_outlook_version
from .core.outlook_worker import OutlookWorker
from .core.outlook_accounts import get_outlook_accounts

logger = logging.getLogger(__name__)

class MailCleanerUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.outlook_version = None
        self.selected_accounts = []
        if not self.check_outlook_compatibility():
            sys.exit(1)
        self.initUI()

    def check_outlook_compatibility(self):
        """Check Outlook version and show compatibility dialog"""
        versions = check_outlook_version()
        if not versions:
            QMessageBox.critical(
                self,
                "Error",
                "Could not detect Outlook. Please make sure Microsoft Outlook is installed and try again."
            )
            return False
            
        if len(versions) > 1:
            # Multiple versions found
            dialog = OutlookVersionSelectionDialog(versions, self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.outlook_version = dialog.get_selected_version()
            else:
                return False
        else:
            # Single version found
            self.outlook_version = versions[0]
            dialog = OutlookVersionDialog(versions[0], self)
            dialog.exec()
        
        if not self.outlook_version["is_modern"]:
            self.log_message(
                "⚠ Using older version of Outlook. Some operations might be slower."
            )
        
        return True

    def initUI(self):
        # Set window properties
        self.setWindowTitle(f'Mail Cleaner Pro - {self.outlook_version["name"]}')
        self.setMinimumWidth(800)
        self.setMinimumHeight(600)
        
        # Center the window on screen
        screen = QApplication.primaryScreen().geometry()
        self.setGeometry(
            (screen.width() - 800) // 2,
            (screen.height() - 600) // 2,
            800, 600
        )
        
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QFrame {
                background-color: white;
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                padding: 15px;
                margin: 5px;
            }
            QPushButton {
                background-color: #0078D4;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                min-width: 120px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #106EBE;
            }
            QPushButton:disabled {
                background-color: #CCE4F7;
                color: #666666;
            }
            QLabel {
                color: #333333;
                font-size: 12px;
                font-weight: bold;
            }
            QComboBox, QSpinBox {
                background-color: white;
                color: #333333;
                padding: 5px;
                border: 1px solid #cccccc;
                border-radius: 4px;
                min-width: 100px;
                min-height: 25px;
            }
            QComboBox::drop-down {
                border: none;
                background-color: transparent;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #666666;
                margin-right: 5px;
            }
            QListWidget {
                background-color: white;
                border: 1px solid #cccccc;
                border-radius: 4px;
                padding: 5px;
                color: #333333;
            }
            QListWidget::item {
                padding: 5px;
                border-radius: 2px;
            }
            QListWidget::item:selected {
                background-color: #e3f2fd;
                color: #333333;
            }
            QProgressBar {
                border: 1px solid #cccccc;
                border-radius: 4px;
                text-align: center;
                background-color: white;
                min-height: 20px;
            }
            QProgressBar::chunk {
                background-color: #0078D4;
                border-radius: 3px;
            }
            QTextEdit {
                background-color: white;
                border: 1px solid #cccccc;
                border-radius: 4px;
                padding: 5px;
                color: #333333;
            }
        """)
        
        # Create main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Account selection frame
        account_frame = QFrame()
        account_layout = QHBoxLayout(account_frame)
        
        account_label = QLabel("Email Accounts:")
        account_label.setStyleSheet("font-size: 13px;")
        
        self.account_button = QPushButton("Select Accounts...")
        self.account_button.clicked.connect(self.select_accounts)
        self.account_button.setStyleSheet("""
            QPushButton {
                background-color: #f5f5f5;
                color: #333;
                border: 1px solid #ddd;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
        """)
        
        self.account_status = QLabel("All accounts")
        self.account_status.setStyleSheet("color: #666; font-weight: normal;")
        
        account_layout.addWidget(account_label)
        account_layout.addWidget(self.account_button)
        account_layout.addWidget(self.account_status)
        account_layout.addStretch()
        
        # Age selection frame
        age_frame = QFrame()
        age_layout = QHBoxLayout(age_frame)
        
        age_label = QLabel("Process emails older than:")
        age_label.setStyleSheet("font-size: 13px;")
        
        self.age_spinbox = QSpinBox()
        self.age_spinbox.setRange(1, 100)
        self.age_spinbox.setValue(6)
        
        self.unit_combo = QComboBox()
        self.unit_combo.addItems(["Months", "Years"])
        
        age_layout.addWidget(age_label)
        age_layout.addWidget(self.age_spinbox)
        age_layout.addWidget(self.unit_combo)
        age_layout.addStretch()
        
        # Folders frame
        folders_frame = QFrame()
        folders_layout = QVBoxLayout(folders_frame)
        folder_label = QLabel("Select folders to process:")
        folder_label.setStyleSheet("font-size: 13px;")
        folders_layout.addWidget(folder_label)
        
        self.folder_list = QListWidget()
        self.folder_list.addItems(["Inbox", "Sent Items"])
        self.folder_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        folders_layout.addWidget(self.folder_list)
        
        # Action frame
        action_frame = QFrame()
        action_layout = QHBoxLayout(action_frame)
        
        action_label = QLabel("Action:")
        action_label.setStyleSheet("font-size: 13px;")
        
        self.action_combo = QComboBox()
        self.action_combo.addItems(["Delete", "Archive"])
        
        action_layout.addWidget(action_label)
        action_layout.addWidget(self.action_combo)
        action_layout.addStretch()
        
        # Progress frame
        progress_frame = QFrame()
        progress_layout = QVBoxLayout(progress_frame)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        self.status_text.setMaximumHeight(100)
        self.status_text.setStyleSheet("font-family: 'Consolas', monospace;")
        
        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.status_text)
        
        # Buttons frame
        buttons_frame = QFrame()
        buttons_layout = QHBoxLayout(buttons_frame)
        
        self.preview_button = QPushButton("Preview")
        self.preview_button.clicked.connect(self.preview_processing)
        
        self.start_button = QPushButton("Start Processing")
        self.start_button.clicked.connect(self.start_processing)
        
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.cancel_processing)
        self.cancel_button.setEnabled(False)
        self.cancel_button.setStyleSheet("""
            QPushButton {
                background-color: #d32f2f;
            }
            QPushButton:hover {
                background-color: #b71c1c;
            }
            QPushButton:disabled {
                background-color: #ffcdd2;
            }
        """)
        
        buttons_layout.addWidget(self.preview_button)
        buttons_layout.addWidget(self.start_button)
        buttons_layout.addWidget(self.cancel_button)
        buttons_layout.addStretch()
        
        # Add all frames to main layout
        layout.addWidget(account_frame)
        layout.addWidget(age_frame)
        layout.addWidget(folders_frame)
        layout.addWidget(action_frame)
        layout.addWidget(progress_frame)
        layout.addWidget(buttons_frame)
        
        self.log_message("Application ready.")

    def select_accounts(self):
        """Show account selection dialog"""
        accounts = get_outlook_accounts()
        if not accounts:
            self.log_message("No accounts detected. Processing default folders.")
            return
            
        dialog = AccountSelectionDialog([acc.to_dict() for acc in accounts], self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.selected_accounts = dialog.get_selected_accounts()
            if self.selected_accounts:
                self.account_status.setText(
                    f"{len(self.selected_accounts)} account{'s' if len(self.selected_accounts) > 1 else ''} selected"
                )
                self.log_message(f"Selected accounts: {', '.join(self.selected_accounts)}")
            else:
                self.account_status.setText("All accounts")
                self.log_message("No accounts selected, processing all accessible folders.")

    def log_message(self, message: str):
        """Add a message to the status text area with timestamp"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_text.append(f"[{timestamp}] {message}")

    def preview_processing(self):
        """Preview emails that will be affected"""
        selected_folders = [item.text() for item in self.folder_list.selectedItems()]
        if not selected_folders:
            QMessageBox.warning(self, "Warning", "Please select at least one folder to process.")
            return
            
        self.preview_button.setEnabled(False)
        self.start_button.setEnabled(False)
        self.progress_bar.setValue(0)
        
        self.worker = OutlookWorker(
            self.age_spinbox.value(),
            self.unit_combo.currentText(),
            selected_folders,
            self.action_combo.currentText(),
            preview_mode=True,
            outlook_version=self.outlook_version,
            selected_accounts=self.selected_accounts
        )
        
        self.worker.progress.connect(self.update_progress)
        self.worker.status.connect(self.log_message)
        self.worker.preview_data.connect(self.show_preview_dialog)
        self.worker.finished.connect(self.preview_finished)
        self.worker.account_detected.connect(self.on_account_detected)
        
        self.worker.start()
        self.log_message("Generating preview...")

    def on_account_detected(self, account_info: Dict):
        """Handle account detection from worker"""
        account_type = account_info['account_type']
        is_primary = account_info['is_primary']
        is_gmail = account_info['is_gmail']
        is_office365 = account_info['is_office365']
        
        status = []
        if is_primary:
            status.append("Primary")
        if is_gmail:
            status.append("Gmail")
        elif is_office365:
            status.append("Office 365")
        
        status_str = f" ({', '.join(status)})" if status else ""
        
        self.log_message(
            f"Detected {account_info['display_name']} "
            f"({account_type}){status_str}"
        )

    def show_preview_dialog(self, emails_by_recipient: Dict[str, List[Tuple]]):
        """Show the preview dialog with email data"""
        dialog = PreviewDialog(emails_by_recipient, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.start_processing(skip_confirmation=True)

    def preview_finished(self, results: Tuple[int, int]):
        """Handle completion of preview generation"""
        self.preview_button.setEnabled(True)
        self.start_button.setEnabled(True)
        self.progress_bar.setValue(0)

    def start_processing(self, skip_confirmation: bool = False):
        """Start the email processing operation"""
        selected_folders = [item.text() for item in self.folder_list.selectedItems()]
        if not selected_folders:
            QMessageBox.warning(self, "Warning", "Please select at least one folder to process.")
            return

        if not skip_confirmation:
            account_str = (
                f" in {len(self.selected_accounts)} selected account{'s' if len(self.selected_accounts) > 1 else ''}"
                if self.selected_accounts else " in all accounts"
            )
            
            confirm_msg = (
                f"This will {self.action_combo.currentText().lower()} emails older than "
                f"{self.age_spinbox.value()} {self.unit_combo.currentText().lower()} "
                f"in the selected folders{account_str}. "
                "It is recommended to preview first. Continue?"
            )
            
            reply = QMessageBox.question(
                self, 'Confirm Operation',
                confirm_msg,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            
            if reply != QMessageBox.StandardButton.Yes:
                return

        self.preview_button.setEnabled(False)
        self.start_button.setEnabled(False)
        self.cancel_button.setEnabled(True)
        self.progress_bar.setValue(0)
        
        self.worker = OutlookWorker(
            self.age_spinbox.value(),
            self.unit_combo.currentText(),
            selected_folders,
            self.action_combo.currentText(),
            outlook_version=self.outlook_version,
            selected_accounts=self.selected_accounts
        )
        
        self.worker.progress.connect(self.update_progress)
        self.worker.status.connect(self.log_message)
        self.worker.finished.connect(self.process_finished)
        self.worker.account_detected.connect(self.on_account_detected)
        
        self.worker.start()
        self.log_message("Processing started...")

    def cancel_processing(self):
        """Cancel the current processing operation"""
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.log_message("Cancelling operation...")

    def update_progress(self, value: int):
        """Update the progress bar value"""
        self.progress_bar.setValue(value)

    def process_finished(self, results: Tuple[int, int]):
        """Handle completion of the processing operation"""
        total_processed, total_affected = results
        self.start_button.setEnabled(True)
        self.cancel_button.setEnabled(False)
        self.progress_bar.setValue(100)
        
        self.log_message(
            f"Operation completed. Processed {total_processed} items, "
            f"{total_affected} items affected."
        ) 