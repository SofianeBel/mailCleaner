import sys
import win32com.client
import logging
from datetime import datetime, timedelta
from typing import List, Optional, Tuple, Dict
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QSpinBox, QComboBox, QCheckBox, QPushButton, QTextEdit,
    QProgressBar, QMessageBox, QListWidget, QFrame, QTreeWidget,
    QTreeWidgetItem, QDialog, QDialogButtonBox, QSplitter, QRadioButton,
    QButtonGroup
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt6.QtGui import QFont, QIcon
import winreg
import psutil
import os

# Configure logging with high precision timestamps
logging.basicConfig(
    level=logging.DEBUG,  # Changed to DEBUG level
    format='%(asctime)s.%(msecs)03d %(levelname)s: %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.FileHandler('mail_cleaner.log'),  # Log to file
        logging.StreamHandler()  # Log to console
    ]
)
logger = logging.getLogger(__name__)

def detect_active_outlook() -> Optional[str]:
    """
    Detect which version of Outlook is currently running by examining processes.
    Returns: "new" for New Outlook, "classic" for Classic Outlook, or None if not found
    """
    try:
        for proc in psutil.process_iter(['name', 'exe']):
            try:
                proc_name = proc.info['name'].lower()
                if proc_name in ['outlook.exe', 'microsoft.office.outlook.hub.exe']:
                    exe_path = proc.info.get('exe', '')
                    if exe_path:
                        logger.debug(f"Found Outlook process: {exe_path}")
                        # Check if it's the new Outlook
                        if any(indicator in exe_path.lower() for indicator in ['windowsapps\\microsoft.outlookforwindows', 'msedge_shell.exe']):
                            logger.debug("Detected New Outlook as active")
                            return "new"
                        else:
                            logger.debug("Detected Classic Outlook as active")
                            return "classic"
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as e:
                logger.debug(f"Skipping process due to: {str(e)}")
                continue
        logger.debug("No running Outlook process found")
        return None
    except Exception as e:
        logger.error(f"Error detecting active Outlook: {str(e)}")
        return None

def check_outlook_version():
    """Check Outlook version and compatibility"""
    try:
        outlook_versions = []
        active_version = detect_active_outlook()
        logger.debug(f"Active Outlook version detected: {active_version}")
        
        # Try to detect both Classic and New Outlook
        try:
            # Check for Classic Outlook
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                classic_version = outlook.Version
                logger.debug(f"Found Classic Outlook version: {classic_version}")
                
                outlook_versions.append({
                    "version": classic_version,
                    "name": "Outlook (Classic)",
                    "is_modern": True,
                    "major_version": ".".join(classic_version.split(".")[:2]),
                    "path": "Classic Installation",
                    "bitness": "Desktop",
                    "is_active": active_version == "classic",
                    "display_name": "Outlook (Classic)",
                    "description": "Traditional desktop version of Outlook with full functionality"
                })
            except Exception as e:
                logger.debug(f"Could not detect Classic Outlook: {str(e)}")
            
            # Check for New Outlook using user-accessible registry paths
            new_outlook_detected = False
            user_registry_paths = [
                (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\16.0\Outlook\Options"),
                (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\16.0\Outlook\Preferences"),
                (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\outlook")
            ]
            
            for root_key, path in user_registry_paths:
                try:
                    key = winreg.OpenKey(root_key, path, 0, winreg.KEY_READ)
                    try:
                        # Check different possible registry values
                        for value_name in ["NewOutlook", "UseNewOutlook", "Microsoft.Office.Outlook.Hub.HubApp"]:
                            try:
                                value, _ = winreg.QueryValueEx(key, value_name)
                                if value:
                                    logger.debug(f"Found New Outlook indicator in registry: {value_name}={value}")
                                    new_outlook_detected = True
                                    break
                            except WindowsError:
                                continue
                    finally:
                        winreg.CloseKey(key)
                except WindowsError as e:
                    logger.debug(f"Could not access registry path {path}: {str(e)}")
                    continue
                
                if new_outlook_detected:
                    break
            
            # If we haven't detected New Outlook yet, try checking common installation paths
            if not new_outlook_detected and active_version == "new":
                logger.debug("New Outlook detected through process but not registry")
                new_outlook_detected = True
            
            if new_outlook_detected:
                outlook_versions.append({
                    "version": "16.0",  # New Outlook is always on latest version
                    "name": "Outlook (New)",
                    "is_modern": True,
                    "major_version": "16.0",
                    "path": "New Installation",
                    "bitness": "Modern",
                    "is_active": active_version == "new",
                    "display_name": "Outlook (New)",
                    "description": "Modern web-based version of Outlook with updated interface"
                })
                logger.debug("Added New Outlook to detected versions")
            
            if len(outlook_versions) > 0:
                # If no active version was detected but we found installations,
                # mark the first one as active
                if active_version is None and outlook_versions:
                    outlook_versions[0]["is_active"] = True
                    logger.debug(f"No active version detected, marking {outlook_versions[0]['name']} as active")
                return outlook_versions
            else:
                logger.error("No Outlook versions found")
                return None
                
        except Exception as e:
            logger.error(f"Error getting Outlook versions: {str(e)}")
            return None
            
    except Exception as e:
        logger.error(f"Error checking Outlook version: {str(e)}")
        return None

class OutlookVersionDialog(QDialog):
    """Dialog to show Outlook version information"""
    def __init__(self, version_info: dict, parent=None):
        super().__init__(parent)
        self.version_info = version_info
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle("Outlook Version Check")
        self.setMinimumWidth(400)
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Title
        title = QLabel("Outlook Version Information")
        title.setStyleSheet("font-size: 16px; font-weight: bold; color: #333333;")
        layout.addWidget(title)
        
        # Version info frame
        info_frame = QFrame()
        info_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                padding: 20px;
            }
            QLabel {
                color: #333333;
                font-size: 13px;
            }
        """)
        
        info_layout = QVBoxLayout(info_frame)
        info_layout.setSpacing(10)
        
        # Version name
        version_label = QLabel(f"Detected Version: {self.version_info['name']}")
        version_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        info_layout.addWidget(version_label)
        
        # Compatibility status
        compatibility_label = QLabel()
        if self.version_info['is_modern']:
            compatibility_label.setText("✓ Your Outlook version is fully compatible")
            compatibility_label.setStyleSheet("color: #2e7d32; font-weight: bold; font-size: 13px;")
        else:
            compatibility_label.setText("⚠ You are using an older version of Outlook")
            compatibility_label.setStyleSheet("color: #f57c00; font-weight: bold; font-size: 13px;")
        info_layout.addWidget(compatibility_label)
        
        # Description
        info_text = QLabel()
        if self.version_info['is_modern']:
            info_text.setText(
                "This version of Outlook supports all features of Mail Cleaner Pro. "
                "You can proceed with full functionality."
            )
        else:
            info_text.setText(
                "While Mail Cleaner Pro will work with your version of Outlook, "
                "some features might be limited or slower. Consider upgrading Outlook "
                "for the best experience."
            )
        info_text.setWordWrap(True)
        info_text.setStyleSheet("color: #666666;")
        info_layout.addWidget(info_text)
        
        layout.addWidget(info_frame)
        
        # Buttons
        button_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        ok_button.setStyleSheet("""
            QPushButton {
                background-color: #0078D4;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                min-width: 100px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #106EBE;
            }
        """)
        ok_button.clicked.connect(self.accept)
        button_layout.addStretch()
        button_layout.addWidget(ok_button)
        
        layout.addLayout(button_layout)

class OutlookVersionSelectionDialog(QDialog):
    def __init__(self, outlook_versions, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Choose Your Outlook Version")
        self.setMinimumWidth(500)
        self.outlook_versions = outlook_versions
        self.selected_version = None
        
        # Main layout
        layout = QVBoxLayout(self)
        
        # Info label at the top
        info_label = QLabel("Multiple Outlook versions detected. Please select the version you want to use:")
        info_label.setWordWrap(True)
        info_label.setStyleSheet("font-size: 10pt; margin-bottom: 10px;")
        layout.addWidget(info_label)
        
        # Create a button group for radio buttons
        self.button_group = QButtonGroup(self)
        
        # Add version options
        for version_info in outlook_versions:
            # Create a frame for each version
            version_frame = QFrame()
            version_frame.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Raised)
            version_frame.setLineWidth(1)
            
            # Set frame style based on active status
            if version_info["is_active"]:
                version_frame.setStyleSheet("""
                    QFrame {
                        background-color: #e3f2fd;
                        border: 2px solid #2196f3;
                        border-radius: 5px;
                        padding: 10px;
                        margin: 5px;
                    }
                """)
            else:
                version_frame.setStyleSheet("""
                    QFrame {
                        background-color: #f5f5f5;
                        border: 1px solid #ddd;
                        border-radius: 5px;
                        padding: 10px;
                        margin: 5px;
                    }
                """)
            
            # Frame layout
            frame_layout = QVBoxLayout(version_frame)
            
            # Radio button with version name
            radio = QRadioButton(version_info["display_name"])
            radio.setFont(QFont("", 10, QFont.Weight.Bold))
            self.button_group.addButton(radio)
            frame_layout.addWidget(radio)
            
            # Version details
            details_layout = QVBoxLayout()
            details_layout.setContentsMargins(20, 0, 0, 0)
            
            # Version info
            version_label = QLabel(f"Version: {version_info['version']}")
            version_label.setStyleSheet("color: #666;")
            details_layout.addWidget(version_label)
            
            # Description
            desc_label = QLabel(version_info["description"])
            desc_label.setWordWrap(True)
            desc_label.setStyleSheet("color: #666;")
            details_layout.addWidget(desc_label)
            
            # Active status
            if version_info["is_active"]:
                active_label = QLabel("✓ Currently Active")
                active_label.setStyleSheet("color: #2196f3; font-weight: bold;")
                details_layout.addWidget(active_label)
            
            frame_layout.addLayout(details_layout)
            layout.addWidget(version_frame)
            
            # Set the active version as default selection
            if version_info["is_active"]:
                radio.setChecked(True)
                self.selected_version = version_info
        
        # Warning label for switching from active version
        self.warning_label = QLabel()
        self.warning_label.setWordWrap(True)
        self.warning_label.setStyleSheet("color: #f44336; margin-top: 10px;")
        self.warning_label.hide()
        layout.addWidget(self.warning_label)
        
        # Note about selection
        note_label = QLabel("Note: This selection will be used for all operations in Mail Cleaner Pro")
        note_label.setWordWrap(True)
        note_label.setStyleSheet("color: #666; font-style: italic; margin-top: 10px;")
        layout.addWidget(note_label)
        
        # Button box
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        
        # Customize the OK button
        ok_button = button_box.button(QDialogButtonBox.StandardButton.Ok)
        ok_button.setText("Use Selected Version")
        ok_button.setStyleSheet("""
            QPushButton {
                background-color: #2196f3;
                color: white;
                padding: 6px 12px;
                border-radius: 4px;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #1976d2;
            }
        """)
        
        layout.addWidget(button_box)
        
        # Connect radio button changes
        self.button_group.buttonClicked.connect(self.on_version_selected)
    
    def on_version_selected(self, radio):
        # Find the corresponding version info
        for version_info in self.outlook_versions:
            if version_info["display_name"] == radio.text():
                self.selected_version = version_info
                
                # Show warning if switching from active version
                active_version = next((v for v in self.outlook_versions if v["is_active"]), None)
                if active_version and active_version != version_info:
                    self.warning_label.setText(
                        f"Warning: You are switching from the currently active {active_version['display_name']}. "
                        "This may affect the application's behavior."
                    )
                    self.warning_label.show()
                else:
                    self.warning_label.hide()
                break
    
    def get_selected_version(self):
        return self.selected_version

class PreviewDialog(QDialog):
    """Dialog for previewing emails before deletion/archiving"""
    def __init__(self, emails_by_recipient: Dict[str, List[Tuple]], parent=None):
        super().__init__(parent)
        self.emails_by_recipient = emails_by_recipient
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle("Preview Affected Emails")
        self.setMinimumSize(800, 600)
        
        layout = QVBoxLayout(self)
        
        # Create splitter for tree and details
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Email tree widget
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Recipient/Subject", "Date", "Size"])
        self.tree.setColumnWidth(0, 400)
        self.tree.itemClicked.connect(self.show_email_details)
        
        # Populate tree
        total_emails = 0
        for recipient, emails in self.emails_by_recipient.items():
            recipient_item = QTreeWidgetItem([recipient, f"{len(emails)} emails"])
            self.tree.addTopLevelItem(recipient_item)
            
            for email in emails:
                subject, date, size, _ = email
                email_item = QTreeWidgetItem([subject, date, size])
                recipient_item.addChild(email_item)
                total_emails += 1
        
        # Email details widget
        details_widget = QWidget()
        details_layout = QVBoxLayout(details_widget)
        
        self.details_text = QTextEdit()
        self.details_text.setReadOnly(True)
        details_layout.addWidget(self.details_text)
        
        # Add widgets to splitter
        splitter.addWidget(self.tree)
        splitter.addWidget(details_widget)
        
        # Add splitter to main layout
        layout.addWidget(splitter)
        
        # Warning label
        warning_label = QLabel(
            f"<b>Warning:</b> This operation will affect {total_emails} emails. "
            "This action cannot be undone!"
        )
        warning_label.setStyleSheet("color: red;")
        layout.addWidget(warning_label)
        
        # Dialog buttons
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | 
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
    def show_email_details(self, item: QTreeWidgetItem, column: int):
        """Show details of selected email"""
        if item.parent():  # This is an email item
            recipient = item.parent().text(0)
            subject = item.text(0)
            
            # Find the email data
            for email in self.emails_by_recipient[recipient]:
                if email[0] == subject:  # Match subject
                    _, date, size, body = email
                    self.details_text.setHtml(
                        f"<h3>{subject}</h3>"
                        f"<p><b>Date:</b> {date}</p>"
                        f"<p><b>Size:</b> {size}</p>"
                        f"<p><b>Recipient:</b> {recipient}</p>"
                        f"<hr>"
                        f"{body}"
                    )
                    break

class OutlookWorker(QThread):
    """Worker thread for Outlook operations to prevent GUI freezing"""
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal(tuple)
    preview_data = pyqtSignal(dict)
    
    def __init__(self, age: int, unit: str, folders: List[str], action: str, 
                 preview_mode: bool = False, outlook_version: dict = None):
        super().__init__()
        self.age = age
        self.unit = unit
        self.folders = folders
        self.action = action
        self.preview_mode = preview_mode
        self.outlook_version = outlook_version
        self.is_running = True
        self.batch_size = 100 if outlook_version and outlook_version['is_modern'] else 20

    def run(self):
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            
            # Calculate threshold date
            multiplier = 365 if self.unit == "Years" else 30
            threshold_date = datetime.now() - timedelta(days=self.age * multiplier)
            
            total_processed = 0
            total_affected = 0
            emails_by_recipient = {}
            
            for folder_name in self.folders:
                if not self.is_running:
                    break
                    
                try:
                    folder = namespace.GetDefaultFolder(6) if folder_name == "Inbox" else \
                            namespace.GetDefaultFolder(5) if folder_name == "Sent Items" else \
                            None
                    
                    if not folder:
                        self.status.emit(f"Skipping folder {folder_name} - not found")
                        continue
                    
                    # Get items and sort by date if using modern Outlook
                    items = folder.Items
                    if self.outlook_version and self.outlook_version['is_modern']:
                        try:
                            items.Sort("[ReceivedTime]", True)  # Descending order
                            self.status.emit("✓ Using optimized date sorting")
                        except:
                            self.status.emit("⚠ Date sorting not available")
                    
                    item_count = items.Count
                    processed_in_batch = 0
                    
                    for i in range(item_count):
                        if not self.is_running:
                            break
                            
                        try:
                            item = items[i]
                            received_time = item.ReceivedTime if hasattr(item, 'ReceivedTime') else \
                                          item.SentOn if hasattr(item, 'SentOn') else None
                                          
                            if received_time and datetime.fromtimestamp(received_time.timestamp()) < threshold_date:
                                if self.preview_mode:
                                    # Get recipient(s)
                                    recipients = []
                                    if hasattr(item, 'To'):
                                        recipients.extend(str(item.To).split(';'))
                                    if hasattr(item, 'CC'):
                                        recipients.extend(str(item.CC).split(';'))
                                    
                                    recipient_key = '; '.join(filter(None, recipients)) or 'No Recipient'
                                    
                                    # Get email size
                                    size = f"{item.Size / 1024:.1f} KB" if hasattr(item, 'Size') else "Unknown"
                                    
                                    # Get email body (safely)
                                    try:
                                        body = item.HTMLBody if hasattr(item, 'HTMLBody') else \
                                              item.Body if hasattr(item, 'Body') else \
                                              "No content available"
                                    except:
                                        body = "Content unavailable"
                                    
                                    email_data = (
                                        item.Subject if hasattr(item, 'Subject') else 'No Subject',
                                        received_time.strftime("%Y-%m-%d %H:%M"),
                                        size,
                                        body
                                    )
                                    
                                    if recipient_key not in emails_by_recipient:
                                        emails_by_recipient[recipient_key] = []
                                    emails_by_recipient[recipient_key].append(email_data)
                                    
                                elif self.action == "Delete":
                                    item.Delete()
                                else:  # Archive
                                    # Implementation for archiving would go here
                                    pass
                                total_affected += 1
                                
                                # Batch processing for older Outlook versions
                                processed_in_batch += 1
                                if processed_in_batch >= self.batch_size:
                                    if not self.outlook_version or not self.outlook_version['is_modern']:
                                        self.status.emit(f"Processing in batches of {self.batch_size} for compatibility...")
                                    processed_in_batch = 0
                                
                            total_processed += 1
                            progress = int((total_processed / item_count) * 100)
                            self.progress.emit(progress)
                            
                        except Exception as e:
                            logger.error(f"Error processing item: {str(e)}")
                            continue
                            
                except Exception as e:
                    logger.error(f"Error accessing folder {folder_name}: {str(e)}")
                    self.status.emit(f"Error in folder {folder_name}: {str(e)}")
                    continue
            
            if self.preview_mode:
                self.preview_data.emit(emails_by_recipient)
            self.finished.emit((total_processed, total_affected))
            
        except Exception as e:
            logger.error(f"Critical error in Outlook operations: {str(e)}")
            self.status.emit(f"Critical error: {str(e)}")
            self.finished.emit((0, 0))

    def stop(self):
        self.is_running = False


class MailCleanerUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.outlook_version = None
        if not self.check_outlook_compatibility():
            sys.exit(1)
        self.initUI()

    def check_outlook_compatibility(self):
        """Check Outlook version and show compatibility dialog"""
        versions = check_outlook_version()
        if versions is None:
            QMessageBox.critical(
                self,
                "Error",
                "Could not detect Outlook. Please make sure Microsoft Outlook is installed and try again."
            )
            return False
            
        if isinstance(versions, list):
            # Multiple versions found
            dialog = OutlookVersionSelectionDialog(versions, self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.outlook_version = dialog.get_selected_version()
            else:
                return False
        else:
            # Single version found
            self.outlook_version = versions
            dialog = OutlookVersionDialog(versions, self)
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
        layout.addWidget(age_frame)
        layout.addWidget(folders_frame)
        layout.addWidget(action_frame)
        layout.addWidget(progress_frame)
        layout.addWidget(buttons_frame)
        
        self.log_message("Application ready.")

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
            outlook_version=self.outlook_version
        )
        
        self.worker.progress.connect(self.update_progress)
        self.worker.status.connect(self.log_message)
        self.worker.preview_data.connect(self.show_preview_dialog)
        self.worker.finished.connect(self.preview_finished)
        
        self.worker.start()
        self.log_message("Generating preview...")

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
            confirm_msg = (
                f"This will {self.action_combo.currentText().lower()} emails older than "
                f"{self.age_spinbox.value()} {self.unit_combo.currentText().lower()} "
                f"in the selected folders. It is recommended to preview first. Continue?"
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
            outlook_version=self.outlook_version
        )
        
        self.worker.progress.connect(self.update_progress)
        self.worker.status.connect(self.log_message)
        self.worker.finished.connect(self.process_finished)
        
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