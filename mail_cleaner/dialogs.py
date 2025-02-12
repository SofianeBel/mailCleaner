from typing import Dict, List, Tuple
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFrame, QButtonGroup, QRadioButton, QDialogButtonBox,
    QTextEdit, QTreeWidget, QTreeWidgetItem, QWidget, QSplitter
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

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