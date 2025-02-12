"""
Account Selection Dialog
Provides a user interface for selecting which Outlook accounts to process.
"""

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QListWidget, QDialogButtonBox, 
    QLabel, QFrame, QListWidgetItem, QWidget, QCheckBox
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QFont, QIcon
from typing import List, Dict
import logging

logger = logging.getLogger(__name__)

class AccountListItem(QListWidgetItem):
    """Custom list widget item for displaying account information"""
    def __init__(self, account_info: Dict):
        super().__init__()
        self.account_info = account_info
        self.setText(account_info['display_name'])
        # Set proper flags to ensure the item is enabled and user-checkable
        self.setFlags(
            Qt.ItemFlag.ItemIsEnabled |
            Qt.ItemFlag.ItemIsUserCheckable |
            Qt.ItemFlag.ItemIsSelectable
        )
        self.setCheckState(Qt.CheckState.Unchecked)
        logger.debug(f"Created AccountListItem for {account_info['display_name']}")

class ClickableWidget(QWidget):
    """A widget that can be clicked and handles mouse events"""
    clicked = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMouseTracking(True)
    
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            clicked_widget = self.childAt(event.position().toPoint())
            if clicked_widget and isinstance(clicked_widget, QCheckBox):
                # When clicking directly on the checkbox, let it handle its own event
                super().mousePressEvent(event)
            else:
                # Only emit clicked signal for non-checkbox clicks
                self.clicked.emit()
                event.accept()
        else:
            super().mousePressEvent(event)

class AccountItemWidget(QWidget):
    """Custom widget for displaying detailed account information"""
    def __init__(self, account_info: Dict, list_item: AccountListItem, parent=None):
        super().__init__(parent)
        self.account_info = account_info
        self.list_item = list_item
        self.setup_ui()
        logger.debug(f"Created AccountItemWidget for {account_info['display_name']}")
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(2)
        
        # Create a clickable container widget
        container = ClickableWidget()
        container.clicked.connect(self.toggle_selection)
        container_layout = QVBoxLayout(container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.setSpacing(2)
        
        # Header layout with checkbox and name
        header_layout = QHBoxLayout()
        
        # Checkbox for selection
        self.checkbox = QCheckBox()
        self.checkbox.setChecked(self.list_item.checkState() == Qt.CheckState.Checked)
        self.checkbox.stateChanged.connect(self.on_checkbox_changed)
        header_layout.addWidget(self.checkbox)
        
        # Account name
        name_label = QLabel(self.account_info['display_name'])
        name_label.setFont(QFont("", weight=QFont.Weight.Bold))
        header_layout.addWidget(name_label)
        header_layout.addStretch()
        
        container_layout.addLayout(header_layout)
        
        # Details layout with increased left margin for better hierarchy
        details_layout = QVBoxLayout()
        details_layout.setContentsMargins(25, 0, 0, 0)  # Increased left margin
        
        # Email address
        if self.account_info['smtp_address']:
            email_label = QLabel(self.account_info['smtp_address'])
            email_label.setStyleSheet("color: #666;")
            details_layout.addWidget(email_label)
        
        # Account type and status
        status_layout = QHBoxLayout()
        
        # Account type badge
        type_label = QLabel(self.account_info['account_type'])
        type_label.setStyleSheet("""
            padding: 2px 6px;
            background-color: #e3f2fd;
            border-radius: 3px;
            color: #1976d2;
            font-size: 11px;
        """)
        status_layout.addWidget(type_label)
        
        # Primary account badge
        if self.account_info['is_primary']:
            primary_label = QLabel("Primary")
            primary_label.setStyleSheet("""
                padding: 2px 6px;
                background-color: #e8f5e9;
                border-radius: 3px;
                color: #2e7d32;
                font-size: 11px;
                margin-left: 5px;
            """)
            status_layout.addWidget(primary_label)
        
        # Gmail/Office 365 badge
        if self.account_info['is_gmail']:
            service_label = QLabel("Gmail")
            service_label.setStyleSheet("""
                padding: 2px 6px;
                background-color: #fce4ec;
                border-radius: 3px;
                color: #c2185b;
                font-size: 11px;
                margin-left: 5px;
            """)
            status_layout.addWidget(service_label)
        elif self.account_info['is_office365']:
            service_label = QLabel("Office 365")
            service_label.setStyleSheet("""
                padding: 2px 6px;
                background-color: #fff3e0;
                border-radius: 3px;
                color: #e65100;
                font-size: 11px;
                margin-left: 5px;
            """)
            status_layout.addWidget(service_label)
        
        status_layout.addStretch()
        details_layout.addLayout(status_layout)
        container_layout.addLayout(details_layout)
        
        layout.addWidget(container)
        
        # Set widget style
        self.setStyleSheet("""
            QWidget {
                background-color: transparent;
            }
            QCheckBox {
                margin-right: 5px;
            }
            QCheckBox:hover {
                background-color: #f5f5f5;
            }
        """)
    
    def toggle_selection(self):
        """Toggle the selection state"""
        logger.debug(f"Toggling selection for {self.account_info['display_name']}")
        # Block signals temporarily to prevent recursive signal emission
        self.checkbox.blockSignals(True)
        self.checkbox.setChecked(not self.checkbox.isChecked())
        self.checkbox.blockSignals(False)
        # Manually trigger the state change handler
        self.on_checkbox_changed(self.checkbox.checkState())
    
    def on_checkbox_changed(self, state):
        """Synchronize checkbox state with list item check state"""
        new_state = Qt.CheckState.Checked if state == Qt.CheckState.Checked else Qt.CheckState.Unchecked
        if self.list_item.checkState() != new_state:
            logger.debug(f"Checkbox state changed for {self.account_info['display_name']}: {state}")
            self.list_item.setCheckState(new_state)
            # Manually trigger the item changed signal since setItemWidget might interfere
            if isinstance(self.parent(), QListWidget):
                self.parent().itemChanged.emit(self.list_item)

class AccountSelectionDialog(QDialog):
    """Dialog for selecting which Outlook accounts to process"""
    def __init__(self, accounts: List[Dict], parent=None):
        super().__init__(parent)
        self.accounts = accounts
        self.selected_accounts = []
        logger.debug(f"Initializing AccountSelectionDialog with {len(accounts)} accounts")
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle("Select Accounts to Process")
        self.setMinimumWidth(450)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Info label
        info_label = QLabel(
            "Select the email accounts you want to process. "
            "Only emails from selected accounts will be affected."
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #666;")
        layout.addWidget(info_label)
        
        # Account list frame
        list_frame = QFrame()
        list_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #e0e0e0;
                border-radius: 8px;
            }
        """)
        list_layout = QVBoxLayout(list_frame)
        list_layout.setContentsMargins(0, 0, 0, 0)  # Remove frame padding
        
        self.account_list = QListWidget()
        self.account_list.setStyleSheet("""
            QListWidget {
                border: none;
                outline: none;
            }
            QListWidget::item {
                padding: 8px;
                border-bottom: 1px solid #f0f0f0;
            }
            QListWidget::item:last {
                border-bottom: none;
            }
            QListWidget::item:hover {
                background-color: #fafafa;
            }
        """)
        
        # Add accounts to list
        for account in self.accounts:
            item = AccountListItem(account)
            self.account_list.addItem(item)
            item_widget = AccountItemWidget(account, item, self.account_list)
            self.account_list.setItemWidget(item, item_widget)
        
        list_layout.addWidget(self.account_list)
        layout.addWidget(list_frame)
        
        # Warning for no selection
        self.warning_label = QLabel(
            "⚠ No accounts selected. All accessible folders will be processed."
        )
        self.warning_label.setStyleSheet("color: #f57c00;")
        self.warning_label.hide()
        layout.addWidget(self.warning_label)
        
        # Buttons
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | 
            QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        
        # Style the buttons
        for button in button_box.buttons():
            if button_box.buttonRole(button) == QDialogButtonBox.ButtonRole.AcceptRole:
                button.setStyleSheet("""
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
            else:
                button.setStyleSheet("""
                    QPushButton {
                        background-color: #f5f5f5;
                        color: #333;
                        border: 1px solid #ddd;
                        padding: 8px 16px;
                        border-radius: 4px;
                        min-width: 100px;
                    }
                    QPushButton:hover {
                        background-color: #e0e0e0;
                    }
                """)
        
        layout.addWidget(button_box)
        
        # Connect signals
        self.account_list.itemChanged.connect(self.on_selection_changed)
    
    def on_selection_changed(self, item):
        """Handle changes in account selection"""
        logger.debug("Selection changed")
        has_selection = False
        for i in range(self.account_list.count()):
            item = self.account_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                has_selection = True
                logger.debug(f"Account selected: {item.account_info['display_name']}")
                break
        self.warning_label.setVisible(not has_selection)
    
    def get_selected_accounts(self) -> List[str]:
        """Get the display names of selected accounts"""
        selected = []
        for i in range(self.account_list.count()):
            item = self.account_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                selected.append(item.account_info['display_name'])
                logger.debug(f"Account in selection: {item.account_info['display_name']}")
        logger.debug(f"Total accounts selected: {len(selected)}")
        return selected 