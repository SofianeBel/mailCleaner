"""
Preview Dialog for Mail Cleaner
Provides a structured view of emails to be processed with granular control and confirmation.
"""

from typing import Dict, List, Tuple, Optional
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
    QTreeWidget, QTreeWidgetItem, QWidget, QSplitter,
    QTextEdit, QPushButton, QMessageBox, QComboBox,
    QFrame, QScrollArea
)
from PyQt6.QtCore import Qt, pyqtSignal
import logging

logger = logging.getLogger(__name__)

class EmailActionWidget(QWidget):
    """Custom widget for email items with action selection"""
    actionChanged = pyqtSignal(str)
    
    def __init__(self, email_data: Tuple, parent=None):
        super().__init__(parent)
        self.email_data = email_data
        self.setObjectName("emailActionWidget")
        self.initUI()
    
    def initUI(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)
        
        # Email info
        subject, received_time, size, _ = self.email_data
        info_label = QLabel(f"{subject} - {received_time} ({size})")
        info_label.setWordWrap(True)
        layout.addWidget(info_label, stretch=1)
        
        # Action selector
        self.action_combo = QComboBox()
        self.action_combo.addItems(["Archive", "Delete"])
        self.action_combo.currentTextChanged.connect(self.on_action_changed)
        layout.addWidget(self.action_combo)
    
    def on_action_changed(self, action: str):
        """Update combo box style based on action"""
        if action == "Delete":
            self.action_combo.setStyleSheet("""
                QComboBox {
                    background-color: #FFEBEE;
                    color: #D32F2F;
                    border-color: #FFCDD2;
                }
            """)
        else:
            self.action_combo.setStyleSheet("")
        self.actionChanged.emit(action)

class PreviewDialog(QDialog):
    """Dialog to preview emails to be processed with selection and confirmation"""
    def __init__(self, emails_by_recipient: Dict[str, List[Tuple]], parent=None):
        super().__init__(parent)
        self.emails_by_recipient = emails_by_recipient
        self.selected_emails = {}  # {recipient: [(email_data, action), ...]}
        self.setWindowTitle("Preview Emails")
        self.setMinimumSize(1000, 700)
        self.initUI()
    
    def initUI(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(15, 15, 15, 15)
        
        # Header frame
        header_frame = QFrame()
        header_frame.setObjectName("headerFrame")
        header_layout = QVBoxLayout(header_frame)
        
        # Title and description
        title_label = QLabel("Email Preview")
        title_label.setObjectName("headerLabel")
        header_layout.addWidget(title_label)
        
        desc_label = QLabel(
            "Select the emails you want to process and choose an action for each. "
            "You can preview the content of each email by clicking on it."
        )
        desc_label.setWordWrap(True)
        header_layout.addWidget(desc_label)
        
        layout.addWidget(header_frame)
        
        # Summary frame
        summary_frame = QFrame()
        summary_frame.setObjectName("contentFrame")
        summary_layout = QHBoxLayout(summary_frame)
        
        # Calculate totals
        total_emails = sum(len(emails) for emails in self.emails_by_recipient.values())
        total_size = 0.0
        for emails in self.emails_by_recipient.values():
            for email in emails:
                try:
                    size_value = float(email[2].split()[0])  # Extract number from "X.X KB"
                    total_size += size_value
                except:
                    pass
        
        summary_label = QLabel(f"Total: {total_emails} emails ({total_size:.1f} KB)")
        summary_label.setObjectName("summaryLabel")
        summary_layout.addWidget(summary_label)
        
        # Selection counter
        self.selection_label = QLabel("Selected: 0 emails")
        summary_layout.addWidget(self.selection_label)
        
        layout.addWidget(summary_frame)
        
        # Main content area with splitter
        content_frame = QFrame()
        content_frame.setObjectName("contentFrame")
        content_layout = QVBoxLayout(content_frame)
        content_layout.setContentsMargins(0, 0, 0, 0)
        
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Left panel: Email tree
        tree_container = QWidget()
        tree_layout = QVBoxLayout(tree_container)
        tree_layout.setContentsMargins(10, 10, 10, 10)
        
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Emails to Process"])
        self.tree.setColumnCount(1)
        self.tree.setMinimumWidth(400)
        
        # Populate tree
        for recipient, emails in self.emails_by_recipient.items():
            recipient_item = QTreeWidgetItem([f"{recipient} ({len(emails)} emails)"])
            recipient_item.setFlags(recipient_item.flags() & ~Qt.ItemFlag.ItemIsUserCheckable)
            self.tree.addTopLevelItem(recipient_item)
            
            for email in emails:
                email_item = QTreeWidgetItem([""])
                email_item.setFlags(
                    email_item.flags() | 
                    Qt.ItemFlag.ItemIsUserCheckable | 
                    Qt.ItemFlag.ItemIsEnabled
                )
                email_item.setCheckState(0, Qt.CheckState.Unchecked)
                recipient_item.addChild(email_item)
                
                # Create and set custom widget with action selector
                action_widget = EmailActionWidget(email)
                self.tree.setItemWidget(email_item, 0, action_widget)
                
                # Store email data
                email_item.setData(0, Qt.ItemDataRole.UserRole, email)
        
        tree_layout.addWidget(self.tree)
        
        # Quick selection buttons
        selection_layout = QHBoxLayout()
        select_all_btn = QPushButton("Select All")
        select_all_btn.setObjectName("secondaryButton")
        select_all_btn.clicked.connect(self.select_all_emails)
        
        select_none_btn = QPushButton("Select None")
        select_none_btn.setObjectName("secondaryButton")
        select_none_btn.clicked.connect(self.select_no_emails)
        
        selection_layout.addWidget(select_all_btn)
        selection_layout.addWidget(select_none_btn)
        selection_layout.addStretch()
        
        tree_layout.addLayout(selection_layout)
        
        splitter.addWidget(tree_container)
        
        # Right panel: Preview area
        preview_container = QWidget()
        preview_layout = QVBoxLayout(preview_container)
        preview_layout.setContentsMargins(10, 10, 10, 10)
        
        # Preview header
        preview_header = QFrame()
        preview_header.setObjectName("headerFrame")
        preview_header_layout = QVBoxLayout(preview_header)
        
        self.preview_subject = QLabel()
        self.preview_subject.setObjectName("headerLabel")
        preview_header_layout.addWidget(self.preview_subject)
        
        self.preview_info = QLabel()
        preview_header_layout.addWidget(self.preview_info)
        
        preview_layout.addWidget(preview_header)
        
        # Preview content in a scroll area
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        
        preview_content = QWidget()
        preview_content_layout = QVBoxLayout(preview_content)
        
        self.preview_text = QTextEdit()
        self.preview_text.setObjectName("previewArea")
        self.preview_text.setReadOnly(True)
        preview_content_layout.addWidget(self.preview_text)
        
        scroll_area.setWidget(preview_content)
        preview_layout.addWidget(scroll_area)
        
        splitter.addWidget(preview_container)
        
        # Set splitter proportions
        splitter.setStretchFactor(0, 1)  # Tree gets 1/3
        splitter.setStretchFactor(1, 2)  # Preview gets 2/3
        
        content_layout.addWidget(splitter)
        layout.addWidget(content_frame)
        
        # Bottom button area
        button_frame = QFrame()
        button_layout = QHBoxLayout(button_frame)
        button_layout.setContentsMargins(0, 0, 0, 0)
        
        # Process button
        self.process_btn = QPushButton("Process Selected")
        self.process_btn.clicked.connect(self.confirm_selection)
        
        # Cancel button
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setObjectName("secondaryButton")
        cancel_btn.clicked.connect(self.reject)
        
        button_layout.addStretch()
        button_layout.addWidget(cancel_btn)
        button_layout.addWidget(self.process_btn)
        
        layout.addWidget(button_frame)
        
        # Connect signals
        self.tree.itemClicked.connect(self.show_email_details)
        self.tree.itemChanged.connect(self.update_selection_count)
        
        # Initial state
        self.tree.expandAll()
        self.update_selection_count()
    
    def show_email_details(self, item: QTreeWidgetItem, column: int):
        """Show email details in preview pane when an email item is clicked"""
        if item.parent():  # This is an email item
            email_data = item.data(0, Qt.ItemDataRole.UserRole)
            if email_data:
                subject, received_time, size, body = email_data
                
                # Update preview header
                self.preview_subject.setText(subject)
                self.preview_info.setText(f"Date: {received_time}  |  Size: {size}")
                
                # Update preview content with styled HTML
                preview_html = f"""
                <div style="font-family: 'Segoe UI', sans-serif; line-height: 1.6;">
                    <div style="white-space: pre-wrap;">{body}</div>
                </div>
                """
                self.preview_text.setHtml(preview_html)
        else:
            # Clear preview for recipient items
            self.preview_subject.clear()
            self.preview_info.clear()
            self.preview_text.clear()
    
    def update_selection_count(self):
        """Update the selection counter and process button state"""
        count = 0
        for i in range(self.tree.topLevelItemCount()):
            recipient_item = self.tree.topLevelItem(i)
            for j in range(recipient_item.childCount()):
                email_item = recipient_item.child(j)
                if email_item.checkState(0) == Qt.CheckState.Checked:
                    count += 1
        
        self.selection_label.setText(f"Selected: {count} emails")
        self.process_btn.setEnabled(count > 0)
    
    def select_all_emails(self):
        """Select all email items in the tree"""
        self._set_all_check_states(Qt.CheckState.Checked)
    
    def select_no_emails(self):
        """Deselect all email items in the tree"""
        self._set_all_check_states(Qt.CheckState.Unchecked)
    
    def _set_all_check_states(self, state: Qt.CheckState):
        """Helper to set all email items to the given check state"""
        for i in range(self.tree.topLevelItemCount()):
            recipient_item = self.tree.topLevelItem(i)
            for j in range(recipient_item.childCount()):
                email_item = recipient_item.child(j)
                email_item.setCheckState(0, state)
    
    def confirm_selection(self):
        """Gather selected emails and their actions, then confirm with user"""
        self.selected_emails.clear()
        total_selected = 0
        delete_count = 0
        archive_count = 0
        
        # Collect selected emails and their actions
        for i in range(self.tree.topLevelItemCount()):
            recipient_item = self.tree.topLevelItem(i)
            recipient = recipient_item.text(0).split(" (")[0]  # Remove email count
            
            for j in range(recipient_item.childCount()):
                email_item = recipient_item.child(j)
                if email_item.checkState(0) == Qt.CheckState.Checked:
                    email_data = email_item.data(0, Qt.ItemDataRole.UserRole)
                    action_widget = self.tree.itemWidget(email_item, 0)
                    action = action_widget.action_combo.currentText()
                    
                    if recipient not in self.selected_emails:
                        self.selected_emails[recipient] = []
                    self.selected_emails[recipient].append((email_data, action))
                    total_selected += 1
                    
                    if action == "Delete":
                        delete_count += 1
                    else:
                        archive_count += 1
        
        if total_selected == 0:
            QMessageBox.warning(
                self,
                "No Selection",
                "Please select at least one email to process."
            )
            return
        
        # Show confirmation dialog with details
        msg = (
            f"You have selected {total_selected} emails to process:\n\n"
            f"• {delete_count} emails to delete\n"
            f"• {archive_count} emails to archive\n\n"
            "This action cannot be undone. Are you sure you want to continue?"
        )
        
        confirm = QMessageBox.question(
            self,
            "Confirm Processing",
            msg,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if confirm == QMessageBox.StandardButton.Yes:
            self.accept()
    
    def get_selected_emails(self) -> Dict[str, List[Tuple]]:
        """Return the dictionary of selected emails with their actions"""
        return self.selected_emails 