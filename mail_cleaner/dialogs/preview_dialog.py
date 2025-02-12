from typing import Dict, List, Tuple
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QTreeWidget, QTreeWidgetItem, QWidget, QSplitter,
    QTextEdit
)
from PyQt6.QtCore import Qt

class PreviewDialog(QDialog):
    """Dialog to preview emails to be processed"""
    def __init__(self, emails_by_recipient: Dict[str, List[Tuple]], parent=None):
        super().__init__(parent)
        self.emails_by_recipient = emails_by_recipient
        self.setWindowTitle("Preview Emails")
        self.setMinimumSize(800, 600)
        self.initUI()
    
    def initUI(self):
        layout = QVBoxLayout(self)
        
        # Create splitter for tree and preview
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Email tree
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Emails to Process"])
        self.tree.setColumnCount(1)
        self.tree.setMinimumWidth(300)
        
        # Populate tree
        total_emails = 0
        total_size = 0.0
        
        for recipient, emails in self.emails_by_recipient.items():
            recipient_item = QTreeWidgetItem([f"{recipient} ({len(emails)} emails)"])
            self.tree.addTopLevelItem(recipient_item)
            
            for email in emails:
                subject, received_time, size, body = email
                # Extract size value for total calculation
                try:
                    size_value = float(size.split()[0])  # Extract number from "X.X KB"
                    total_size += size_value
                except:
                    pass
                
                email_item = QTreeWidgetItem([f"{subject} - {received_time} ({size})"])
                recipient_item.addChild(email_item)
                total_emails += 1
                
                # Store email data for preview
                email_item.setData(0, Qt.ItemDataRole.UserRole, email)
        
        # Add summary at the top
        summary_label = QLabel(
            f"Total: {total_emails} emails "
            f"({total_size:.1f} KB)"
        )
        summary_label.setStyleSheet("font-weight: bold; padding: 5px;")
        layout.addWidget(summary_label)
        
        # Preview area
        preview_widget = QWidget()
        preview_layout = QVBoxLayout(preview_widget)
        
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        preview_layout.addWidget(self.preview_text)
        
        # Add widgets to splitter
        splitter.addWidget(self.tree)
        splitter.addWidget(preview_widget)
        
        # Set splitter proportions
        splitter.setStretchFactor(0, 1)  # Tree gets 1/3
        splitter.setStretchFactor(1, 2)  # Preview gets 2/3
        
        layout.addWidget(splitter)
        
        # Connect signals
        self.tree.itemClicked.connect(self.show_email_details)
        
        # Expand all items
        self.tree.expandAll()
    
    def show_email_details(self, item: QTreeWidgetItem, column: int):
        """Show email details in preview pane when an email item is clicked"""
        # Only show details for email items (not recipient items)
        if item.parent():  # This is an email item
            email_data = item.data(0, Qt.ItemDataRole.UserRole)
            if email_data:
                subject, received_time, size, body = email_data
                
                # Format the preview
                preview_html = f"""
                <h3>{subject}</h3>
                <p><b>Date:</b> {received_time}</p>
                <p><b>Size:</b> {size}</p>
                <hr>
                <div style="white-space: pre-wrap;">{body}</div>
                """
                
                self.preview_text.setHtml(preview_html)
        else:
            # Clear preview for recipient items
            self.preview_text.clear() 