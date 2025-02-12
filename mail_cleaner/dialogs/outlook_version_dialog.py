from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFrame
)
from PyQt6.QtCore import Qt

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