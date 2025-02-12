from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QPushButton,
    QFrame, QButtonGroup, QRadioButton, QDialogButtonBox
)
from PyQt6.QtGui import QFont

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