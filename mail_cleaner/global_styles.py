"""
Global styles for Mail Cleaner
Provides consistent styling across all dialogs and windows.
"""

global_styles = """
/* Base styles */
QMainWindow, QDialog {
    background-color: #ECEFF1;
    color: #263238;
}

QWidget {
    font-family: 'Segoe UI', sans-serif;
    font-size: 12pt;
}

QLabel {
    color: #263238;
}

/* Buttons */
QPushButton {
    background-color: #1976D2;
    color: #FFFFFF;
    border: none;
    padding: 8px 16px;
    border-radius: 4px;
    font-weight: bold;
}

QPushButton:hover {
    background-color: #1565C0;
}

QPushButton:disabled {
    background-color: #90A4AE;
}

QPushButton#destructiveButton {
    background-color: #D32F2F;
}

QPushButton#destructiveButton:hover {
    background-color: #C62828;
}

QPushButton#secondaryButton {
    background-color: #78909C;
}

QPushButton#secondaryButton:hover {
    background-color: #607D8B;
}

/* List and Tree Widgets */
QListWidget, QTreeWidget {
    background-color: #FFFFFF;
    color: #263238;
    border: 1px solid #B0BEC5;
    border-radius: 4px;
    padding: 4px;
}

QListWidget::item, QTreeWidget::item {
    padding: 10px;
    border-bottom: 1px solid #ECEFF1;
}

QListWidget::item:hover, QTreeWidget::item:hover {
    background-color: #ECEFF1;
}

QListWidget::item:selected, QTreeWidget::item:selected {
    background-color: #E3F2FD;
    color: #1976D2;
}

/* Input and Display Widgets */
QTextEdit, QComboBox {
    background-color: #FFFFFF;
    color: #263238;
    border: 1px solid #B0BEC5;
    border-radius: 4px;
    padding: 4px;
}

QComboBox:hover {
    border-color: #90A4AE;
}

QComboBox::drop-down {
    border: none;
    padding-right: 8px;
}

QComboBox::down-arrow {
    image: url(resources/down-arrow.png);
    width: 12px;
    height: 12px;
}

/* Frames and Containers */
QFrame#contentFrame {
    background-color: #FFFFFF;
    border: 1px solid #B0BEC5;
    border-radius: 8px;
}

QFrame#headerFrame {
    background-color: #FFFFFF;
    border-bottom: 1px solid #B0BEC5;
    padding: 10px;
}

/* Splitters */
QSplitter::handle {
    background-color: #B0BEC5;
}

QSplitter::handle:horizontal {
    width: 2px;
}

QSplitter::handle:vertical {
    height: 2px;
}

/* Scroll Bars */
QScrollBar:vertical {
    border: none;
    background-color: #ECEFF1;
    width: 10px;
    margin: 0px;
}

QScrollBar::handle:vertical {
    background-color: #90A4AE;
    border-radius: 5px;
    min-height: 20px;
}

QScrollBar::handle:vertical:hover {
    background-color: #78909C;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}

/* Custom Widgets */
QLabel#warningLabel {
    color: #D32F2F;
    font-weight: bold;
}

QLabel#summaryLabel {
    font-size: 14pt;
    font-weight: bold;
    color: #1976D2;
}

QLabel#headerLabel {
    font-size: 16pt;
    font-weight: bold;
    color: #263238;
}

/* Preview Area */
QTextEdit#previewArea {
    font-family: 'Segoe UI', sans-serif;
    font-size: 11pt;
    line-height: 1.4;
    padding: 12px;
}

/* Email Action Widget */
QWidget#emailActionWidget {
    background-color: transparent;
    padding: 4px;
}

QWidget#emailActionWidget QLabel {
    font-size: 11pt;
}

QWidget#emailActionWidget QComboBox {
    min-width: 100px;
    max-width: 120px;
}
""" 