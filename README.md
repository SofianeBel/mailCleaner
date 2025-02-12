# Mail Cleaner Pro

A high-performance Outlook email management tool built with Python and PyQt6. This application helps you efficiently manage your Outlook mailbox by providing automated cleanup capabilities with a user-friendly interface.

## Features

- **Advanced Email Processing**: Process emails based on age (months or years)
- **Multiple Folder Support**: Clean up Inbox, Sent Items, or both simultaneously
- **Flexible Actions**: Choose between deletion and archiving
- **Real-time Progress Tracking**: Monitor cleanup progress with a detailed status log
- **Cancellation Support**: Safely cancel operations mid-process
- **Error Handling**: Robust error handling with detailed logging
- **Modern UI**: Clean, intuitive interface built with PyQt6

## System Requirements

- Windows 10 or later
- Python 3.9 or later
- Microsoft Outlook (tested with Outlook 2016 and later)
- Administrator privileges (for installation)

## Installation

1. **Clone or download this repository**:
   ```bash
   git clone 
   cd mail-cleaner
   ```

2. **Create and activate a virtual environment** (recommended):
   ```bash
   python -m venv venv
   .\venv\Scripts\activate
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Build the executable** (optional):
   ```bash
   pyinstaller --onefile --windowed mail_cleaner.py
   ```
   The executable will be created in the `dist` directory.

## Usage

1. **Start the application**:
   - If running from source: `python mail_cleaner.py`
   - If using executable: Run `mail_cleaner.exe` from the `dist` directory

2. **Configure cleanup parameters**:
   - Set the age threshold (e.g., 6 months, 1 year)
   - Select the folders to process
   - Choose the action (Delete or Archive)

3. **Start processing**:
   - Click "Start Processing"
   - Confirm the operation when prompted
   - Monitor progress in the status window

4. **Cancel if needed**:
   - Click "Cancel" to safely stop the operation
   - The application will complete the current item before stopping

## Performance Considerations

- Processing speed depends on the size of your mailbox and system resources
- For large mailboxes (>10,000 items), the initial scan may take several minutes
- The application uses a separate thread for email processing to maintain UI responsiveness

## Security Notes

- The application interacts with Outlook through the official Microsoft Office COM interface
- No email data is stored or transmitted outside of Outlook
- All operations are performed locally on your machine

## Troubleshooting

1. **Outlook Not Found**:
   - Ensure Outlook is installed and properly configured
   - Try running Outlook manually first

2. **Permission Errors**:
   - Run the application as administrator
   - Check Outlook security settings

3. **Performance Issues**:
   - Close other applications using Outlook
   - Reduce the number of folders being processed simultaneously

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- PyQt6 team for the excellent GUI framework
- Microsoft for the Outlook COM interface
- All contributors and users of this project 