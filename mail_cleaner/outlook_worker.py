import win32com.client
import logging
from datetime import datetime, timedelta
from typing import List, Dict, Tuple
from PyQt6.QtCore import QThread, pyqtSignal

logger = logging.getLogger(__name__)

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