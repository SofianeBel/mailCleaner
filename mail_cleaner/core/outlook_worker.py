"""
Worker thread for Outlook operations to prevent GUI freezing.
Handles email processing operations in a separate thread.
"""

import win32com.client
from datetime import datetime, timedelta
from typing import List, Dict, Tuple, Optional
from PyQt6.QtCore import QThread, pyqtSignal
import logging
from .outlook_accounts import get_account_by_folder, get_outlook_accounts

logger = logging.getLogger(__name__)

# Outlook folder constants
FOLDER_INBOX = 6  # olFolderInbox
FOLDER_SENT = 5   # olFolderSentMail
FOLDER_DELETED = 3  # olFolderDeletedItems

# Mapping of folder types to their possible localized names
FOLDER_MAPPING = {
    "Inbox": ["Inbox", "Boîte de réception", "Bandeja de entrada", "Posteingang"],
    "Sent Items": ["Sent Items", "Éléments envoyés", "Elementos enviados", "Gesendete Elemente"],
    "Deleted Items": ["Deleted Items", "Éléments supprimés", "Elementos eliminados", "Gelöschte Elemente"]
}

class OutlookWorker(QThread):
    """Worker thread for Outlook operations to prevent GUI freezing"""
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal(tuple)
    preview_data = pyqtSignal(dict)
    account_detected = pyqtSignal(dict)  # Signal for account detection
    
    def __init__(self, age: int, unit: str, folders: List[str], action: str, 
                 preview_mode: bool = False, outlook_version: dict = None,
                 selected_accounts: List[str] = None):
        super().__init__()
        self.age = age
        self.unit = unit
        self.folders = folders
        self.action = action
        self.preview_mode = preview_mode
        self.outlook_version = outlook_version
        self.selected_accounts = selected_accounts
        self.is_running = True
        self.batch_size = 100 if outlook_version and outlook_version['is_modern'] else 20
        
    def get_folder_constant(self, folder_name: str) -> Optional[int]:
        """Get the Outlook folder constant for default folders"""
        folder_name = folder_name.lower()
        if folder_name == "inbox":
            return FOLDER_INBOX
        elif folder_name == "sent items":
            return FOLDER_SENT
        elif folder_name == "deleted items":
            return FOLDER_DELETED
        return None
    
    def find_folder_in_store(self, store, folder_name: str) -> Optional[object]:
        """Find a folder in a store using either default folder constant or localized name"""
        try:
            # First try using default folder constant
            folder_constant = self.get_folder_constant(folder_name)
            if folder_constant is not None:
                try:
                    return store.GetDefaultFolder(folder_constant)
                except Exception as e:
                    logger.debug(f"Could not get default folder using constant {folder_constant}: {str(e)}")
            
            # If that fails, try searching by localized names
            root_folder = store.GetRootFolder()
            localized_names = FOLDER_MAPPING.get(folder_name, [folder_name])
            
            for folder in root_folder.Folders:
                if folder.Name in localized_names:
                    return folder
            
            return None
        except Exception as e:
            logger.error(f"Error finding folder {folder_name} in store: {str(e)}")
            return None
    
    def get_store_email(self, store, namespace) -> Optional[str]:
        """Get the email address associated with a store"""
        try:
            # Try to get the account associated with this store
            for account in namespace.Accounts:
                try:
                    if hasattr(account, 'DeliveryStore'):
                        if account.DeliveryStore.StoreID == store.StoreID:
                            return account.SmtpAddress
                except:
                    continue
            
            # If no match found, try to get from the store's properties
            try:
                pr_smtp_address = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
                email = store.PropertyAccessor.GetProperty(pr_smtp_address)
                if email:
                    return email
            except:
                pass
            
            return None
        except Exception as e:
            logger.error(f"Error getting store email: {str(e)}")
            return None

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
            
            # Get all stores (root folders)
            stores = namespace.Stores
            logger.info(f"Found {stores.Count} stores in Outlook")
            
            # Get all accounts for reference
            outlook_accounts = get_outlook_accounts()
            
            for store in stores:
                if not self.is_running:
                    break
                
                try:
                    store_name = store.DisplayName
                    logger.info(f"Processing store: {store_name}")
                    
                    # Get the email address for this store
                    store_email = self.get_store_email(store, namespace)
                    if store_email:
                        logger.info(f"Store email: {store_email}")
                    
                    # Check if this store's account is selected
                    store_matches = False
                    if not self.selected_accounts:
                        store_matches = True
                    elif store_email:
                        # Try matching by email
                        store_matches = any(acc for acc in outlook_accounts 
                                         if acc.smtp_address == store_email 
                                         and acc.display_name in self.selected_accounts)
                    
                    if not store_matches:
                        self.status.emit(f"Skipping store {store_name} - account not selected")
                        continue
                    
                    for folder_name in self.folders:
                        if not self.is_running:
                            break
                        
                        try:
                            # Find the target folder in this store
                            folder = self.find_folder_in_store(store, folder_name)
                            
                            if not folder:
                                self.status.emit(f"Folder {folder_name} not found in {store_name}")
                                continue
                            
                            # Detect account for this folder
                            account = get_account_by_folder(namespace, folder)
                            if account:
                                # Emit account info
                                self.account_detected.emit(account.to_dict())
                            
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
                            logger.error(f"Error accessing folder {folder_name} in store {store_name}: {str(e)}")
                            self.status.emit(f"Error in folder {folder_name}: {str(e)}")
                            continue
                            
                except Exception as e:
                    logger.error(f"Error accessing store {store_name}: {str(e)}")
                    self.status.emit(f"Error accessing store {store_name}: {str(e)}")
                    continue
            
            if self.preview_mode:
                self.preview_data.emit(emails_by_recipient)
            self.finished.emit((total_processed, total_affected))
            
        except Exception as e:
            logger.error(f"Critical error in Outlook operations: {str(e)}")
            self.status.emit(f"Critical error: {str(e)}")
            self.finished.emit((0, 0))

    def stop(self):
        """Stop the worker thread"""
        self.is_running = False 