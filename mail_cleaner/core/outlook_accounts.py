"""
Outlook Account Detection Module
Provides functionality to detect and analyze Outlook account configurations.
"""

import win32com.client
import logging
from enum import IntEnum
from typing import List, Dict, Optional
from dataclasses import dataclass

logger = logging.getLogger(__name__)

class AccountType(IntEnum):
    """Outlook account types"""
    UNKNOWN = 0
    EXCHANGE = 1
    IMAP = 2
    POP3 = 3
    HTTP = 4  # Including Gmail, Outlook.com
    EAS = 5   # Exchange ActiveSync
    OFFICE365 = 6

@dataclass
class OutlookAccount:
    """Data class representing an Outlook account configuration"""
    display_name: str
    user_name: str
    smtp_address: str
    account_type: AccountType
    is_primary: bool
    store_path: Optional[str] = None
    
    def is_gmail(self) -> bool:
        """Check if this is a Gmail account"""
        return any(
            indicator in (self.smtp_address or "").lower()
            for indicator in ["@gmail.com", "googlemail.com"]
        )
    
    def is_office365(self) -> bool:
        """Check if this is an Office 365 account"""
        return (
            self.account_type == AccountType.OFFICE365 or
            any(
                indicator in (self.smtp_address or "").lower()
                for indicator in ["@outlook.com", "@hotmail.com", "@live.com", "onmicrosoft.com"]
            )
        )
    
    def to_dict(self) -> Dict:
        """Convert to dictionary format"""
        return {
            "display_name": self.display_name,
            "user_name": self.user_name,
            "smtp_address": self.smtp_address,
            "account_type": self.account_type.name,
            "is_primary": self.is_primary,
            "store_path": self.store_path,
            "is_gmail": self.is_gmail(),
            "is_office365": self.is_office365()
        }

def detect_account_type(account) -> AccountType:
    """
    Detect the type of an Outlook account based on its properties
    """
    try:
        # Try to get the AccountType property directly
        account_type = getattr(account, 'AccountType', None)
        if account_type is not None:
            logger.debug(f"Found AccountType property: {account_type}")
            # Map Outlook's account type values to our enum
            # Note: These mappings may need adjustment based on Outlook versions
            type_map = {
                0: AccountType.UNKNOWN,
                1: AccountType.EXCHANGE,
                2: AccountType.IMAP,
                3: AccountType.POP3,
                4: AccountType.HTTP,
                5: AccountType.EAS,
                6: AccountType.OFFICE365
            }
            return type_map.get(account_type, AccountType.UNKNOWN)
        
        # If AccountType not available, try to infer from other properties
        smtp_address = getattr(account, 'SmtpAddress', '').lower()
        if '@gmail.com' in smtp_address or '@googlemail.com' in smtp_address:
            return AccountType.HTTP
        elif any(domain in smtp_address for domain in ['@outlook.com', '@hotmail.com', '@live.com', 'onmicrosoft.com']):
            return AccountType.OFFICE365
        
        # Check for Exchange-specific properties
        if hasattr(account, 'ExchangeConnectionMode'):
            return AccountType.EXCHANGE
            
        # Default to UNKNOWN if we can't determine the type
        return AccountType.UNKNOWN
        
    except Exception as e:
        logger.error(f"Error detecting account type: {str(e)}")
        return AccountType.UNKNOWN

def get_outlook_accounts() -> List[OutlookAccount]:
    """
    Retrieve and analyze all configured Outlook accounts.
    
    Returns:
        List of OutlookAccount objects representing each configured account.
    """
    accounts = []
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Get all accounts
        outlook_accounts = namespace.Accounts
        logger.info(f"Found {outlook_accounts.Count} account(s) configured in Outlook")
        
        # Get default account/store for primary detection
        default_store = namespace.DefaultStore
        default_store_path = getattr(default_store, 'FilePath', None)
        
        for account in outlook_accounts:
            try:
                # Get basic account properties
                display_name = getattr(account, 'DisplayName', 'Unknown Account')
                user_name = getattr(account, 'UserName', '')
                smtp_address = getattr(account, 'SmtpAddress', '')
                
                # Get store path if available
                store_path = None
                try:
                    if hasattr(account, 'DeliveryStore'):
                        store_path = account.DeliveryStore.FilePath
                except:
                    pass
                
                # Determine if this is the primary account
                is_primary = (store_path == default_store_path if store_path 
                            else account.DisplayName == namespace.CurrentUser.Name)
                
                # Create account object
                account_obj = OutlookAccount(
                    display_name=display_name,
                    user_name=user_name,
                    smtp_address=smtp_address,
                    account_type=detect_account_type(account),
                    is_primary=is_primary,
                    store_path=store_path
                )
                
                logger.info(
                    f"Account detected: {display_name} "
                    f"({account_obj.account_type.name}, "
                    f"{'Primary' if is_primary else 'Secondary'})"
                )
                
                if account_obj.is_gmail():
                    logger.info(f"Gmail account detected: {display_name}")
                if account_obj.is_office365():
                    logger.info(f"Office 365 account detected: {display_name}")
                
                accounts.append(account_obj)
                
            except Exception as e:
                logger.error(f"Error processing account {getattr(account, 'DisplayName', 'Unknown')}: {str(e)}")
                continue
        
    except Exception as e:
        logger.error(f"Error accessing Outlook accounts: {str(e)}")
    
    return accounts

def get_account_by_folder(namespace, folder) -> Optional[OutlookAccount]:
    """
    Attempt to determine which account a folder belongs to.
    
    Args:
        namespace: Outlook MAPI namespace
        folder: Outlook folder object
        
    Returns:
        OutlookAccount object if the account can be determined, None otherwise
    """
    try:
        # Get the store that contains this folder
        store = folder.Store
        store_path = getattr(store, 'FilePath', None)
        
        # Get all accounts
        accounts = get_outlook_accounts()
        
        # First try matching by store path
        if store_path:
            for account in accounts:
                if account.store_path == store_path:
                    return account
        
        # If no match by path, try matching by display name
        store_display_name = getattr(store, 'DisplayName', '')
        for account in accounts:
            if account.display_name in store_display_name:
                return account
        
        return None
        
    except Exception as e:
        logger.error(f"Error determining account for folder: {str(e)}")
        return None 