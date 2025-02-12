"""
Outlook Detection Module
Provides functionality to detect and verify Outlook installations and versions.
"""

import win32com.client
import winreg
import psutil
import logging
from typing import List, Dict, Optional
from dataclasses import dataclass
from enum import Enum, auto

logger = logging.getLogger(__name__)

class OutlookType(Enum):
    """Enumeration of supported Outlook types"""
    CLASSIC = auto()
    NEW = auto()

@dataclass
class OutlookVersion:
    """Data class representing an Outlook version"""
    version: str
    name: str
    is_modern: bool
    major_version: str
    path: str
    bitness: str
    is_active: bool
    display_name: str
    description: str

    def to_dict(self) -> Dict:
        """Convert to dictionary format for UI compatibility"""
        return {
            "version": self.version,
            "name": self.name,
            "is_modern": self.is_modern,
            "major_version": self.major_version,
            "path": self.path,
            "bitness": self.bitness,
            "is_active": self.is_active,
            "display_name": self.display_name,
            "description": self.description
        }

def detect_active_outlook() -> Optional[OutlookType]:
    """
    Detect which version of Outlook is currently running by examining processes.
    Returns: OutlookType.NEW for New Outlook, OutlookType.CLASSIC for Classic Outlook,
            or None if not found
    """
    try:
        for proc in psutil.process_iter(['name', 'exe']):
            try:
                proc_info = proc.info
                if not proc_info or 'name' not in proc_info:
                    continue
                    
                proc_name = proc_info['name'].lower()
                if proc_name not in ['outlook.exe', 'microsoft.office.outlook.hub.exe']:
                    continue
                    
                exe_path = proc_info.get('exe', '')
                if not exe_path:
                    continue
                    
                logger.debug(f"Found Outlook process: {exe_path}")
                
                # Check if it's the new Outlook
                if any(indicator in exe_path.lower() for indicator in 
                      ['windowsapps\\microsoft.outlookforwindows', 'msedge_shell.exe']):
                    logger.debug("Detected New Outlook as active")
                    return OutlookType.NEW
                else:
                    logger.debug("Detected Classic Outlook as active")
                    return OutlookType.CLASSIC
                    
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as e:
                logger.debug(f"Skipping process due to: {str(e)}")
                continue
                
        logger.debug("No running Outlook process found")
        return None
        
    except Exception as e:
        logger.error(f"Error detecting active Outlook: {str(e)}")
        return None

def detect_new_outlook_installation() -> bool:
    """
    Check for New Outlook installation using registry and process detection.
    Returns: True if New Outlook is installed, False otherwise
    """
    try:
        # First check registry
        user_registry_paths = [
            (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\16.0\Outlook\Options"),
            (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\16.0\Outlook\Preferences"),
            (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\outlook")
        ]
        
        registry_keys = ["NewOutlook", "UseNewOutlook", "Microsoft.Office.Outlook.Hub.HubApp"]
        
        for root_key, path in user_registry_paths:
            try:
                with winreg.OpenKey(root_key, path, 0, winreg.KEY_READ) as key:
                    for value_name in registry_keys:
                        try:
                            value, _ = winreg.QueryValueEx(key, value_name)
                            if value:
                                logger.debug(f"Found New Outlook indicator in registry: {value_name}={value}")
                                return True
                        except OSError:  # Using OSError instead of WindowsError for better compatibility
                            continue
            except OSError as e:
                logger.debug(f"Could not access registry path {path}: {str(e)}")
                continue
        
        # If registry check fails, check if New Outlook is currently running
        active_version = detect_active_outlook()
        if active_version == OutlookType.NEW:
            logger.debug("New Outlook detected through process inspection")
            return True
            
        return False
        
    except Exception as e:
        logger.error(f"Error checking for New Outlook installation: {str(e)}")
        return False

def create_classic_outlook_version(version: str, is_active: bool) -> OutlookVersion:
    """Create a Classic Outlook version object"""
    return OutlookVersion(
        version=version,
        name="Outlook (Classic)",
        is_modern=True,
        major_version=".".join(version.split(".")[:2]),
        path="Classic Installation",
        bitness="Desktop",
        is_active=is_active,
        display_name="Outlook (Classic)",
        description="Traditional desktop version of Outlook with full functionality"
    )

def create_new_outlook_version(is_active: bool) -> OutlookVersion:
    """Create a New Outlook version object"""
    return OutlookVersion(
        version="16.0",  # New Outlook is always on latest version
        name="Outlook (New)",
        is_modern=True,
        major_version="16.0",
        path="New Installation",
        bitness="Modern",
        is_active=is_active,
        display_name="Outlook (New)",
        description="Modern web-based version of Outlook with updated interface"
    )

def check_outlook_version() -> List[Dict]:
    """
    Check Outlook version and compatibility.
    Returns a list of detected Outlook versions with their details.
    The list will always be returned, even if empty or containing only one version.
    Each version is represented as a dictionary for UI compatibility.
    """
    try:
        outlook_versions = []
        active_version = detect_active_outlook()
        logger.debug(f"Active Outlook version detected: {active_version}")
        
        # Try to detect Classic Outlook
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            classic_version = outlook.Version
            logger.debug(f"Found Classic Outlook version: {classic_version}")
            
            outlook_versions.append(
                create_classic_outlook_version(
                    classic_version, 
                    is_active=(active_version == OutlookType.CLASSIC)
                ).to_dict()
            )
        except Exception as e:
            logger.debug(f"Could not detect Classic Outlook: {str(e)}")
        
        # Check for New Outlook
        if detect_new_outlook_installation() or active_version == OutlookType.NEW:
            outlook_versions.append(
                create_new_outlook_version(
                    is_active=(active_version == OutlookType.NEW)
                ).to_dict()
            )
            logger.debug("Added New Outlook to detected versions")
        
        if not outlook_versions:
            logger.error("No Outlook versions found")
            return []
            
        # If no active version was detected but we found installations,
        # mark the first one as active
        if active_version is None and outlook_versions:
            outlook_versions[0]["is_active"] = True
            logger.debug(f"No active version detected, marking {outlook_versions[0]['name']} as active")
        
        return outlook_versions
            
    except Exception as e:
        logger.error(f"Error checking Outlook version: {str(e)}")
        return [] 