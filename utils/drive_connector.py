import os
import glob
from typing import List, Optional

class DriveConnector:
    """Handles connection to shared drive and file operations"""
    
    def __init__(self):
        pass
    
    def get_project_folders(self, construction_folder: str) -> List[str]:
        """Get list of project folders in the construction directory"""
        try:
            if not os.path.exists(construction_folder):
                return []
            
            folders = []
            for item in os.listdir(construction_folder):
                item_path = os.path.join(construction_folder, item)
                if os.path.isdir(item_path):
                    folders.append(item)
            
            return sorted(folders)
        
        except Exception as e:
            print(f"Error getting project folders: {str(e)}")
            return []
    
    def get_excel_files(self, project_path: str) -> List[str]:
        """Get list of Excel files in a project folder"""
        try:
            if not os.path.exists(project_path):
                return []
            
            excel_files = []
            
            # Look for Excel files with common extensions
            patterns = ['*.xlsx', '*.xls', '*.xlsm']
            
            for pattern in patterns:
                files = glob.glob(os.path.join(project_path, pattern))
                for file_path in files:
                    filename = os.path.basename(file_path)
                    if not filename.startswith('~'):  # Skip temporary files
                        excel_files.append(filename)
            
            return sorted(list(set(excel_files)))  # Remove duplicates and sort
        
        except Exception as e:
            print(f"Error getting Excel files: {str(e)}")
            return []
    
    def get_file_path(self, construction_folder: str, project: str, filename: str) -> str:
        """Get full path to a specific file"""
        return os.path.join(construction_folder, project, filename)
    
    def file_exists(self, file_path: str) -> bool:
        """Check if file exists"""
        return os.path.exists(file_path)
    
    def get_file_info(self, file_path: str) -> Optional[dict]:
        """Get file information"""
        try:
            if not os.path.exists(file_path):
                return None
            
            stat = os.stat(file_path)
            return {
                'size': stat.st_size,
                'modified': stat.st_mtime,
                'created': stat.st_ctime
            }
        
        except Exception as e:
            print(f"Error getting file info: {str(e)}")
            return None
