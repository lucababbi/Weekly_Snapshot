import os
from datetime import datetime
import shutil

def SharePointUpload(Excel_Path, Sharepoint_Folder):

    if not os.path.exists(Excel_Path):
        raise FileNotFoundError(f"Excel file not found: {Excel_Path}")
    
    # VBA filename logic
    timestamp = datetime.now().strftime("%Y%m%d")
    file_name = f"{timestamp}_Weekly_Benchmarks_Snapshot.xlsx"
    target_path = os.path.join(Sharepoint_Folder, file_name)
    
    # Copy (triggers sync, like VBA "should")
    shutil.copy2(Excel_Path, target_path)
    
    print(f"Copied to synced folder: {target_path}")
