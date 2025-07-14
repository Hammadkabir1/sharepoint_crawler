import os
import json
from datetime import datetime
from dotenv import load_dotenv
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

def format_file_size(size_bytes):
    """Convert bytes to human readable format"""
    if size_bytes is None or size_bytes == "" or size_bytes == "Unknown":
        return "Unknown"
    
    try:
        if isinstance(size_bytes, str):
            size_bytes = int(size_bytes)
        elif not isinstance(size_bytes, int):
            return "Unknown"
    except (ValueError, TypeError):
        return "Unknown"
    
    if size_bytes <= 0:
        return "0 B"
    
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.1f} MB"

def get_file_extension(filename):
    """Extract file extension from filename"""
    return os.path.splitext(filename)[1].lower()

def extract_folder_contents(ctx, folder_path, current_level=0, max_depth=5):
    """Recursively extract all folder contents"""
    items = []
    
    if current_level > max_depth:
        return items
    
    try:
        # Get folder contents
        folder = ctx.web.get_folder_by_server_relative_url(folder_path)
        files = folder.files
        subfolders = folder.folders
        ctx.load(files)
        ctx.load(subfolders)
        ctx.execute_query()
        
        # Process files in current folder
        for file in files:
            file_name = file.properties["Name"]
            file_size = file.properties.get("Length", 0)
            
            # Ensure file_size is an integer
            try:
                if isinstance(file_size, str):
                    file_size = int(file_size) if file_size.isdigit() else 0
                elif file_size is None:
                    file_size = 0
                else:
                    file_size = int(file_size)
            except (ValueError, TypeError):
                file_size = 0
            
            size_str = format_file_size(file_size)
            extension = get_file_extension(file_name)
            
            items.append({
                "type": "file",
                "name": file_name,
                "path": folder_path,
                "size": size_str,
                "size_bytes": file_size,
                "extension": extension,
                "level": current_level
            })
        
        # Process subfolders
        for subfolder in subfolders:
            folder_name = subfolder.properties["Name"]
            if not folder_name.startswith('.'):
                subfolder_path = f"{folder_path}/{folder_name}"
                
                # Add folder info
                items.append({
                    "type": "folder",
                    "name": folder_name,
                    "path": subfolder_path,
                    "level": current_level
                })
                
                # Recursively get subfolder contents
                subfolder_items = extract_folder_contents(ctx, subfolder_path, current_level + 1, max_depth)
                items.extend(subfolder_items)
    
    except Exception as e:
        pass  # Silent error handling for production
    
    return items

def display_hierarchical_structure(items):
    """Display items in a hierarchical tree structure (deprecated - for legacy compatibility)"""
    # Function kept for backward compatibility but no longer outputs to console
    pass

def generate_summary(items):
    """Generate summary statistics (deprecated - for legacy compatibility)"""
    # Function kept for backward compatibility but no longer outputs to console
    pass

def get_files_by_folder(items):
    """Group files by their parent folder names"""
    folder_files = {}
    
    for item in items:
        if item['type'] == 'file':
            # Extract folder name from path or use 'root' for top-level files
            path_parts = item['path'].split('/')
            if len(path_parts) > 1:
                folder_name = path_parts[-2]  # Parent folder name
            else:
                folder_name = 'root'
            
            if folder_name not in folder_files:
                folder_files[folder_name] = []
            folder_files[folder_name].append(item['name'])
    
    return folder_files

def format_for_frontend_api(items, organization_name="HEDP"):
    """Format data for simple frontend API consumption - grouped by folders"""
    return get_files_by_folder(items)

def format_detailed_json(items, organization_name="HEDP"):
    """Format data as detailed JSON with metadata and statistics (deprecated)"""
    # Simplified to match user requirements - just return simple format
    return format_for_frontend_api(items, organization_name)

def extract_sharepoint_data_as_json(format_type="simple", organization_name="HEDP"):
    """Extract SharePoint data and return as simple JSON
    
    Args:
        format_type (str): Ignored - always returns simple format
        organization_name (str): Name of the organization (used as key in JSON)
    
    Returns:
        dict: Simple JSON with organization name as key and filenames as array
    """
    # Load environment variables
    load_dotenv()
    USERNAME = os.getenv('sharepoint_email')
    PASSWORD = os.getenv('sharepoint_password')
    SITE_URL = os.getenv('sharepoint_url_site')
    HEDP_FOLDER_PATH = os.getenv('hedp_folder_path')
    
    if not all([USERNAME, PASSWORD, SITE_URL, HEDP_FOLDER_PATH]):
        return {organization_name: []}
    
    if SITE_URL.endswith('/'):
        SITE_URL = SITE_URL[:-1]
    
    try:
        # Handle OneDrive personal folders
        if "personal/" in HEDP_FOLDER_PATH:
            personal_site_url = f"{SITE_URL}/personal/muhammad_fawwaz_tmcltd_com"
            ctx = ClientContext(personal_site_url).with_credentials(UserCredential(USERNAME, PASSWORD))
            
            web = ctx.web
            ctx.load(web, ["ServerRelativeUrl"])
            ctx.execute_query()
            
            base_path = f"{web.properties['ServerRelativeUrl']}/Documents/HEDP"
        else:
            ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, PASSWORD))
            web = ctx.web
            ctx.load(web, ["ServerRelativeUrl"])
            ctx.execute_query()
            
            base_path = f"{web.properties['ServerRelativeUrl']}/{HEDP_FOLDER_PATH}"
        
        # Extract all data recursively
        all_items = extract_folder_contents(ctx, base_path)
        
        # Always return simple format
        return format_for_frontend_api(all_items, organization_name)
        
    except Exception as e:
        return {organization_name: []}

def extract_all_sharepoint_data():
    """Main function to extract all SharePoint data (legacy display format - deprecated)"""
    # Deprecated function - use extract_sharepoint_data_as_json() instead
    return extract_sharepoint_data_as_json("simple", "HEDP")

def main():
    """Main function - returns simple JSON data for production use"""
    # Production version - returns simple data format
    return extract_sharepoint_data_as_json("simple", "HEDP")

if __name__ == "__main__":
    # For direct script execution, output JSON to stdout
    result = main()
    print(json.dumps(result, indent=2))