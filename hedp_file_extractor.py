import os
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
        print(f"⚠️  Error accessing folder {folder_path}: {str(e)}")
    
    return items

def display_hierarchical_structure(items):
    """Display items in a hierarchical tree structure"""
    if not items:
        print("📂 No items found.")
        return
    
    # Group items by level for better display
    levels = {}
    for item in items:
        level = item['level']
        if level not in levels:
            levels[level] = []
        levels[level].append(item)
    
    # Display items level by level
    for level in sorted(levels.keys()):
        level_items = levels[level]
        
        # Separate folders and files
        folders = [item for item in level_items if item['type'] == 'folder']
        files = [item for item in level_items if item['type'] == 'file']
        
        if folders:
            print(f"\n📁 LEVEL {level} - FOLDERS ({len(folders)} folders):")
            print("-" * 60)
            for folder in folders:
                indent = "  " * level
                print(f"{indent}📁 {folder['name']}/")
        
        if files:
            print(f"\n📄 LEVEL {level} - FILES ({len(files)} files):")
            print("-" * 60)
            for file in files:
                indent = "  " * level
                print(f"{indent}📄 {file['name']} ({file['size']})")

def generate_summary(items):
    """Generate summary statistics"""
    if not items:
        return
    
    folders = [item for item in items if item['type'] == 'folder']
    files = [item for item in items if item['type'] == 'file']
    
    # Calculate total size
    total_size_bytes = sum(item.get('size_bytes', 0) for item in files)
    total_size_str = format_file_size(total_size_bytes)
    
    # File type distribution
    file_types = {}
    for file in files:
        ext = file.get('extension', 'no extension')
        file_types[ext] = file_types.get(ext, 0) + 1
    
    print("\n" + "=" * 70)
    print("📊 SUMMARY STATISTICS")
    print("=" * 70)
    print(f"📁 Total Folders: {len(folders)}")
    print(f"📄 Total Files: {len(files)}")
    print(f"💾 Total Size: {total_size_str}")
    print(f"📊 Total Items: {len(items)}")
    
    if file_types:
        print("\n📋 File Types Distribution:")
        for ext, count in sorted(file_types.items()):
            ext_display = ext if ext else 'no extension'
            print(f"   {ext_display}: {count} files")

def extract_all_sharepoint_data():
    """Main function to extract all SharePoint data"""
    # Load environment variables
    load_dotenv()
    USERNAME = os.getenv('sharepoint_email')
    PASSWORD = os.getenv('sharepoint_password')
    SITE_URL = os.getenv('sharepoint_url_site')
    HEDP_FOLDER_PATH = os.getenv('hedp_folder_path')
    
    if not all([USERNAME, PASSWORD, SITE_URL, HEDP_FOLDER_PATH]):
        print("❌ Missing environment variables. Check .env file.")
        return
    
    if SITE_URL.endswith('/'):
        SITE_URL = SITE_URL[:-1]
    
    try:
        print("🔗 Connecting to SharePoint...")
        
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
        
        print(f"📂 Extracting data from: {base_path}")
        print("⏳ This may take a moment for large folder structures...")
        
        # Extract all data recursively
        all_items = extract_folder_contents(ctx, base_path)
        
        # Display results
        print("\n" + "=" * 70)
        print("🗂️  COMPLETE SHAREPOINT FOLDER STRUCTURE")
        print("=" * 70)
        
        if all_items:
            display_hierarchical_structure(all_items)
            generate_summary(all_items)
        else:
            print("📂 No items found in the specified folder.")
        
        print("\n✅ Data extraction completed!")
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        print("Check your credentials and folder path.")

def main():
    """Main function"""
    print("=" * 70)
    print("🗂️  HEDP SharePoint Complete Data Extractor")
    print("=" * 70)
    extract_all_sharepoint_data()

if __name__ == "__main__":
    main()