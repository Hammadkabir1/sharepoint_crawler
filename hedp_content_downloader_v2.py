import os
from pathlib import Path
from dotenv import load_dotenv
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from langchain_community.document_loaders import PyPDFLoader, TextLoader, Docx2txtLoader

def setup_sharepoint_connection():
    load_dotenv()
    USERNAME = os.getenv('sharepoint_email')
    PASSWORD = os.getenv('sharepoint_password')
    SITE_URL = os.getenv('sharepoint_url_site')
    HEDP_FOLDER_PATH = os.getenv('hedp_folder_path')
    
    if not all([USERNAME, PASSWORD, SITE_URL, HEDP_FOLDER_PATH]):
        raise ValueError("Missing required environment variables")
    
    if SITE_URL.endswith('/'):
        SITE_URL = SITE_URL[:-1]
    
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
    
    return ctx, base_path

def download_file_directly(ctx, base_path, filename, downloads_dir):
    try:
        file_url = f"{base_path}/{filename}"
        local_file_path = os.path.join(downloads_dir, filename)
        
        file = ctx.web.get_file_by_server_relative_url(file_url)
        
        with open(local_file_path, 'wb') as local_file:
            file.download(local_file)
            ctx.execute_query()
        
        if os.path.exists(local_file_path) and os.path.getsize(local_file_path) > 0:
            return local_file_path
        return None
            
    except Exception as e:
        return None

def get_loader_for_file(file_path):
    extension = Path(file_path).suffix.lower()
    
    if extension == '.pdf':
        return PyPDFLoader(file_path)
    elif extension == '.txt':
        return TextLoader(file_path, encoding='utf-8')
    elif extension in ['.docx', '.doc']:
        return Docx2txtLoader(file_path)
    else:
        raise ValueError(f"Unsupported file type: {extension}")

def process_with_langchain(local_file_path):
    try:
        loader = get_loader_for_file(local_file_path)
        documents = loader.load()
        
        content = ""
        for doc in documents:
            content += doc.page_content + "\n"
        
        return content.strip()
        
    except Exception as e:
        return f"Error processing file: {str(e)}"

def download_and_extract_content(filenames):
    if not isinstance(filenames, list):
        return {"error": "Input must be a list of file names"}
    
    try:
        ctx, base_path = setup_sharepoint_connection()
    except Exception as e:
        return {"error": f"SharePoint connection failed: {str(e)}"}
    
    downloads_dir = "downloads"
    os.makedirs(downloads_dir, exist_ok=True)
    
    results = {}
    
    for filename in filenames:
        local_file_path = download_file_directly(ctx, base_path, filename, downloads_dir)
        
        if not local_file_path:
            results[filename] = {
                "status": "error",
                "message": "Failed to download file from SharePoint"
            }
            continue
        
        content = process_with_langchain(local_file_path)
        
        if content.startswith("Error processing"):
            results[filename] = {
                "status": "error",
                "message": content
            }
        else:
            results[filename] = {
                "status": "success",
                "content": content,
                "content_length": len(content),
                "file_type": Path(filename).suffix.lower(),
                "local_path": local_file_path
            }
        

    
    return results

def main():
    test_files = [
        "SAP-HCM( FAQs).pdf",
        "Finance FAQs.txt",
        "Lab and Estate mangement Manual.txt"
    ]
    
    results = download_and_extract_content(test_files)
    
    if "error" in results:
        print(f"Error: {results['error']}")
    else:
        for filename, result in results.items():
            status = result.get('status', 'unknown')
            if status == 'success':
                print(f"SUCCESS: {filename} - {result.get('content_length', 0)} characters")
                print(f"Saved to: {result.get('local_path', 'N/A')}")
            else:
                print(f"ERROR: {filename} - {result.get('message', 'Unknown error')}")
    
    return results

if __name__ == "__main__":
    main()