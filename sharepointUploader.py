import os
import requests
from urllib.parse import quote
import msal
import logging
from datetime import datetime

# Configuration constants
CLIENT_ID = 'XXXXX-42za-XXXXX1f28f'
CLIENT_SECRET = '_jXXXXX~vXXXXXCotWJROq.pWYXXXXXBcSi'
TENANT_ID = 'bXXXXX91c-4e21-4b97-XXXXX'
AUTHORITY = f'https://login.microsoftonline.com/{TENANT_ID}'
SHAREPOINT_SITE_URL = 'https://graph.microsoft.com/v1.0/sites/XXXXXX:/sites/XXXXXX'
LOCAL_DIRECTORY = 'E:/workspace/EXTRACT/opt'  # Local directory to synchronize
SHAREPOINT_BASE_FOLDER = '/data_tree'

# Setup logging
def setup_logging():
    """Setup logging configuration."""
    # Create logs directory if it doesn't exist
    logs_dir = 'logs/synchro'
    os.makedirs(logs_dir, exist_ok=True)
    
    # Create log filename with timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_filename = os.path.join(logs_dir, f'sharepoint_sync_{timestamp}.log')
    
    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filename, encoding='utf-8'),
            logging.StreamHandler()  # Also log to console
        ]
    )
    
    return logging.getLogger(__name__)

# Get access token
def get_access_token():
    """Get access token from Azure AD."""
    logger.info("Requesting access token from Azure AD")
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if 'access_token' in result:
        logger.info("Access token obtained successfully")
        return result['access_token']
    else:
        error_msg = f"Error: Token not obtained, {result.get('error_description')}"
        logger.error(error_msg)
        raise Exception(error_msg)

# Get site ID
def get_site_id(access_token):
    """Retrieve SharePoint site ID."""
    logger.info("Retrieving SharePoint site ID")
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(SHAREPOINT_SITE_URL, headers=headers)
    if response.status_code == 200:
        site_id = response.json().get('id')
        logger.info(f"Site ID retrieved successfully: {site_id}")
        return site_id
    else:
        error_msg = f"Error: Unable to get site ID, {response.text}"
        logger.error(error_msg)
        raise Exception(error_msg)

# Recursively delete SharePoint folder contents
def delete_folder_contents(access_token, site_id, folder_path):
    """
    Recursively delete all contents of a SharePoint folder.
    """
    logger.info(f"Starting deletion of folder contents: {folder_path}")
    headers = {'Authorization': f'Bearer {access_token}'}
    
    # Normalize folder path
    folder_path = folder_path.strip('/')
    if not folder_path:
        logger.warning("Empty folder path provided, skipping deletion")
        return
    
    folder_path = '/' + folder_path
    folder_path_encoded = quote(folder_path)
    
    # Check if folder exists
    check_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{folder_path_encoded}"
    response = requests.get(check_url, headers=headers)
    
    if response.status_code == 404:
        logger.info(f"Folder {folder_path} does not exist yet")
        return
    elif response.status_code != 200:
        error_msg = f"Error checking folder {folder_path}: {response.status_code} - {response.text}"
        logger.error(error_msg)
        raise Exception(error_msg)
    
    # List folder contents
    list_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{folder_path_encoded}:/children"
    response = requests.get(list_url, headers=headers)
    
    if response.status_code != 200:
        error_msg = f"Error retrieving contents of {folder_path}: {response.status_code} - {response.text}"
        logger.error(error_msg)
        raise Exception(error_msg)
    
    items = response.json().get('value', [])
    
    if not items:
        logger.info(f"Folder {folder_path} is already empty")
        return
    
    logger.info(f"Deleting {len(items)} item(s) in {folder_path}")
    
    # Delete each item
    for item in items:
        item_id = item['id']
        item_name = item['name']
        item_type = "folder" if 'folder' in item else "file"
        
        delete_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}"
        delete_response = requests.delete(delete_url, headers=headers)
        
        if delete_response.status_code == 204:
            logger.info(f"{item_type.capitalize()} deleted: {item_name}")
        else:
            logger.error(f"Error deleting {item_name}: {delete_response.status_code} - {delete_response.text}")

# Check and create folders on SharePoint
def ensure_folder_exists(access_token, site_id, folder_path):
    """
    Check if a folder exists on SharePoint, create it if it doesn't.
    """
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    
    # Normalize folder path
    folder_path = folder_path.strip('/')
    if not folder_path:  # If path is empty (root), exit
        return
    
    folder_path = '/' + folder_path
    folder_path_encoded = quote(folder_path)
    
    # Check if folder exists
    check_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{folder_path_encoded}"
    response = requests.get(check_url, headers=headers)

    if response.status_code == 200:
        logger.debug(f"Folder {folder_path} already exists")
        return
    elif response.status_code == 404:  # Folder not found, create it
        logger.info(f"Creating folder: {folder_path}")
        
        # Separate parent path and folder name
        path_parts = folder_path.strip('/').split('/')
        folder_name = path_parts[-1]
        
        if len(path_parts) > 1:
            parent_path = '/' + '/'.join(path_parts[:-1])
            # Create parent folders first
            ensure_folder_exists(access_token, site_id, parent_path)
        else:
            parent_path = ''
        
        # Create target folder
        if parent_path:
            parent_path_encoded = quote(parent_path)
            create_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{parent_path_encoded}:/children"
        else:
            create_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children"
        
        payload = {
            "name": folder_name, 
            "folder": {}, 
            "@microsoft.graph.conflictBehavior": "replace"
        }
        
        create_response = requests.post(create_url, headers=headers, json=payload)
        if create_response.status_code == 201:
            logger.info(f"Folder {folder_path} created successfully")
        elif create_response.status_code == 409:
            logger.info(f"Folder {folder_path} already exists (conflict resolved)")
        else:
            error_msg = f"Error creating {folder_path}: {create_response.status_code} - {create_response.text}"
            logger.error(error_msg)
            raise Exception(error_msg)
    else:
        error_msg = f"Error checking {folder_path}: {response.status_code} - {response.text}"
        logger.error(error_msg)
        raise Exception(error_msg)

# Upload files to SharePoint
def upload_file(access_token, site_id, sharepoint_folder_path, local_file_path):
    """
    Upload a file to SharePoint in the specified folder.
    """
    headers = {'Authorization': f'Bearer {access_token}'}
    
    # Normalize folder path
    sharepoint_folder_path = sharepoint_folder_path.strip('/')
    if sharepoint_folder_path:
        sharepoint_folder_path = '/' + sharepoint_folder_path
    
    folder_path_encoded = quote(sharepoint_folder_path) if sharepoint_folder_path else ''
    file_name = os.path.basename(local_file_path)
    
    if sharepoint_folder_path:
        upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{folder_path_encoded}/{file_name}:/content"
    else:
        upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_name}:/content"

    try:
        with open(local_file_path, 'rb') as file_data:
            response = requests.put(upload_url, headers=headers, data=file_data)
            if response.status_code in [200, 201]:
                logger.info(f"File uploaded: {local_file_path} → {sharepoint_folder_path}/{file_name}")
            else:
                logger.error(f"Error uploading {local_file_path}: {response.status_code} - {response.text}")
    except Exception as e:
        logger.error(f"Error reading file {local_file_path}: {e}")

# Local to SharePoint synchronization
def sync_local_to_sharepoint(access_token, site_id, local_directory, sharepoint_folder):
    """
    Synchronize all files and subfolders from a local directory to SharePoint.
    """
    # First check that local directory exists
    if not os.path.exists(local_directory):
        error_msg = f"Local directory {local_directory} does not exist"
        logger.error(error_msg)
        raise Exception(error_msg)
    
    logger.info(f"Starting synchronization from {local_directory} to {sharepoint_folder}")
    
    for root, dirs, files in os.walk(local_directory):
        # Build relative path for current folder
        relative_path = os.path.relpath(root, local_directory).replace("\\", "/")
        
        if relative_path == '.':
            current_sharepoint_folder = sharepoint_folder
        else:
            current_sharepoint_folder = f"{sharepoint_folder}/{relative_path}"

        # Check or create current folder in SharePoint
        logger.info(f"Checking folder: {current_sharepoint_folder}")
        ensure_folder_exists(access_token, site_id, current_sharepoint_folder)

        # Upload files in current folder
        for file_name in files:
            local_file_path = os.path.join(root, file_name)
            logger.info(f"Uploading: {local_file_path}")
            upload_file(access_token, site_id, current_sharepoint_folder, local_file_path)


def list_drive_ids(access_token, site_id):
    """
    List all SharePoint document libraries (drives) for the site
    and display their Drive IDs.
    """
    logger.info("Retrieving SharePoint drives")

    headers = {
        'Authorization': f'Bearer {access_token}'
    }

    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        drives = response.json().get('value', [])

        if not drives:
            logger.warning("No drives found for this SharePoint site")
            return []

        logger.info(f"Found {len(drives)} drive(s):")

        print("\n" + "=" * 80)
        print("SHAREPOINT DRIVES")
        print("=" * 80)

        for drive in drives:
            drive_id = drive.get('id')
            drive_name = drive.get('name')
            drive_type = drive.get('driveType')
            web_url = drive.get('webUrl')

            print(f"\nName     : {drive_name}")
            print(f"Drive ID : {drive_id}")
            print(f"Type     : {drive_type}")
            print(f"Web URL  : {web_url}")

            logger.info(
                f"Drive found - Name: {drive_name} - ID: {drive_id}"
            )

        print("\n" + "=" * 80 + "\n")

        return drives

    else:
        error_msg = (
            f"Error retrieving drives: "
            f"{response.status_code} - {response.text}"
        )
        logger.error(error_msg)
        raise Exception(error_msg)

# Main program
if __name__ == "__main__":
    logger = setup_logging()

    try:
        logger.info("Starting SharePoint synchronization process")

        logger.info("Getting access token...")
        access_token = get_access_token()

        logger.info("Retrieving site ID...")
        site_id = get_site_id(access_token)

        # Display all available SharePoint Drive IDs
        logger.info("Retrieving available SharePoint drives...")
        list_drive_ids(access_token, site_id)

        logger.info("Synchronizing local files with SharePoint...")
        sync_local_to_sharepoint(
            access_token,
            site_id,
            LOCAL_DIRECTORY,
            SHAREPOINT_BASE_FOLDER
        )

        logger.info("Synchronization completed successfully!")

    except Exception as e:
        logger.error(f"An error occurred: {e}")
        import traceback
        logger.error("Error details:")
        logger.error(traceback.format_exc())
        
        
