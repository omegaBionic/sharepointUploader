# Local to SharePoint Synchronization with Microsoft Graph

This Python script recursively synchronizes the contents of a local directory to a SharePoint document library using Microsoft Graph and Microsoft Entra ID application authentication.

It can also retrieve and display the `Drive ID` values of the document libraries available on the SharePoint site.

## Features

- Authentication to Microsoft Entra ID using `CLIENT_ID`, `CLIENT_SECRET`, and `TENANT_ID`.
- Automatic retrieval of the SharePoint `Site ID`.
- Lists SharePoint document libraries with:
  - library name;
  - `Drive ID`;
  - drive type;
  - web URL.
- Automatically creates missing SharePoint folders.
- Recursively scans the local directory.
- Uploads files while preserving the local folder structure.
- Logs activity to both a timestamped log file and the console.
- Includes a function for deleting the contents of a SharePoint folder.

## Requirements

- Python 3.9 or later recommended.
- A registered Microsoft Entra ID application.
- A valid client secret.
- Appropriate Microsoft Graph application permissions.
- Network access to Microsoft Entra ID and Microsoft Graph.

## Installation

Create a virtual environment:

```bash
python -m venv .venv
```

On Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

Install the dependencies:

```bash
pip install requests msal
```

Alternatively, create a `requirements.txt` file:

```text
msal
requests
```

Then run:

```bash
pip install -r requirements.txt
```

## Microsoft Entra ID Configuration

The application must be registered in Microsoft Entra ID and have Microsoft Graph application permissions that allow access to the required SharePoint sites and files.

Use the minimum permissions required by your environment. If access should be restricted to specific SharePoint sites, prefer a restricted site access model rather than tenant-wide access.

## Script Configuration

Update the constants at the beginning of the script:

```python
CLIENT_ID = 'YOUR_CLIENT_ID'
CLIENT_SECRET = 'YOUR_CLIENT_SECRET'
TENANT_ID = 'YOUR_TENANT_ID'

AUTHORITY = f'https://login.microsoftonline.com/{TENANT_ID}'

SHAREPOINT_SITE_URL = (
    'https://graph.microsoft.com/v1.0/sites/'
    'your-sharepoint-hostname:/sites/your-site-name'
)

LOCAL_DIRECTORY = 'E:/workspace/EXTRACT/opt'
SHAREPOINT_BASE_FOLDER = '/data_tree'
```

### Parameters

- `CLIENT_ID`: Microsoft Entra ID application/client ID.
- `CLIENT_SECRET`: application client secret.
- `TENANT_ID`: Microsoft Entra ID tenant ID.
- `SHAREPOINT_SITE_URL`: Microsoft Graph endpoint used to retrieve the SharePoint site.
- `LOCAL_DIRECTORY`: local directory to synchronize.
- `SHAREPOINT_BASE_FOLDER`: destination folder in the default SharePoint document library.

## Credentials Security

Do not store a real `CLIENT_SECRET` directly in source code or commit it to Git.

Environment variables are preferable:

```python
CLIENT_ID = os.environ['SP_CLIENT_ID']
CLIENT_SECRET = os.environ['SP_CLIENT_SECRET']
TENANT_ID = os.environ['SP_TENANT_ID']
```

Example with PowerShell:

```powershell
$env:SP_CLIENT_ID = "your-client-id"
$env:SP_CLIENT_SECRET = "your-client-secret"
$env:SP_TENANT_ID = "your-tenant-id"
python .\sharepoint_sync.py
```

A `.gitignore` should also exclude sensitive files and logs:

```gitignore
.venv/
__pycache__/
*.pyc
.env
logs/
```

## Running the Script

Save the Python script, for example as:

```text
sharepoint_sync.py
```

Run it with:

```bash
python sharepoint_sync.py
```

The script then:

1. initializes logging;
2. obtains a Microsoft Graph access token;
3. retrieves the SharePoint `Site ID`;
4. retrieves and displays available document libraries and their `Drive ID` values;
5. creates `/data_tree` and its subfolders when necessary;
6. uploads files from the local directory.

## Drive ID Output

`list_drive_ids()` requests:

```text
GET /sites/{site_id}/drives
```

Example console output:

```text
================================================================================
SHAREPOINT DRIVES
================================================================================

Name     : Documents
Drive ID : b!xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
Type     : documentLibrary
Web URL  : https://tenant.sharepoint.com/sites/site/Documents

================================================================================
```

The value required for a `Drive ID` field is the complete value displayed after `Drive ID`, often beginning with `b!`.

## SharePoint Library Used by the Script

The synchronization functions currently use Microsoft Graph endpoints in this form:

```text
/sites/{site_id}/drive/root:...
```

This means the script targets the site's default document library.

If the required destination is stored in another document library, use its `Drive ID` and change the relevant endpoints to:

```text
/drives/{drive_id}/root:...
```

For example:

```python
check_url = (
    f"https://graph.microsoft.com/v1.0/drives/{drive_id}"
    f"/root:{folder_path_encoded}"
)
```

In that design, the synchronization functions should receive `drive_id` as an argument instead of implicitly using the default library.

## Folder Structure

With:

```python
LOCAL_DIRECTORY = 'E:/workspace/EXTRACT/opt'
SHAREPOINT_BASE_FOLDER = '/data_tree'
```

A local file such as:

```text
E:/workspace/EXTRACT/opt/site1/config/settings.json
```

is uploaded to:

```text
/data_tree/site1/config/settings.json
```

The relative directory structure below `LOCAL_DIRECTORY` is therefore preserved under `SHAREPOINT_BASE_FOLDER`.

## Logging

Logs are written to:

```text
logs/synchro/
```

Each execution generates a timestamped file using this naming format:

```text
sharepoint_sync_YYYYMMDD_HHMMSS.log
```

The same log messages are also written to the console.

## Main Functions

### `setup_logging()`

Creates the log directory, configures file and console handlers, and returns the application's logger.

### `get_access_token()`

Uses `msal.ConfidentialClientApplication` and the Microsoft Graph `.default` scope to request an application access token.

### `get_site_id(access_token)`

Calls `SHAREPOINT_SITE_URL` and returns the SharePoint site identifier from the response.

### `list_drive_ids(access_token, site_id)`

Retrieves the site's document libraries and prints their name, `Drive ID`, drive type, and web URL.

### `ensure_folder_exists(access_token, site_id, folder_path)`

Checks whether the requested folder exists. Missing parent and target folders are created recursively.

### `upload_file(access_token, site_id, sharepoint_folder_path, local_file_path)`

Uploads one local file to the requested SharePoint folder.

### `sync_local_to_sharepoint(access_token, site_id, local_directory, sharepoint_folder)`

Walks through `local_directory`, recreates the folder hierarchy in SharePoint, and uploads each file.

### `delete_folder_contents(access_token, site_id, folder_path)`

Deletes the items returned from the specified SharePoint folder. This function is present in the script but is not called by the current main program.

## Current Synchronization Behavior

The current implementation should be considered an upload synchronization rather than a strict mirror:

- local directories are created remotely when required;
- local files are uploaded to the corresponding SharePoint path;
- existing remote files with the same destination may be overwritten by the upload operation;
- remote files that no longer exist locally are not automatically removed by the main synchronization flow.

## Important Limitations

### Large Files

The script currently uploads files using a simple Microsoft Graph content upload:

```http
PUT /content
```

For large files, an upload-session implementation is more robust because it can transfer the file in chunks and support resumable uploads.

### Pagination

`list_drive_ids()` currently processes the first response page. If Microsoft Graph returns an `@odata.nextLink`, the implementation would need to follow that link to process additional pages.

The same consideration applies when listing large collections of SharePoint items.

### File Name Encoding

Folder paths are encoded with `urllib.parse.quote`. File names containing reserved URL characters such as spaces, `#`, or `%` should also be encoded safely when building upload URLs.

### Error Handling and Retries

The current script logs HTTP errors but does not implement a general retry strategy for transient Microsoft Graph errors or throttling responses.

## HTML Copy/Paste Issue

Make sure the Python source contains the real `>` operator here:

```python
if len(path_parts) > 1:
```

If the Python file literally contains `&gt;`, it must be replaced with `>` or Python will report a syntax error. This can happen when code is copied through an HTML-formatted interface.

## Troubleshooting

### HTTP 401

Check:

- `CLIENT_ID`;
- `CLIENT_SECRET`;
- `TENANT_ID`;
- whether the client secret is still valid;
- whether an access token is successfully returned.

### HTTP 403

Check:

- Microsoft Graph application permissions;
- administrator consent;
- site-specific authorization if the application is intentionally restricted to selected SharePoint sites.

### HTTP 404 When Retrieving the Site

Verify the structure of `SHAREPOINT_SITE_URL`, for example:

```text
https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site-name}
```

### HTTP 404 for a Folder

Check:

- whether the folder is in the default document library;
- `SHAREPOINT_BASE_FOLDER`;
- the generated folder path;
- the `Drive ID` if the target folder is stored in another document library.

### No Drive IDs Are Displayed

Check that the application can successfully request:

```text
GET /sites/{site_id}/drives
```

and that the site exposes document libraries accessible to the application.

## Production Recommendations

Before using the script in production:

- move credentials to environment variables or a secret-management solution;
- add explicit timeouts to `requests` calls;
- add retry handling for transient failures and throttling;
- implement upload sessions for large files;
- handle Microsoft Graph pagination;
- safely encode file names used in URLs;
- test deletion behavior in a non-production SharePoint site;
- grant only the Microsoft Graph permissions required by the application.

## License

Internal project. Update this section according to your organization's licensing requirements.
