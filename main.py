import os
import pickle
import re
import io
import argparse
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow

# If modifying the file permission, we need the 'https://www.googleapis.com/auth/drive' scope
SCOPES = ['https://www.googleapis.com/auth/drive']

def authenticate_google_account():
    """Authenticate and create the Google Drive API client."""
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('drive', 'v3', credentials=creds)
    return service

def list_files_in_folder(service, folder_id):
    """Lists all files inside the given Google Drive folder."""
    try:
        query = f"'{folder_id}' in parents"
        results = service.files().list(
            q=query,
            fields="files(id, name, mimeType, parents, trashed)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True
        ).execute()

        # Exclude trashed files
        return list(filter(lambda file: not file['trashed'], results.get('files', [])))

    except Exception as e:
        print(f"An error occurred while listing files: {e}")
        return []

def list_files_in_folder_recursive(service, folder_id):
    """Lists all files inside the given Google Drive folder and its subfolders."""
    try:
        items = list_files_in_folder(service, folder_id)

        for item in items:
            if item['mimeType'] == 'application/vnd.google-apps.folder':
                # Recursively list files in subfolders
                subfolder_items = list_files_in_folder_recursive(service, item['id'])
                items.extend(subfolder_items)

        return items

    except Exception as e:
        print(f"An error occurred while listing files: {e}")
        return []

def export_drive_files(service, source_folder_id, target_folder_id, recursive=False, file_regex=None):
    """Exports Google Docs and Slides to PDF and copies other files to the target folder."""
    try:
        files = list_files_in_folder_recursive(service, source_folder_id) if recursive else list_files_in_folder(service, source_folder_id)

        if file_regex is not None:
            files = list(filter(lambda file: re.match(file_regex, file['name']), files))

        if not files:
            print("No matching files found in the source folder.")
            return

        for file in files:
            mime_type = file['mimeType']

            if mime_type == 'application/vnd.google-apps.folder':
                # Do not process Google Drive folders
                continue

            print(f"Processing {file['name']}...")

            if mime_type == 'application/vnd.google-apps.document':
                # Export Google Docs to PDF
                request = service.files().export_media(fileId=file['id'], mimeType='application/pdf')
                pdf_data = io.BytesIO(request.execute())

                pdf_metadata = {
                    'name': f"{file['name']}.pdf",
                    'parents': [target_folder_id]
                }

                media = MediaIoBaseUpload(pdf_data, mimetype='application/pdf', resumable=True)
                service.files().create(body=pdf_metadata, media_body=media, supportsAllDrives=True).execute()
                print(f"Exported and uploaded {file['name']} as a PDF.")

            elif mime_type == 'application/vnd.google-apps.presentation':
                # Export Google Slides to PDF
                request = service.files().export_media(fileId=file['id'], mimeType='application/pdf')
                pdf_data = io.BytesIO(request.execute())

                pdf_metadata = {
                    'name': f"{file['name']}.pdf",
                    'parents': [target_folder_id]
                }

                media = MediaIoBaseUpload(pdf_data, mimetype='application/pdf', resumable=True)
                service.files().create(body=pdf_metadata, media_body=media, supportsAllDrives=True).execute()
                print(f"Exported and uploaded {file['name']} as a PDF.")

            else:
                # Copy other files directly
                copied_file_metadata = {
                    'name': file['name'],
                    'parents': [target_folder_id]
                }
                
                service.files().copy(
                    fileId=file['id'],
                    body=copied_file_metadata,
                    supportsAllDrives=True
                ).execute()

                print(f"Copied {file['name']} to the target folder.")

    except Exception as e:
        print(f"An error occurred while processing files: {e}")

def create_shareable_link(service, file_id):
    """Creates a shareable link to the file with public access."""
    try:
        # permission = {
        #     'type': 'domain',
        #     'role': 'reader',
        #     'domain': 'fysiksektionen.se'
        # }
        #
        # service.permissions().create(fileId=file_id, body=permission).execute()
        file = service.files().get(fileId=file_id, fields="webViewLink", supportsAllDrives=True).execute()
        return file.get('webViewLink')
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

def sort_and_indent_files(files):
    """Sorts files into alphabetical and hierarchical order with proper indentation."""
    def parse_numbering(name):
        match = re.match(r"^((\d+\.?)+)\s", name)
        if match:
            return match.group(1).split('.')
        return []

    unnumbered_files = []
    numbered_files = []

    for file in files:
        if parse_numbering(file['name']):
            numbered_files.append(file)
        else:
            unnumbered_files.append(file)

    unnumbered_files.sort(key=lambda x: x['name'].lower())

    def sort_key(file):
        numbering = parse_numbering(file['name'])
        return ([int(part) for part in numbering if part.isdigit()], file['name'].lower())

    numbered_files.sort(key=sort_key)
    sorted_files = unnumbered_files + numbered_files

    toc_lines = []
    for file in sorted_files:
        numbering = parse_numbering(file['name'])
        indent = '    ' * max(len(numbering) - 2, 0)
        name_without_extension = re.sub(r"\.[^.]+$", "", file['name'])
        link = f"{indent}- [{name_without_extension}]({file['link']})"
        toc_lines.append(link)

    return toc_lines

def create_links_for_drive_folder(service, folder_id, recursive=False, file_regex=None):
    """Creates a nicely formatted Table of Contents for all matching files in a Google Drive folder."""
    try:
        files = list_files_in_folder_recursive(service, folder_id) if recursive else list_files_in_folder(service, folder_id)

        if file_regex is not None:
            files = list(filter(lambda file: re.match(file_regex, file['name']), files))

        if not files:
            print("No files found to process.")
            return

        for file in files:
            print(f"Creating shareable link for: {file['name']}...")
            link = create_shareable_link(service, file['id'])
            if link:
                file['link'] = link
            else:
                print(f"Failed to create a shareable link for {file['name']}.")
                file['link'] = ""

        folder_link = create_shareable_link(service, folder_id)
        toc_lines = sort_and_indent_files(files)

        print("\nBilagor:\n")
        if folder_link:
            print(f"[Samtliga bilagor]({folder_link})\n")
        if toc_lines:
            print("\n".join(toc_lines))
        else:
            print("No shareable links were created.")

    except Exception as e:
        print(f"An error occurred while processing the folder: {e}")

def print_files_in_drive_folder(service, folder_id, recursive=False, file_regex=None, print_dirs=False):
    """Prints a a list of files in a Google Drive folder."""
    files = list_files_in_folder_recursive(service, folder_id) if recursive else list_files_in_folder(service, folder_id)

    if file_regex is not None:
        files = list(filter(lambda file: re.match(file_regex, file['name']), files))

    # Filter out directories
    if not print_dirs:
        files = list(filter(lambda file: file['mimeType'] != 'application/vnd.google-apps.folder', files))

    file_names = sorted(map(lambda file: file['name'], files))

    for file_name in file_names:
        print(file_name)

def main():
    service = authenticate_google_account()

    parser = argparse.ArgumentParser(description="Attachments helper")
    parser.add_argument("operation", choices=["table", "pdfs", "print"], help="Operation to perform: 'table' to generate a Table of Contents, 'print' to print the files or 'pdfs' to export files")
    parser.add_argument("source_folder", help="Source Google Drive folder ID")
    parser.add_argument("destination_folder", nargs="?", help="Target Google Drive folder ID (required for 'pdfs' operation)")
    parser.add_argument("-r", "--recursive", action="store_true", help="Include subdirectories in the operation")
    parser.add_argument("-x", "--regex", nargs="?", help="Optional regex to copy only some files for 'pdfs'", default=None)
    parser.add_argument("-d", "--print-dirs", action="store_true", help="Print both files and directories for 'print'")

    args = parser.parse_args()

    if args.operation == "table":
        create_links_for_drive_folder(service, args.source_folder, args.recursive, args.regex)
    elif args.operation == "print":
        print_files_in_drive_folder(service, args.source_folder, args.recursive, args.regex, args.print_dirs)
    elif args.operation == "pdfs":
        if not args.destination_folder:
            print("Error: Destination folder ID is required for the 'pdfs' operation.")
            return
        export_drive_files(service, args.source_folder, args.destination_folder, args.recursive, args.regex)

if __name__ == '__main__':
    main()
