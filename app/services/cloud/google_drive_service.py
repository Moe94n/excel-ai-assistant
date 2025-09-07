"""
Google Drive integration service for Excel AI Assistant.
Handles OAuth2 authentication and file operations with Google Drive API.
"""

import os
import io
import json
import pickle
import webbrowser
from pathlib import Path
from typing import List, Dict, Optional, Tuple, Any
from urllib.parse import urlparse, parse_qs

import requests
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from googleapiclient.errors import HttpError

from app.config import AppConfig
from app.utils.logger import setup_logger


class GoogleDriveService:
    """Service for Google Drive integration with OAuth2 authentication"""

    # Google Drive API scopes
    SCOPES = [
        'https://www.googleapis.com/auth/drive.file',
        'https://www.googleapis.com/auth/drive.metadata.readonly'
    ]

    # MIME types for Excel/CSV files
    EXCEL_MIME_TYPES = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',  # .xlsx
        'application/vnd.ms-excel',  # .xls
        'text/csv',  # .csv
        'application/csv'  # Alternative CSV MIME type
    ]

    def __init__(self, config: AppConfig, logger=None):
        """Initialize Google Drive service"""
        self.config = config
        self.logger = logger or setup_logger("GoogleDriveService")
        self.service = None
        self.credentials = None

        # Get token storage path
        self.token_path = self._get_token_path()

        # Load credentials if available
        self._load_credentials()

    def _get_token_path(self) -> Path:
        """Get the path for storing OAuth tokens"""
        # Use the same config directory as the main app
        config_dir = Path.home() / 'AppData' / 'Roaming' / 'ExcelAIAssistant'
        config_dir.mkdir(parents=True, exist_ok=True)
        return config_dir / 'google_drive_token.pickle'

    def _load_credentials(self) -> bool:
        """Load existing credentials from token file"""
        try:
            if self.token_path.exists():
                with open(self.token_path, 'rb') as token:
                    self.credentials = pickle.load(token)

                # Check if credentials are valid
                if self.credentials and self.credentials.valid:
                    self.service = build('drive', 'v3', credentials=self.credentials)
                    self.logger.info("Google Drive credentials loaded successfully")
                    return True
                elif self.credentials and self.credentials.expired and self.credentials.refresh_token:
                    # Refresh expired credentials
                    self.logger.info("Refreshing expired Google Drive credentials")
                    self.credentials.refresh(Request())
                    self._save_credentials()
                    self.service = build('drive', 'v3', credentials=self.credentials)
                    return True
                else:
                    self.logger.warning("Invalid or expired Google Drive credentials")
                    return False
            else:
                self.logger.info("No Google Drive credentials found")
                return False
        except Exception as e:
            self.logger.error(f"Error loading Google Drive credentials: {str(e)}")
            return False

    def _save_credentials(self):
        """Save credentials to token file"""
        try:
            with open(self.token_path, 'wb') as token:
                pickle.dump(self.credentials, token)
            self.logger.info("Google Drive credentials saved")
        except Exception as e:
            self.logger.error(f"Error saving Google Drive credentials: {str(e)}")

    def authenticate(self, client_secrets_path: str = None) -> bool:
        """
        Authenticate with Google Drive using OAuth2

        Args:
            client_secrets_path: Path to client secrets JSON file (optional)

        Returns:
            bool: True if authentication successful
        """
        try:
            # If no client secrets path provided, look for it in config
            if not client_secrets_path:
                client_secrets_path = self.config.get('google_client_secrets_path')

            if not client_secrets_path or not os.path.exists(client_secrets_path):
                error_msg = (
                    "Google Drive client secrets file not found. "
                    "Please set up Google Drive integration by following the guide at: "
                    "docs/google_drive_setup.md"
                )
                self.logger.error(error_msg)

                # Try to provide a helpful message to the user
                print("\n" + "="*60)
                print("GOOGLE DRIVE SETUP REQUIRED")
                print("="*60)
                print("To use Google Drive integration, you need to:")
                print("1. Go to https://console.cloud.google.com/")
                print("2. Create a new project or select existing one")
                print("3. Enable Google Drive API")
                print("4. Create OAuth 2.0 credentials")
                print("5. Download client_secrets.json")
                print("6. Place it in your project directory")
                print("7. Set the path in application settings")
                print("\nSee docs/google_drive_setup.md for detailed instructions")
                print("="*60 + "\n")

                return False

            # Validate that the file is a proper JSON file
            try:
                with open(client_secrets_path, 'r') as f:
                    secrets_data = json.load(f)

                # Check if it has the required structure
                if 'installed' not in secrets_data and 'web' not in secrets_data:
                    self.logger.error("Invalid client secrets file format")
                    return False

            except json.JSONDecodeError:
                self.logger.error("Client secrets file is not valid JSON")
                return False
            except Exception as e:
                self.logger.error(f"Error reading client secrets file: {str(e)}")
                return False

            # Create OAuth2 flow
            flow = InstalledAppFlow.from_client_secrets_file(
                client_secrets_path, self.SCOPES
            )

            print("\n" + "="*60)
            print("GOOGLE DRIVE AUTHENTICATION")
            print("="*60)
            print("A browser window will open for Google authentication.")
            print("Please grant the requested permissions.")
            print("You can close this window after authentication.")
            print("="*60 + "\n")

            # Run local server for authentication
            self.credentials = flow.run_local_server(port=0)

            # Save credentials
            self._save_credentials()

            # Build service
            self.service = build('drive', 'v3', credentials=self.credentials)

            self.logger.info("Google Drive authentication successful")
            print("\n✓ Google Drive authentication successful!\n")
            return True

        except Exception as e:
            error_msg = f"Google Drive authentication failed: {str(e)}"
            self.logger.error(error_msg)
            print(f"\n✗ {error_msg}")
            print("Please check your client secrets file and try again.\n")
            return False

    def is_authenticated(self) -> bool:
        """Check if user is authenticated with Google Drive"""
        return self.service is not None and self.credentials is not None

    def logout(self):
        """Logout and clear stored credentials"""
        try:
            if self.token_path.exists():
                self.token_path.unlink()
            self.credentials = None
            self.service = None
            self.logger.info("Google Drive logout successful")
        except Exception as e:
            self.logger.error(f"Error during Google Drive logout: {str(e)}")

    def list_files(self, folder_id: Optional[str] = None) -> List[Dict[str, Any]]:
        """
        List files in a Google Drive folder (simplified version for cloud manager)

        Args:
            folder_id: ID of the folder to list (None for root)

        Returns:
            List of file dictionaries
        """
        if not self.is_authenticated():
            self.logger.error("Not authenticated with Google Drive")
            return []

        try:
            # Build query
            queries = []

            # Add folder filter
            if folder_id:
                queries.append(f"'{folder_id}' in parents")
            else:
                queries.append("'root' in parents")

            # Add trashed filter (exclude deleted files)
            queries.append("trashed=false")

            # Combine queries
            full_query = ' and '.join(queries)

            # List files and folders
            results = self.service.files().list(
                q=full_query,
                pageSize=1000,
                fields="nextPageToken, files(id, name, mimeType, size, modifiedTime, parents)",
                orderBy="folder, name"
            ).execute()

            files = results.get('files', [])

            self.logger.debug(f"Listed {len(files)} items from Google Drive folder {folder_id or 'root'}")
            return files

        except HttpError as e:
            self.logger.error(f"Google Drive API error: {str(e)}")
            return []
        except Exception as e:
            self.logger.error(f"Error listing Google Drive files: {str(e)}")
            return []

    def list_files_detailed(self, folder_id: str = 'root', mime_types: List[str] = None,
                           query: str = None) -> List[Dict[str, Any]]:
        """
        List files in a Google Drive folder (detailed version)

        Args:
            folder_id: ID of the folder to list (default: root)
            mime_types: List of MIME types to filter by
            query: Additional query string

        Returns:
            List of file dictionaries
        """
        if not self.is_authenticated():
            self.logger.error("Not authenticated with Google Drive")
            return []

        try:
            # Build query
            queries = []

            # Add folder filter
            if folder_id == 'root':
                queries.append("'root' in parents")
            else:
                queries.append(f"'{folder_id}' in parents")

            # Add MIME type filter
            if mime_types:
                mime_query = ' or '.join([f"mimeType='{mt}'" for mt in mime_types])
                queries.append(f"({mime_query})")

            # Add trashed filter (exclude deleted files)
            queries.append("trashed=false")

            # Add custom query
            if query:
                queries.append(query)

            # Combine queries
            full_query = ' and '.join(queries)

            # List files
            results = self.service.files().list(
                q=full_query,
                pageSize=100,
                fields="nextPageToken, files(id, name, mimeType, size, modifiedTime, parents, webViewLink)",
                orderBy="modifiedTime desc"
            ).execute()

            files = results.get('files', [])

            # Get next page if available
            while 'nextPageToken' in results:
                results = self.service.files().list(
                    q=full_query,
                    pageSize=100,
                    pageToken=results['nextPageToken'],
                    fields="nextPageToken, files(id, name, mimeType, size, modifiedTime, parents, webViewLink)",
                    orderBy="modifiedTime desc"
                ).execute()
                files.extend(results.get('files', []))

            self.logger.info(f"Listed {len(files)} files from Google Drive")
            return files

        except HttpError as e:
            self.logger.error(f"Google Drive API error: {str(e)}")
            return []
        except Exception as e:
            self.logger.error(f"Error listing Google Drive files: {str(e)}")
            return []

    def list_excel_files(self, folder_id: str = 'root') -> List[Dict[str, Any]]:
        """List Excel and CSV files in a Google Drive folder"""
        return self.list_files(folder_id, self.EXCEL_MIME_TYPES)

    def download_file(self, file_id: str, local_path: str) -> bool:
        """
        Download a file from Google Drive

        Args:
            file_id: Google Drive file ID
            local_path: Local path to save the file

        Returns:
            bool: True if download successful
        """
        if not self.is_authenticated():
            self.logger.error("Not authenticated with Google Drive")
            return False

        try:
            # Get file metadata
            file_metadata = self.service.files().get(fileId=file_id).execute()
            file_name = file_metadata.get('name', 'downloaded_file')

            # Create download request
            request = self.service.files().get_media(fileId=file_id)

            # Download file
            with io.FileIO(local_path, 'wb') as fh:
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while done is False:
                    status, done = downloader.next_chunk()
                    self.logger.debug(f"Download progress: {int(status.progress() * 100)}%")

            self.logger.info(f"Downloaded file '{file_name}' from Google Drive")
            return True

        except HttpError as e:
            self.logger.error(f"Google Drive API error during download: {str(e)}")
            return False
        except Exception as e:
            self.logger.error(f"Error downloading file from Google Drive: {str(e)}")
            return False

    def download_file_content(self, file_id: str) -> Optional[bytes]:
        """
        Download a file from Google Drive and return content as bytes

        Args:
            file_id: Google Drive file ID

        Returns:
            bytes: File content, or None if failed
        """
        if not self.is_authenticated():
            self.logger.error("Not authenticated with Google Drive")
            return None

        try:
            # Get file metadata first
            file_metadata = self.service.files().get(fileId=file_id, fields='name,mimeType').execute()
            mime_type = file_metadata.get('mimeType', '')

            # Handle Google Workspace files (Docs, Sheets, etc.)
            if mime_type.startswith('application/vnd.google-apps.'):
                if mime_type == 'application/vnd.google-apps.spreadsheet':
                    # Export Google Sheets to Excel
                    request = self.service.files().export(
                        fileId=file_id,
                        mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                else:
                    self.logger.error(f"Unsupported Google Workspace file type: {mime_type}")
                    return None
            else:
                # Regular file download
                request = self.service.files().get_media(fileId=file_id)

            # Download the file content
            file_content = io.BytesIO()
            downloader = MediaIoBaseDownload(file_content, request)

            done = False
            while done is False:
                status, done = downloader.next_chunk()
                self.logger.debug(f"Download progress: {int(status.progress() * 100)}%")

            self.logger.info(f"Successfully downloaded file content: {file_metadata.get('name', 'Unknown')}")
            return file_content.getvalue()

        except HttpError as e:
            self.logger.error(f"Google Drive API error downloading file content: {str(e)}")
            return None
        except Exception as e:
            self.logger.error(f"Error downloading file content from Google Drive: {str(e)}")
            return None

    def upload_file(self, local_path: str, folder_id: str = 'root',
                   file_name: str = None) -> Optional[str]:
        """
        Upload a file to Google Drive

        Args:
            local_path: Local path of file to upload
            folder_id: Google Drive folder ID to upload to
            file_name: Name for the uploaded file (optional)

        Returns:
            str: File ID of uploaded file, or None if failed
        """
        if not self.is_authenticated():
            self.logger.error("Not authenticated with Google Drive")
            return None

        try:
            # Get file info
            if not os.path.exists(local_path):
                self.logger.error(f"Local file not found: {local_path}")
                return None

            if not file_name:
                file_name = os.path.basename(local_path)

            # Determine MIME type
            file_extension = os.path.splitext(local_path)[1].lower()
            mime_type = 'application/octet-stream'  # Default

            if file_extension == '.xlsx':
                mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            elif file_extension == '.xls':
                mime_type = 'application/vnd.ms-excel'
            elif file_extension == '.csv':
                mime_type = 'text/csv'

            # Create file metadata
            file_metadata = {
                'name': file_name,
                'parents': [folder_id]
            }

            # Upload file
            media = MediaIoBaseUpload(
                io.FileIO(local_path, 'rb'),
                mimetype=mime_type,
                resumable=True
            )

            file = self.service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id'
            ).execute()

            file_id = file.get('id')
            self.logger.info(f"Uploaded file '{file_name}' to Google Drive with ID: {file_id}")
            return file_id

        except HttpError as e:
            self.logger.error(f"Google Drive API error during upload: {str(e)}")
            return None
        except Exception as e:
            self.logger.error(f"Error uploading file to Google Drive: {str(e)}")
            return None

    def create_share_link(self, file_id: str) -> Optional[str]:
        """
        Create a shareable link for a Google Drive file

        Args:
            file_id: Google Drive file ID

        Returns:
            str: Shareable link, or None if failed
        """
        if not self.is_authenticated():
            self.logger.error("Not authenticated with Google Drive")
            return None

        try:
            # Create public permission
            permission = {
                'type': 'anyone',
                'role': 'reader'
            }

            self.service.permissions().create(
                fileId=file_id,
                body=permission
            ).execute()

            # Get web view link
            file_metadata = self.service.files().get(
                fileId=file_id,
                fields='webViewLink'
            ).execute()

            share_link = file_metadata.get('webViewLink')
            self.logger.info(f"Created share link for file {file_id}")
            return share_link

        except HttpError as e:
            self.logger.error(f"Google Drive API error creating share link: {str(e)}")
            return None
        except Exception as e:
            self.logger.error(f"Error creating share link: {str(e)}")
            return None

    def get_file_info(self, file_id: str) -> Optional[Dict[str, Any]]:
        """
        Get detailed information about a Google Drive file

        Args:
            file_id: Google Drive file ID

        Returns:
            Dict containing file information, or None if failed
        """
        if not self.is_authenticated():
            self.logger.error("Not authenticated with Google Drive")
            return None

        try:
            file_info = self.service.files().get(
                fileId=file_id,
                fields='id, name, mimeType, size, modifiedTime, createdTime, parents, webViewLink, thumbnailLink'
            ).execute()

            return file_info

        except HttpError as e:
            self.logger.error(f"Google Drive API error getting file info: {str(e)}")
            return None
        except Exception as e:
            self.logger.error(f"Error getting file info: {str(e)}")
            return None

    def create_folder(self, name: str, parent_id: str = 'root') -> Optional[str]:
        """
        Create a new folder in Google Drive

        Args:
            name: Name of the folder
            parent_id: Parent folder ID

        Returns:
            str: Folder ID, or None if failed
        """
        if not self.is_authenticated():
            self.logger.error("Not authenticated with Google Drive")
            return None

        try:
            folder_metadata = {
                'name': name,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [parent_id]
            }

            folder = self.service.files().create(
                body=folder_metadata,
                fields='id'
            ).execute()

            folder_id = folder.get('id')
            self.logger.info(f"Created folder '{name}' with ID: {folder_id}")
            return folder_id

        except HttpError as e:
            self.logger.error(f"Google Drive API error creating folder: {str(e)}")
            return None
        except Exception as e:
            self.logger.error(f"Error creating folder: {str(e)}")
            return None