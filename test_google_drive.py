#!/usr/bin/env python3
"""
Test script for Google Drive integration.
Run this to verify your Google Drive setup is working.
"""

import os
import sys

# Add the current directory to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app.config import AppConfig
from app.services.cloud.google_drive_service import GoogleDriveService
from app.utils.logger import setup_logger


def test_google_drive_integration():
    """Test Google Drive integration"""
    print("Testing Google Drive Integration")
    print("=" * 50)

    # Initialize config and logger
    config = AppConfig()
    logger = setup_logger("TestGoogleDrive")

    # Create Google Drive service
    gdrive = GoogleDriveService(config, logger)

    # Check if client secrets are configured
    client_secrets_path = config.get('google_client_secrets_path')
    if not client_secrets_path:
        print("❌ Google Drive client secrets not configured")
        print("Please run the application and go to:")
        print("File > Cloud Storage > Cloud Storage Settings")
        print("Then configure your Google Drive credentials.")
        return False

    if not os.path.exists(client_secrets_path):
        print(f"❌ Client secrets file not found: {client_secrets_path}")
        print("Please check the file path in settings.")
        return False

    print(f"✓ Client secrets file found: {client_secrets_path}")

    # Test authentication
    print("\nTesting authentication...")
    try:
        success = gdrive.authenticate()
        if success:
            print("✓ Authentication successful!")
        else:
            print("❌ Authentication failed")
            return False
    except Exception as e:
        print(f"❌ Authentication error: {str(e)}")
        return False

    # Test basic API calls
    print("\nTesting API access...")
    try:
        # Try to list files in root
        files = gdrive.list_files()
        print(f"✓ Successfully listed {len(files)} files/folders")

        # Show first few items
        if files:
            print("\nFirst few items:")
            for i, file in enumerate(files[:5]):
                file_type = "Folder" if file.get('mimeType') == 'application/vnd.google-apps.folder' else "File"
                print(f"  {i+1}. {file.get('name', 'Unknown')} ({file_type})")

    except Exception as e:
        print(f"❌ API access error: {str(e)}")
        return False

    print("\n" + "=" * 50)
    print("✅ Google Drive integration test PASSED!")
    print("You can now use Google Drive in the Excel AI Assistant.")
    print("=" * 50)

    return True


if __name__ == "__main__":
    success = test_google_drive_integration()
    sys.exit(0 if success else 1)