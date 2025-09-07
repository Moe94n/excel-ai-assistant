# Google Drive Integration Setup

This guide will help you set up Google Drive integration for the Excel AI Assistant.

## Prerequisites

1. A Google account with Google Drive access
2. Python environment with required dependencies installed

## Step 1: Create Google Cloud Project

1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the Google Drive API:
   - Go to "APIs & Services" > "Library"
   - Search for "Google Drive API"
   - Click "Enable"

## Step 2: Create OAuth 2.0 Credentials

1. Go to "APIs & Services" > "Credentials"
2. Click "Create Credentials" > "OAuth 2.0 Client IDs"
3. Configure the OAuth consent screen if prompted
4. Select "Desktop application" as the application type
5. Download the client secrets JSON file
6. Rename it to `client_secrets.json` and place it in the project root directory

## Step 3: Configure the Application

1. Open the Excel AI Assistant
2. Go to File > Cloud Storage > Cloud Storage Settings
3. Set the path to your `client_secrets.json` file
4. Click "Test Connection" to verify the setup

## Step 4: Authenticate

1. When you first try to access Google Drive, you'll be prompted to authenticate
2. A browser window will open asking you to grant permissions
3. Grant the requested permissions
4. The authentication token will be saved for future use

## Security Notes

- Keep your `client_secrets.json` file secure and never share it
- The application stores OAuth tokens locally for convenience
- You can revoke access anytime from your Google Account settings
- The application only requests minimal permissions needed for file operations

## Troubleshooting

### "Client secrets file not found"
- Ensure the `client_secrets.json` file is in the correct location
- Check the file path in the settings

### "Access denied" or "Invalid credentials"
- Try re-authenticating by deleting the token file and trying again
- Check that your Google account has access to the files you're trying to use

### "Daily limit exceeded"
- Google Drive API has usage limits for free accounts
- Consider upgrading to a paid plan if you need higher limits

## Supported File Types

The integration supports:
- Excel files (.xlsx, .xls)
- CSV files (.csv)
- Google Sheets (automatically converted to Excel format)

## Permissions Required

The application requests these Google Drive permissions:
- `https://www.googleapis.com/auth/drive.file` - View and manage files created by this app
- `https://www.googleapis.com/auth/drive.metadata.readonly` - View metadata for files in your Google Drive

These are minimal permissions focused on the files created or opened by the application.