# Google Drive File Access - Step-by-Step Guide

## Overview
This guide will help you enable Google Drive file access in your FIFA Draw Simulator application. Currently, your app uses Google Sheets API. To access files from Google Drive (including spreadsheets, CSV files, JSON files, etc.), you'll need to integrate the Google Drive API.

---

## Step 1: Set Up Google Cloud Console

### 1.1 Create/Select a Project
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Click on the project dropdown at the top
3. Either select an existing project or click "New Project"
4. If creating new: Enter a project name (e.g., "FIFA Draw App") and click "Create"

### 1.2 Enable Google Drive API
1. In the left sidebar, go to **APIs & Services** → **Library**
2. Search for "Google Drive API"
3. Click on "Google Drive API"
4. Click the **Enable** button
5. Wait for it to enable (may take a few seconds)

### 1.3 Enable Google Sheets API (if not already enabled)
- Your app still needs Google Sheets API for reading spreadsheet data
- Follow the same steps above but search for "Google Sheets API"
- Make sure it's enabled

---

## Step 2: Create OAuth 2.0 Credentials

**Why OAuth?** Google Drive API requires OAuth 2.0 authentication to access user files securely. An API key alone won't work for Drive files.

### 2.1 Create OAuth Client ID
1. Go to **APIs & Services** → **Credentials**
2. Click **+ CREATE CREDENTIALS** → **OAuth client ID**
3. If prompted, configure the OAuth consent screen first:
   - Choose **External** (unless you have a Google Workspace)
   - Fill in required fields:
     - App name: "FIFA Draw Simulator"
     - User support email: Your email
     - Developer contact: Your email
   - Click **Save and Continue** through the steps
   - Add scopes: `https://www.googleapis.com/auth/drive.readonly` (for reading files)
   - Add test users if needed (for testing)
   - Click **Save and Continue** → **Back to Dashboard**

### 2.2 Create OAuth Client ID (continued)
1. Application type: **Web application**
2. Name: "FIFA Draw Web Client"
3. **Authorized JavaScript origins:**
   - For local testing: `http://localhost:8000` (or your local port)
   - For production: `https://yourdomain.com`
4. **Authorized redirect URIs:**
   - For local testing: `http://localhost:8000` (or your local port)
   - For production: `https://yourdomain.com`
5. Click **Create**
6. **IMPORTANT:** Copy the **Client ID** (you'll need this in your code)
   - It looks like: `123456789-abcdefghijklmnop.apps.googleusercontent.com`

---

## Step 3: Understanding Authentication Methods

### Method A: OAuth 2.0 (Recommended)
- **Pros:** Works with private files, secure, proper user consent
- **Cons:** Requires user to sign in and grant permissions
- **Use when:** You need to access user's private files or files shared with them

### Method B: API Key (Limited)
- **Pros:** Simple, no user interaction needed
- **Cons:** Only works with publicly shared files
- **Use when:** Files are shared as "Anyone with the link can view"

**For this app, we'll implement OAuth 2.0** as it's more flexible and secure.

---

## Step 4: Extract File ID from Google Drive URL

When you have a Google Drive file URL, you need to extract the File ID:

**Google Sheets URL format:**
```
https://docs.google.com/spreadsheets/d/FILE_ID/edit
```

**Google Drive file URL format:**
```
https://drive.google.com/file/d/FILE_ID/view
```

**Shared link format:**
```
https://drive.google.com/open?id=FILE_ID
```

The FILE_ID is the long alphanumeric string between `/d/` and `/` or after `id=`.

---

## Step 5: File Sharing Settings

### For OAuth 2.0 (Recommended):
- Files can be private or shared
- User must grant access when signing in
- No need to change sharing settings

### For API Key Only:
- File must be shared publicly:
  1. Right-click the file in Google Drive
  2. Click **Share**
  3. Change to **"Anyone with the link can view"**
  4. Copy the link

---

## Step 6: Supported File Types

The Google Drive API can access various file types. For your app, you might want to:

1. **Google Sheets** (.gsheet) - Already supported via Sheets API
2. **CSV files** (.csv) - Can be downloaded and parsed
3. **JSON files** (.json) - Can be downloaded and parsed
4. **Text files** (.txt) - Can be downloaded and parsed
5. **Excel files** (.xlsx, .xls) - Can be downloaded (requires conversion)

---

## Step 7: Implementation Options

### Option 1: Use Google Drive API to Access Sheets
- Use Drive API to get file metadata
- Then use Sheets API to read the data (if it's a spreadsheet)

### Option 2: Use Google Drive API to Download Files
- Download CSV/JSON/TXT files directly
- Parse them in JavaScript

### Option 3: Hybrid Approach
- Try Sheets API first (for spreadsheets)
- Fall back to Drive API for other file types

---

## Step 8: Testing

1. **Test OAuth Flow:**
   - Open your app
   - Click "Connect to Google Drive"
   - Sign in with Google
   - Grant permissions
   - Verify access token is received

2. **Test File Access:**
   - Enter a Google Drive file URL
   - Verify file is loaded correctly
   - Check console for any errors

---

## Security Notes

1. **Never expose your Client Secret** in client-side code
2. **Client ID is safe** to expose in frontend code
3. **Access tokens** expire after 1 hour
4. **Refresh tokens** can be used to get new access tokens
5. For production, consider using a backend server to handle OAuth

---

## Troubleshooting

### "Access blocked: This app's request is invalid"
- Check that your redirect URI matches exactly in Google Cloud Console
- Check that your JavaScript origin matches your app URL

### "Error 403: Access Denied"
- File might not be shared properly
- User might not have granted necessary permissions
- Check OAuth scopes are correct

### "File not found"
- Verify the File ID is correct
- Check that the file exists and is accessible
- For API key method, ensure file is publicly shared

---

## Next Steps

After completing these steps, you'll need to:
1. Add the OAuth 2.0 code to your application (see code changes)
2. Update the UI to include Google Drive file selection
3. Handle file downloads and parsing based on file type
4. Test with different file types

See the updated `app.js` file for implementation details.








