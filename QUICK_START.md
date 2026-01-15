# Quick Start Guide - Google Drive Integration

## Summary

Your app now supports **two methods** to access Google files:

1. **Google Sheets API** (existing) - Uses API key, files must be public
2. **Google Drive API** (new) - Uses OAuth, works with private files, supports CSV/JSON/TXT

---

## Quick Setup Steps

### For Google Drive API (Recommended):

1. **Get OAuth Client ID:**
   - Go to https://console.cloud.google.com/
   - Create/select a project
   - Enable "Google Drive API"
   - Go to Credentials → Create Credentials → OAuth client ID
   - Application type: Web application
   - Add authorized origins: `http://localhost:8000` (or your domain)
   - Copy the Client ID

2. **In Your App:**
   - Select "Use Google Drive API (OAuth - Recommended)"
   - Paste your Client ID
   - Click "Connect to Google Drive"
   - Sign in with Google and grant permissions
   - Paste your Google Drive file URL
   - Continue to next step

### For Google Sheets API (Existing Method):

1. **Get API Key:**
   - Go to https://console.cloud.google.com/
   - Enable "Google Sheets API"
   - Create API key in Credentials
   - Make your sheet publicly shared

2. **In Your App:**
   - Select "Use Google Sheets API (API Key)"
   - Paste sheet URL and API key
   - Continue to next step

---

## Supported File Types (Google Drive API)

- ✅ Google Sheets (.gsheet)
- ✅ CSV files (.csv)
- ✅ JSON files (.json)
- ✅ Text files (.txt)

---

## File URL Formats Supported

- `https://drive.google.com/file/d/FILE_ID/view`
- `https://drive.google.com/open?id=FILE_ID`
- `https://docs.google.com/spreadsheets/d/FILE_ID/edit`
- Direct file ID: `FILE_ID`

---

## Troubleshooting

**"Access blocked" error:**
- Check that your redirect URI in Google Cloud Console matches your app URL exactly

**"Not authenticated" error:**
- Click "Connect to Google Drive" button first
- Make sure you've signed in and granted permissions

**"File not found" error:**
- Verify the file ID is correct
- For API key method, ensure file is publicly shared

---

## Next Steps

See `GOOGLE_DRIVE_SETUP.md` for detailed step-by-step instructions.








