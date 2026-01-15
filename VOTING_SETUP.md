# Hackathon Team Voting Solution

## Overview

This voting solution allows you to create anonymous Google Forms for team voting at the conclusion of your hackathon event. The solution ensures that:

1. **Voting is anonymous** - Google Forms can be configured to collect anonymous responses
2. **Team members cannot vote for their own team** - Each form excludes the team of the participants using it
3. **Multiple forms are supported** - One form is created per team

## How It Works

The solution creates **one Google Form per team**. Each form:
- Excludes that team from the voting options
- Is assigned to all members of that team
- Allows team members to vote for all other teams

### Example:
- **Team A** members get a form that excludes Team A (they can vote for Teams B, C, D, etc.)
- **Team B** members get a form that excludes Team B (they can vote for Teams A, C, D, etc.)
- And so on...

## Usage Instructions

### Step 1: Complete the Draw

1. Run your draw simulation as usual
2. Complete the draw (either manually or using "INSTANT DRAW")
3. Wait for the "DRAW COMPLETE!" message

### Step 2: Enable Required APIs

**You already have OAuth credentials set up - you just need to enable two APIs!**

#### 2.1 Enable Google Drive API

**Why:** The app needs Drive API to create folders and move forms to the correct location.

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Make sure you're in the correct project (check the project dropdown at the top)
   - If you see project ID `17017615455` or similar in the error, that's your project
3. In the left sidebar, click **"APIs & Services"**
4. Click **"Library"** (or "Enabled APIs and services" → "Library")
5. In the search box at the top, type: **"Google Drive API"**
6. Click on **"Google Drive API"** from the search results
7. Click the blue **"ENABLE"** button
8. Wait for it to enable (you'll see a green checkmark or "API enabled" message)

#### 2.2 Enable Google Forms API

**Why:** The app needs Forms API to create the voting forms.

1. Still in the **"APIs & Services" → "Library"** section
2. In the search box, type: **"Google Forms API"**
3. Click on **"Google Forms API"** from the search results
4. Click the blue **"ENABLE"** button
5. Wait for it to enable

#### 2.3 Enable Google People API (Optional - for extracting emails from contact tags)

**Why:** If your spreadsheet uses Google contact tags (like @Amit Kashi), the app needs People API to extract email addresses from those tags.

1. Still in the **"APIs & Services" → "Library"** section
2. In the search box, type: **"People API"**
3. Click on **"People API"** from the search results
4. Click the blue **"ENABLE"** button
5. Wait for it to enable

**Note:** If you don't use contact tags and have a separate email column in your spreadsheet, you can skip this step.

**Important:** After enabling all APIs, wait 2-3 minutes for the changes to propagate through Google's systems before trying to create forms.

#### 2.4 Re-authenticate (One Time)

After enabling the APIs, you need to grant the new permissions:

1. In your app, click **"Connect to Google Drive"** (or the equivalent button)
2. Grant the new permissions when prompted
   - You'll see a consent screen asking for permissions
   - Make sure to check all the requested scopes (Drive, Sheets, Forms, Contacts)
3. That's it - you're done!

**Note:** If you enabled People API, you'll need to re-authenticate to get the contacts.readonly scope.

---

**Note:** If you DON'T have OAuth credentials yet, see section 2.5 below. Otherwise, skip it.

**Skip this if you already have OAuth credentials set up.**

1. In Google Cloud Console, go to **"APIs & Services"** → **"Credentials"**
2. If you see a warning about configuring the OAuth consent screen:
   - Click **"CONFIGURE CONSENT SCREEN"**
   - Choose **"External"** (unless you have Google Workspace)
   - Fill in required fields:
     - **App name:** "FIFA Draw Simulator" (or any name you prefer)
     - **User support email:** Your email address
     - **Developer contact information:** Your email address
   - Click **"SAVE AND CONTINUE"**
   - On the Scopes page, click **"SAVE AND CONTINUE"** (scopes will be requested automatically)
   - On the Test users page, click **"SAVE AND CONTINUE"** (you can add test users later if needed)
   - Click **"BACK TO DASHBOARD"**
3. Back on the Credentials page, click **"+ CREATE CREDENTIALS"**
4. Select **"OAuth client ID"**
5. If prompted, select **"Web application"** as the application type
6. Fill in:
   - **Name:** "FIFA Draw Web Client" (or any name)
   - **Authorized JavaScript origins:** 
     - Click **"+ ADD URI"**
     - Enter: `http://localhost:8000` (or your domain if deployed)
   - **Authorized redirect URIs:**
     - Click **"+ ADD URI"**
     - Enter: `http://localhost:8000` (or your domain if deployed)
7. Click **"CREATE"**
8. **Copy the Client ID** (it looks like: `123456789-abc.apps.googleusercontent.com`)
   - You can also copy it later from the Credentials page

#### 2.5 Get OAuth Client ID (ONLY if you don't have one yet - SKIP if you already have credentials)

**Skip this entire section if you already have OAuth credentials!**

1. In Google Cloud Console, go to **"APIs & Services"** → **"Credentials"**
2. If you see a warning about configuring the OAuth consent screen:
   - Click **"CONFIGURE CONSENT SCREEN"**
   - Choose **"External"** (unless you have Google Workspace)
   - Fill in required fields (App name, support email, etc.)
   - Click through the screens and save
3. Back on the Credentials page, click **"+ CREATE CREDENTIALS"** → **"OAuth client ID"**
4. Select **"Web application"**
5. Fill in:
   - **Name:** "FIFA Draw Web Client"
   - **Authorized JavaScript origins:** `http://localhost:8000`
   - **Authorized redirect URIs:** `http://localhost:8000`
6. Click **"CREATE"** and copy the Client ID
7. Enter it in your app's Google Drive Client ID field

### Step 3: Generate and Create Voting Forms

1. After the draw is complete, you'll see a new button: **"GENERATE VOTING FORMS"**
2. Click this button
3. **If authenticated:** Forms will be created automatically!
   - The app will create all forms using the Google Forms API
   - You'll see a list of created forms with links
4. **If not authenticated:** You'll see instructions for manual creation
   - Follow the instructions to create forms manually

### Step 4: Anonymous Settings (Automatically Configured)

**✅ GOOD NEWS:** Anonymous settings are now automatically configured when forms are created!

The following settings are applied automatically:
- **Email collection:** DISABLED (ensures anonymous responses)
- **Login requirement:** DISABLED (allows anonymous voting)
- **Response editing:** DISABLED (prevents vote changes)
- **Summary charts:** HIDDEN (maintains anonymity during voting)
- **Progress bar:** ENABLED (better user experience)

**No manual configuration required!** Forms are ready for anonymous voting immediately after creation.

**Note:** If you see any warnings in the browser console about settings not being supported by the API, you may need to manually verify the settings:
1. Click the "Edit Form" link
2. Check Settings (gear icon) → Responses
3. Ensure "Collect email addresses" is set to "Do not collect"
4. Ensure other privacy settings are configured as needed

### Step 5: Distribute Forms

1. Share the appropriate form link with each team's members
2. You can:
   - Send links via email
   - Post links in a shared document
   - Share links in a messaging app
   - Create a simple webpage with links

### Step 6: Collect Votes

1. After the hackathon concludes, participants vote using their assigned forms
2. Responses are automatically collected in Google Forms
3. You can view results in the Google Forms response sheet

## Important Notes

### Automatic Form Creation

- Forms are created automatically using the Google Forms API
- You need Google OAuth authentication set up (same as Google Drive integration)
- The Google Forms API must be enabled in your Google Cloud project
- Forms are created with questions pre-configured AND anonymous settings automatically applied
- No manual configuration required - forms are ready to use immediately

### Anonymity

- Google Forms supports anonymous responses when configured correctly
- **Anonymous settings are automatically configured** during form creation
- Email collection and login requirements are automatically disabled
- Summary charts are hidden to prevent participants from seeing results before voting closes
- No manual configuration required - forms are anonymous by default

### Team Member Verification

- The solution assumes that names in your spreadsheet match participant names/tags
- Make sure names are consistent between the draw and the voting forms
- You may want to verify participant assignments before distributing forms

### Multiple Forms

- Each team gets its own form
- This ensures team members cannot vote for their own team
- You can handle distributing multiple form links as needed

## Exporting Configuration

The voting solution allows you to:
- **Export Configuration (JSON)**: Download the complete voting configuration as a JSON file
- **Copy Report**: Copy the summary report to your clipboard for easy sharing

## Troubleshooting

### "Google Drive API has not been used" or "API is disabled" Error

**This means Google Drive API is not enabled in your project.**

**Solution:**
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Make sure you're in the correct project (check project dropdown at top)
3. Go to **APIs & Services** → **Library**
4. Search for **"Google Drive API"**
5. Click on it and click **"ENABLE"**
6. Wait 2-3 minutes, then try again

**Quick Link:** If you see a project ID in the error message, use this link (replace YOUR_PROJECT_ID):
```
https://console.developers.google.com/apis/api/drive.googleapis.com/overview?project=YOUR_PROJECT_ID
```

### "Draw state is empty" Error
- Make sure you've completed the draw before generating voting forms
- The draw must be marked as complete

### Missing Teams
- Verify that your draw has been completed successfully
- Check that groups have entries assigned to them

### Form Not Appearing
- Refresh the page after completing the draw
- Check the browser console for any errors

### Authorization Dialog Appears Every Time
- This should be fixed in the latest version
- The app now reuses your existing token
- If it still happens, check browser console for errors

### Forms Not in Folder
- Check browser console (F12) for error messages
- Look for messages starting with "=== FOLDER CREATION DEBUG ==="
- Verify your spreadsheet URL is detected correctly
- Forms might be in your Drive root folder if folder creation failed

## Technical Details

The voting solution:
- Reads from `drawState.groups` to get team assignments
- Uses `config.groupNames` to identify teams
- Creates form configurations that exclude each team from its own form
- Generates participant assignments based on team membership

## Support

If you encounter issues:
1. Check the browser console for error messages
2. Verify that the draw is complete
3. Ensure names in your spreadsheet match participant names/tags
4. Make sure the voting-solution.js file is loaded (check the HTML file)
