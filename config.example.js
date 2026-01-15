// ============================================================================
// Configuration Template for FIFA Draw Simulator
// ============================================================================
// Copy this file to 'config.js' and fill in your API keys
// ============================================================================

const APP_CONFIG = {
    // Google Sheets API Configuration
    googleSheets: {
        // Get your API key from: https://console.cloud.google.com/apis/credentials
        // 1. Enable Google Sheets API in your project
        // 2. Create an API key
        // 3. Paste it here (e.g., 'AIzaSy...')
        apiKey: '', 
        
        // Default sheet name/tab name to use (usually 'Sheet1')
        defaultSheetName: 'Participants'
    },

    // Google Drive API Configuration (Optional - only if using OAuth method)
    googleDrive: {
        // Get your OAuth Client ID from: https://console.cloud.google.com/apis/credentials
        // 1. Enable Google Drive API in your project
        // 2. Create OAuth 2.0 Client ID (Web application)
        // 3. Add authorized origins: http://localhost:8000 (or your domain)
        // 4. Paste Client ID here (e.g., '123456789-abc.apps.googleusercontent.com')
        clientId: ''
    },

    // Default settings (these populate the form when you open the app)
    defaults: {
        eventTitle: 'FIFA World Cup',  // Default event name
        numGroups: 8,                  // Default number of groups
        numPots: 4                     // Default number of pots
    }
};








