// Custom Draw Simulator - Fully Configurable

// ==================== SECRET CHEAT CONFIGURATION ====================
// This section is YOUR SECRET SETUP - only visible in the code!
//
// CONSTRAINT TYPES:
// 1. "cannotBeWith" - Two entries cannot be in the same group
// 2. "mustBeWith" - Two entries MUST be in the same group
// 3. "mustBeInGroup" - Force an entry into a specific group
// 4. "cannotBeInGroup" - Prevent an entry from being in a specific group
//
// HOW TO USE:
// - Names must match EXACTLY as you enter them in the UI (case-sensitive)
// - Add your constraints to the CHEAT_CONSTRAINTS array below
// - The draw will secretly respect these rules while appearing random

const CHEAT_CONSTRAINTS = {
    enabled: true,  // Set to false to disable all constraints

    // Entries that CANNOT be in the same group
    // Loaded from "CannotBeWith" sheet (columns A and B, each row is a pair)
    // Example: ["Brazil", "Argentina"] means they will never be in the same group
    cannotBeWith: [
        // Will be loaded from CannotBeWith sheet
    ],

    // Entries that MUST be in the same group
    // Loaded from "MustBeWith" sheet (columns A and B, each row is a pair)
    // Example: ["USA", "Canada"] means they will always be together
    mustBeWith: [
        // Will be loaded from MustBeWith sheet
    ],

    // Force specific entries into specific groups
    // Example: { "Brazil": "A" } forces Brazil into Group A
    mustBeInGroup: {
        // "EntryName": "GroupName",
    },

    // Prevent entries from being in specific groups
    // Example: { "Brazil": ["B", "C"] } prevents Brazil from being in Group B or C
    cannotBeInGroup: {
        // "EntryName": ["GroupName1", "GroupName2"],
    },
};

// ==================== END SECRET CONFIGURATION ====================

// ==================== GOOGLE SHEETS INTEGRATION ====================

// Extract spreadsheet ID from Google Sheets URL
function extractSpreadsheetId(url) {
    const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : null;
}

// Fetch data from Google Sheets
async function fetchGoogleSheetData(apiKey, spreadsheetId, sheetName = 'Participants') {
    try {
        const range = `${sheetName}!A1:Z1000`; // Fetch a large range
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}?key=${apiKey}`;

        console.log('Fetching from URL:', url.replace(apiKey, 'API_KEY_HIDDEN'));
        const response = await fetch(url);

        if (!response.ok) {
            const errorData = await response.json();
            console.error('Google Sheets API error:', errorData);
            throw new Error(errorData.error?.message || 'Failed to fetch Google Sheets data');
        }

        const data = await response.json();
        console.log('Raw Google Sheets response:', data);
        console.log('Values array:', data.values);
        console.log('Number of rows:', data.values?.length || 0);
        return data.values || [];
    } catch (error) {
        console.error('Google Sheets fetch error:', error);
        throw error;
    }
}

// Write data to Google Sheets with formatting using OAuth2
async function writeToGoogleSheetOAuth(accessToken, spreadsheetId, sheetName, data, colors) {
    try {
        // First, get sheet metadata to find sheet ID
        const metadataUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}`;
        const metadataResponse = await fetch(metadataUrl, {
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });
        
        if (!metadataResponse.ok) {
            const errorData = await metadataResponse.json();
            throw new Error(errorData.error?.message || 'Failed to get spreadsheet metadata');
        }
        
        const metadata = await metadataResponse.json();
        let targetSheetId = 0; // Default to first sheet
        
        // Find or create the target sheet
        const existingSheet = metadata.sheets?.find(s => s.properties.title === sheetName);
        if (existingSheet) {
            targetSheetId = existingSheet.properties.sheetId;
        } else {
            // Create new sheet
            const createSheetUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`;
            const createResponse = await fetch(createSheetUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${accessToken}`
                },
                body: JSON.stringify({
                    requests: [{
                        addSheet: {
                            properties: {
                                title: sheetName
                            }
                        }
                    }]
                })
            });
            
            if (createResponse.ok) {
                const createData = await createResponse.json();
                targetSheetId = createData.replies[0].addSheet.properties.sheetId;
            }
        }

        // Clear the sheet first for clean overwrite
        const clearUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${sheetName}!A:ZZ:clear`;
        await fetch(clearUrl, {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });
        
        // Write the values
        const range = `${sheetName}!A1`;
        const valuesUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}?valueInputOption=RAW`;
        
        const valuesResponse = await fetch(valuesUrl, {
            method: 'PUT',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${accessToken}`
            },
            body: JSON.stringify({
                values: data
            })
        });

        if (!valuesResponse.ok) {
            const errorData = await valuesResponse.json();
            throw new Error(errorData.error?.message || 'Failed to write values');
        }

        // Then apply formatting (colors)
        if (colors && colors.length > 0) {
            const requests = [];
            
            colors.forEach((rowColors, rowIndex) => {
                rowColors.forEach((color, colIndex) => {
                    if (color) {
                        const rgb = hexToRgb(color);
                        if (rgb) {
                            requests.push({
                                repeatCell: {
                                    range: {
                                        sheetId: targetSheetId,
                                        startRowIndex: rowIndex,
                                        endRowIndex: rowIndex + 1,
                                        startColumnIndex: colIndex,
                                        endColumnIndex: colIndex + 1
                                    },
                                    cell: {
                                        userEnteredFormat: {
                                            backgroundColor: rgb,
                                            textFormat: {
                                                foregroundColor: {red: 1, green: 1, blue: 1}, // White text
                                                bold: true
                                            }
                                        }
                                    },
                                    fields: 'userEnteredFormat(backgroundColor,textFormat)'
                                }
                            });
                        }
                    }
                });
            });

            if (requests.length > 0) {
                const batchUpdateUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`;
                
                // Process in batches of 100 (Google Sheets API limit)
                for (let i = 0; i < requests.length; i += 100) {
                    const batch = requests.slice(i, i + 100);
                    const formatResponse = await fetch(batchUpdateUrl, {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json',
                            'Authorization': `Bearer ${accessToken}`
                        },
                        body: JSON.stringify({
                            requests: batch
                        })
                    });

                    if (!formatResponse.ok) {
                        const errorData = await formatResponse.json();
                        console.warn('Failed to apply some formatting:', errorData);
                    }
                }
            }
        }

        return true;
    } catch (error) {
        console.error('Error writing to Google Sheet:', error);
        throw error;
    }
}

// Export draw results to Google Sheets (requires OAuth2)
async function exportToGoogleSheet() {
    if (!drawState.groups || Object.keys(drawState.groups).length === 0) {
        updateStatus('No draw results to export!');
        return;
    }

    // Check if user is signed in with OAuth2
    // Check both the Drive section and the export-specific field, then localStorage, then config
    let clientId = document.getElementById('googleDriveClientIdForExport')?.value.trim() ||
                   document.getElementById('googleDriveClientId')?.value.trim() ||
                   localStorage.getItem('googleDriveClientId') ||
                   (typeof APP_CONFIG !== 'undefined' ? APP_CONFIG.googleDrive?.clientId || '' : '');
    
    // Save to localStorage if found
    if (clientId && clientId.trim()) {
        localStorage.setItem('googleDriveClientId', clientId.trim());
    }
    
    // If no client ID, try to get it from config or prompt user
    if (!clientId) {
        const useOAuth = confirm('Export requires Google OAuth2. Do you have a Google OAuth Client ID? (Click OK to enter it, Cancel to skip)');
        if (useOAuth) {
            clientId = prompt('Enter your Google OAuth Client ID:');
            if (!clientId || !clientId.trim()) {
                updateStatus('Export cancelled - OAuth Client ID required');
                return;
            }
            // Save it to the input field if it exists
            const clientIdInput = document.getElementById('googleDriveClientId');
            if (clientIdInput) {
                clientIdInput.value = clientId.trim();
            }
        } else {
            updateStatus('Export requires Google OAuth2. Please set up Google Drive integration first.');
            return;
        }
    }

    // Ensure user is signed in
    if (!isSignedIn() || !accessToken) {
        try {
            updateStatus('Signing in to Google...');
            await initGoogleDriveAPI(clientId);
            await signInWithGoogle();
            if (!accessToken) {
                updateStatus('Failed to get access token. Please try again.');
                return;
            }
        } catch (error) {
            updateStatus(`Sign-in error: ${error.message}. Please check your Client ID.`);
            return;
        }
    }

    // Get source spreadsheet info to find parent folder
    const sourceUrl = document.getElementById('googleDriveFileUrl')?.value.trim() || 
                     document.getElementById('googleSheetUrl')?.value.trim() || '';
    const sourceId = sourceUrl ? extractSpreadsheetId(sourceUrl) : null;
    
    let parentFolderId = null;
    let spreadsheetId = null;
    const exportFileName = `${config.eventTitle} - Draw Results`;
    
    try {
        updateStatus('Finding source folder...');
        
        // Get parent folder of source spreadsheet
        if (sourceId) {
            const parentResponse = await fetch(
                `https://www.googleapis.com/drive/v3/files/${sourceId}?fields=parents`,
                { headers: { 'Authorization': `Bearer ${accessToken}` } }
            );
            if (parentResponse.ok) {
                const parentData = await parentResponse.json();
                parentFolderId = parentData.parents?.[0] || null;
                console.log('Source folder ID:', parentFolderId);
            }
        }
        
        // Search for existing export file in the folder
        if (parentFolderId) {
            const searchQuery = `name='${exportFileName}' and '${parentFolderId}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`;
            const searchResponse = await fetch(
                `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(searchQuery)}&fields=files(id,name)`,
                { headers: { 'Authorization': `Bearer ${accessToken}` } }
            );
            if (searchResponse.ok) {
                const searchData = await searchResponse.json();
                if (searchData.files && searchData.files.length > 0) {
                    spreadsheetId = searchData.files[0].id;
                    console.log('Found existing export file:', spreadsheetId);
                    updateStatus('Updating existing export file...');
                }
            }
        }
        
        // Create new spreadsheet if not found
        if (!spreadsheetId) {
            updateStatus('Creating new spreadsheet...');
            const createResponse = await fetch('https://sheets.googleapis.com/v4/spreadsheets', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${accessToken}`
                },
                body: JSON.stringify({
                    properties: { title: exportFileName }
                })
            });

            if (!createResponse.ok) {
                const errorData = await createResponse.json();
                throw new Error(errorData.error?.message || 'Failed to create spreadsheet');
            }

            const createData = await createResponse.json();
            spreadsheetId = createData.spreadsheetId;
            
            // Move to source folder if we have one
            if (parentFolderId) {
                await fetch(`https://www.googleapis.com/drive/v3/files/${spreadsheetId}?addParents=${parentFolderId}&removeParents=root`, {
                    method: 'PATCH',
                    headers: { 'Authorization': `Bearer ${accessToken}` }
                });
                console.log('Moved export file to source folder');
            }
            
            updateStatus(`Created: ${exportFileName}`);
        }
    } catch (error) {
        updateStatus(`Error setting up export: ${error.message}`);
        return;
    }

    try {
        updateStatus('Exporting results...');

        // Prepare data: groups as columns, entries as rows
        const sheetData = [];
        const colorData = [];

        // Header row: Group names with rooms
        const headerRow = [];
        const headerColors = [];
        config.groupNames.forEach(name => {
            const room = config.groupRooms?.[name] || '';
            const headerText = room ? `${name}\n(${room})` : name;
            headerRow.push(headerText);
            headerColors.push(null);
        });
        sheetData.push(headerRow);
        colorData.push(headerColors);

        // Find maximum number of entries in any group
        let maxEntries = 0;
        config.groupNames.forEach(name => {
            const entries = drawState.groups[name] || [];
            maxEntries = Math.max(maxEntries, entries.length);
        });

        // Data rows: one row per entry position (no row label column)
        for (let i = 0; i < maxEntries; i++) {
            const row = [];
            const rowColors = [];
            
            config.groupNames.forEach(groupName => {
                const entries = drawState.groups[groupName] || [];
                const entryData = entries[i];
                
                if (entryData) {
                    const entryName = typeof entryData === 'string' ? entryData : entryData.entry;
                    const potIndex = typeof entryData === 'string' ? -1 : entryData.potIndex;
                    const potColor = potIndex >= 0 ? POT_COLORS[potIndex % POT_COLORS.length] : null;
                    
                    row.push(entryName || '');
                    rowColors.push(potColor);
                } else {
                    row.push('');
                    rowColors.push(null);
                }
            });
            
            sheetData.push(row);
            colorData.push(rowColors);
        }

        // Write to sheet using OAuth2
        await writeToGoogleSheetOAuth(accessToken, spreadsheetId, 'Draw Results', sheetData, colorData);
        
        const sheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}`;
        updateStatus(`✓ Exported! <a href="${sheetUrl}" target="_blank" style="color: #4CAF50; text-decoration: underline; font-weight: 600;">Open Sheet</a>`);
        
        // Also open in new tab automatically
        window.open(sheetUrl, '_blank');
    } catch (error) {
        updateStatus(`Export error: ${error.message}`);
        console.error('Export error:', error);
    }
}

// Parse Google Sheets data into pots format
function parseSheetDataToPots(sheetData) {
    if (!sheetData || sheetData.length === 0) {
        throw new Error('Sheet is empty');
    }

    console.log('Parsing sheet data:', sheetData);
    console.log('Sheet data length:', sheetData.length);
    console.log('First row (headers):', sheetData[0]);

    const headers = sheetData[0] || []; // First row contains pot names
    const pots = [];

    console.log('Headers from first row:', headers);

    // Create a pot for each column
    for (let colIndex = 0; colIndex < headers.length; colIndex++) {
        // Get pot name from first row (header)
        const potName = (headers[colIndex] || '').trim();
        console.log(`Column ${colIndex}: header = "${potName}"`);
        
        // Use the header as pot name - if empty, use default name
        const finalPotName = potName || `Pot ${colIndex + 1}`;

        const entries = [];

        // Collect entries from rows below the header (starting from row 2, index 1)
        // Only count up to row 10 (index 1-10, so rows 2-11)
        const maxRowToLoad = Math.min(11, sheetData.length); // Check up to row 11 (index 10)
        for (let rowIndex = 1; rowIndex < maxRowToLoad; rowIndex++) {
            const row = sheetData[rowIndex];
            if (!row || !Array.isArray(row)) continue;
            const entry = (row[colIndex] || '').trim();
            if (entry) {
                entries.push(entry);
            }
        }

        console.log(`Pot "${finalPotName}" has ${entries.length} entries:`, entries);

        pots.push({
            name: finalPotName, // Pot name comes from first row of sheet
            entries: entries
        });
    }

    console.log('Parsed pots:', pots);
    return pots;
}

// Load pots from Google Sheets
async function loadPotsFromGoogleSheets(apiKey, sheetUrl, sheetName) {
    const spreadsheetId = extractSpreadsheetId(sheetUrl);
    if (!spreadsheetId) {
        throw new Error('Invalid Google Sheets URL');
    }

    const sheetData = await fetchGoogleSheetData(apiKey, spreadsheetId, sheetName);
    const pots = parseSheetDataToPots(sheetData);

    if (pots.length === 0) {
        throw new Error('No pots found in the sheet');
    }

    return pots;
}

// Load constraints from CannotBeWith sheet
async function loadCannotBeWithConstraints(apiKey, sheetUrl, accessToken = null) {
    try {
        let spreadsheetId = extractSpreadsheetId(sheetUrl);
        
        // If not found, try extracting from Drive URL format
        if (!spreadsheetId) {
            const fileId = extractFileIdFromDriveUrl(sheetUrl);
            if (fileId) {
                spreadsheetId = fileId;
            }
        }
        
        if (!spreadsheetId) {
            console.warn('Invalid Google Sheets URL, skipping CannotBeWith sheet');
            return null;
        }

        console.log('Loading CannotBeWith sheet from spreadsheet:', spreadsheetId);
        let sheetData;
        
        // Use OAuth if access token provided, otherwise use API key
        if (accessToken) {
            const range = 'CannotBeWith!A1:B1000';
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}`;
            console.log('Fetching CannotBeWith with OAuth from:', url);
            const response = await fetch(url, {
                headers: {
                    'Authorization': `Bearer ${accessToken}`
                }
            });
            
            if (!response.ok) {
                const errorData = await response.json();
                console.error('Error fetching CannotBeWith sheet with OAuth:', errorData);
                return null;
            }
            
            const data = await response.json();
            sheetData = data.values || [];
            console.log('CannotBeWith sheet data received:', sheetData.length, 'rows');
        } else {
            // Use API key
            if (!apiKey) {
                console.warn('No API key provided for CannotBeWith sheet');
                return null;
            }
            console.log('Fetching CannotBeWith with API key');
            sheetData = await fetchGoogleSheetData(apiKey, spreadsheetId, 'CannotBeWith');
            console.log('CannotBeWith sheet data received:', sheetData.length, 'rows');
        }
        
        if (!sheetData || sheetData.length === 0) {
            console.warn('CannotBeWith sheet is empty or not found');
            return null;
        }

        // Extract pairs from columns A and B (each row is a pair)
        // Skip first row if it looks like a header (contains common header words)
        const pairs = [];
        const headerKeywords = ['entry', 'name', 'participant', 'person', 'team', 'cannot', 'must'];
        let startRow = 0;
        
        // Check if first row looks like a header
        if (sheetData.length > 0 && sheetData[0] && sheetData[0][0]) {
            const firstCell = sheetData[0][0].toLowerCase().trim();
            if (headerKeywords.some(keyword => firstCell.includes(keyword))) {
                startRow = 1; // Skip header row
            }
        }
        
        for (let i = startRow; i < sheetData.length; i++) {
            const row = sheetData[i];
            if (row && row[0] && row[1]) {
                const entry1 = row[0].trim();
                const entry2 = row[1].trim();
                if (entry1 && entry2 && entry1.length > 0 && entry2.length > 0) {
                    pairs.push([entry1, entry2]);
                }
            }
        }

        if (pairs.length === 0) {
            console.warn('No valid pairs found in CannotBeWith sheet');
            return null;
        }

        console.log('Loaded CannotBeWith constraints:', pairs);
        return pairs;
    } catch (error) {
        console.warn('Error loading CannotBeWith sheet:', error);
        return null;
    }
}

// Load constraints from MustBeWith sheet
async function loadMustBeWithConstraints(apiKey, sheetUrl, accessToken = null) {
    try {
        let spreadsheetId = extractSpreadsheetId(sheetUrl);
        
        // If not found, try extracting from Drive URL format
        if (!spreadsheetId) {
            const fileId = extractFileIdFromDriveUrl(sheetUrl);
            if (fileId) {
                spreadsheetId = fileId;
            }
        }
        
        if (!spreadsheetId) {
            console.warn('Invalid Google Sheets URL, skipping MustBeWith sheet');
            return null;
        }

        console.log('Loading MustBeWith sheet from spreadsheet:', spreadsheetId);
        let sheetData;
        
        // Use OAuth if access token provided, otherwise use API key
        if (accessToken) {
            const range = 'MustBeWith!A1:B1000';
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}`;
            console.log('Fetching MustBeWith with OAuth from:', url);
            const response = await fetch(url, {
                headers: {
                    'Authorization': `Bearer ${accessToken}`
                }
            });
            
            if (!response.ok) {
                const errorData = await response.json();
                console.error('Error fetching MustBeWith sheet with OAuth:', errorData);
                return null;
            }
            
            const data = await response.json();
            sheetData = data.values || [];
            console.log('MustBeWith sheet data received:', sheetData.length, 'rows');
        } else {
            // Use API key
            if (!apiKey) {
                console.warn('No API key provided for MustBeWith sheet');
                return null;
            }
            console.log('Fetching MustBeWith with API key');
            sheetData = await fetchGoogleSheetData(apiKey, spreadsheetId, 'MustBeWith');
            console.log('MustBeWith sheet data received:', sheetData.length, 'rows');
        }
        
        if (!sheetData || sheetData.length === 0) {
            console.warn('MustBeWith sheet is empty or not found');
            return null;
        }

        // Extract pairs from columns A and B (each row is a pair)
        // Skip first row if it looks like a header (contains common header words)
        const pairs = [];
        const headerKeywords = ['entry', 'name', 'participant', 'person', 'team', 'cannot', 'must'];
        let startRow = 0;
        
        // Check if first row looks like a header
        if (sheetData.length > 0 && sheetData[0] && sheetData[0][0]) {
            const firstCell = sheetData[0][0].toLowerCase().trim();
            if (headerKeywords.some(keyword => firstCell.includes(keyword))) {
                startRow = 1; // Skip header row
            }
        }
        
        for (let i = startRow; i < sheetData.length; i++) {
            const row = sheetData[i];
            if (row && row[0] && row[1]) {
                const entry1 = row[0].trim();
                const entry2 = row[1].trim();
                if (entry1 && entry2 && entry1.length > 0 && entry2.length > 0) {
                    pairs.push([entry1, entry2]);
                }
            }
        }

        if (pairs.length === 0) {
            console.warn('No valid pairs found in MustBeWith sheet');
            return null;
        }

        console.log('Loaded MustBeWith constraints:', pairs);
        return pairs;
    } catch (error) {
        console.warn('Error loading MustBeWith sheet:', error);
        return null;
    }
}

// Load constraints from MustBeInTopic sheet
async function loadMustBeInTopicConstraints(apiKey, sheetUrl, accessToken = null) {
    try {
        let spreadsheetId = extractSpreadsheetId(sheetUrl);
        
        // If not found, try extracting from Drive URL format
        if (!spreadsheetId) {
            const fileId = extractFileIdFromDriveUrl(sheetUrl);
            if (fileId) {
                spreadsheetId = fileId;
            }
        }
        
        if (!spreadsheetId) {
            console.warn('Invalid Google Sheets URL, skipping MustBeInTopic sheet');
            return null;
        }

        console.log('Loading MustBeInTopic sheet from spreadsheet:', spreadsheetId);
        let sheetData;
        
        // Use OAuth if access token provided, otherwise use API key
        if (accessToken) {
            const range = 'MustBeInTopic!A1:B1000';
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}`;
            console.log('Fetching MustBeInTopic with OAuth from:', url);
            const response = await fetch(url, {
                headers: {
                    'Authorization': `Bearer ${accessToken}`
                }
            });
            
            if (!response.ok) {
                const errorData = await response.json();
                console.error('Error fetching MustBeInTopic sheet with OAuth:', errorData);
                return null;
            }
            
            const data = await response.json();
            sheetData = data.values || [];
            console.log('MustBeInTopic sheet data received:', sheetData.length, 'rows');
        } else {
            // Use API key
            if (!apiKey) {
                console.warn('No API key provided for MustBeInTopic sheet');
                return null;
            }
            console.log('Fetching MustBeInTopic with API key');
            sheetData = await fetchGoogleSheetData(apiKey, spreadsheetId, 'MustBeInTopic');
            console.log('MustBeInTopic sheet data received:', sheetData.length, 'rows');
        }
        
        if (!sheetData || sheetData.length === 0) {
            console.warn('MustBeInTopic sheet is empty or not found');
            return null;
        }

        // Extract mappings: Column A = Entry Name, Column B = Topic/Group Name
        // Skip first row if it looks like a header
        const mappings = {};
        const headerKeywords = ['entry', 'name', 'participant', 'person', 'team', 'topic', 'group'];
        let startRow = 0;
        
        // Check if first row looks like a header
        if (sheetData.length > 0 && sheetData[0] && sheetData[0][0]) {
            const firstCell = sheetData[0][0].toLowerCase().trim();
            if (headerKeywords.some(keyword => firstCell.includes(keyword))) {
                startRow = 1; // Skip header row
            }
        }
        
        for (let i = startRow; i < sheetData.length; i++) {
            const row = sheetData[i];
            if (row && row[0] && row[1]) {
                const entryName = row[0].trim();
                const topicName = row[1].trim();
                if (entryName && topicName && entryName.length > 0 && topicName.length > 0) {
                    mappings[entryName] = topicName;
                }
            }
        }

        if (Object.keys(mappings).length === 0) {
            console.warn('No valid mappings found in MustBeInTopic sheet');
            return null;
        }

        console.log('Loaded MustBeInTopic constraints:', mappings);
        return mappings;
    } catch (error) {
        console.warn('Error loading MustBeInTopic sheet:', error);
        return null;
    }
}

// Load rooms from Rooms sheet
async function loadRoomsFromSheet(apiKey, sheetUrl, accessToken = null) {
    try {
        let spreadsheetId = extractSpreadsheetId(sheetUrl);
        
        // If not found, try extracting from Drive URL format
        if (!spreadsheetId) {
            const fileId = extractFileIdFromDriveUrl(sheetUrl);
            if (fileId) {
                spreadsheetId = fileId;
            }
        }
        
        if (!spreadsheetId) {
            console.warn('Invalid Google Sheets URL, skipping Rooms sheet');
            return null;
        }

        let sheetData;
        
        // Use OAuth if access token provided, otherwise use API key
        if (accessToken) {
            const range = 'Rooms!A1:A1000';
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}`;
            const response = await fetch(url, {
                headers: {
                    'Authorization': `Bearer ${accessToken}`
                }
            });
            
            if (!response.ok) {
                const errorData = await response.json();
                console.warn('Error fetching Rooms sheet with OAuth:', errorData);
                return null;
            }
            
            const data = await response.json();
            sheetData = data.values || [];
        } else {
            // Use API key
            if (!apiKey) {
                console.warn('No API key provided for Rooms sheet');
                return null;
            }
            sheetData = await fetchGoogleSheetData(apiKey, spreadsheetId, 'Rooms');
        }
        
        if (!sheetData || sheetData.length === 0) {
            console.warn('Rooms sheet is empty or not found');
            return null;
        }

        // Extract rooms from column A (single column)
        // Skip first row if it looks like a header
        const headerKeywords = ['room', 'name'];
        let startRow = 0;
        
        if (sheetData.length > 0 && sheetData[0] && sheetData[0][0]) {
            const firstCell = sheetData[0][0].toLowerCase().trim();
            if (headerKeywords.some(keyword => firstCell.includes(keyword))) {
                startRow = 1;
                console.log('Header row detected in Rooms sheet, skipping first row');
            }
        }
        
        const rooms = [];
        for (let i = startRow; i < sheetData.length; i++) {
            const row = sheetData[i];
            if (row && row[0]) {
                const room = row[0].trim();
                if (room) {
                    rooms.push(room);
                }
            }
        }

        if (rooms.length === 0) {
            console.warn('No valid rooms found in Rooms sheet');
            return null;
        }

        console.log('Loaded rooms from Rooms sheet:', rooms);
        return rooms; // Return as array, will be mapped to groups by index
    } catch (error) {
        console.warn('Error loading Rooms sheet:', error);
        return null;
    }
}

// Load group names from Topics sheet
async function loadGroupNamesFromTopicsSheet(apiKey, sheetUrl, accessToken = null) {
    try {
        let spreadsheetId = extractSpreadsheetId(sheetUrl);
        
        // If not found, try extracting from Drive URL format
        if (!spreadsheetId) {
            const fileId = extractFileIdFromDriveUrl(sheetUrl);
            if (fileId) {
                // If it's a Google Sheets file, the file ID is the spreadsheet ID
                spreadsheetId = fileId;
            }
        }
        
        if (!spreadsheetId) {
            console.warn('Invalid Google Sheets URL, skipping Topics sheet');
            return null;
        }

        let topicsData;
        
        // Use OAuth if access token provided, otherwise use API key
        if (accessToken) {
            const range = 'Topics!A1:A1000';
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}`;
            const response = await fetch(url, {
                headers: {
                    'Authorization': `Bearer ${accessToken}`
                }
            });
            
            if (!response.ok) {
                const errorData = await response.json();
                console.warn('Error fetching Topics sheet with OAuth:', errorData);
                return null;
            }
            
            const data = await response.json();
            topicsData = data.values || [];
        } else {
            // Use API key
            topicsData = await fetchGoogleSheetData(apiKey, spreadsheetId, 'Topics');
        }
        
        if (!topicsData || topicsData.length === 0) {
            console.warn('Topics sheet is empty or not found');
            return null;
        }

        // Extract group names from column A (first column)
        // Assume each row contains a group name, starting from row 1
        const groupNames = [];
        for (let i = 0; i < topicsData.length; i++) {
            const row = topicsData[i];
            if (row && row[0]) {
                const name = row[0].trim();
                if (name) {
                    groupNames.push(name);
                }
            }
        }

        if (groupNames.length === 0) {
            console.warn('No group names found in Topics sheet');
            return null;
        }

        console.log('Loaded group names from Topics sheet:', groupNames);
        return groupNames;
    } catch (error) {
        console.warn('Error loading Topics sheet:', error);
        // Don't throw - just return null so we can fall back to default names
        return null;
    }
}

// Detect number of pots and groups from a Google Sheet
async function detectSheetStructure(apiKey, sheetUrl, sheetName) {
    try {
        const spreadsheetId = extractSpreadsheetId(sheetUrl);
        if (!spreadsheetId) {
            return null;
        }

        const sheetData = await fetchGoogleSheetData(apiKey, spreadsheetId, sheetName);
        if (!sheetData || sheetData.length === 0) {
            return null;
        }

        const headers = sheetData[0] || [];
        // Count columns with non-empty headers in first row (these are the pots)
        const numPots = headers.filter(h => h && h.trim()).length;

        // Count entries in each column up to row 10 (index 1-10, so rows 2-11) and find minimum
        const entryCounts = [];
        const maxRowToCheck = Math.min(11, sheetData.length); // Check up to row 11 (index 10)
        
        for (let colIndex = 0; colIndex < headers.length; colIndex++) {
            if (!headers[colIndex] || !headers[colIndex].trim()) continue;
            
            let count = 0;
            // Start from row 2 (index 1), check up to row 11 (index 10)
            for (let rowIndex = 1; rowIndex < maxRowToCheck; rowIndex++) {
                const row = sheetData[rowIndex];
                if (row && row[colIndex] && row[colIndex].trim()) {
                    count++;
                }
            }
            if (count > 0) {
                entryCounts.push(count);
            }
        }

        const numGroups = entryCounts.length > 0 ? Math.min(...entryCounts) : 0;

        return {
            numPots: numPots,
            numGroups: numGroups,
            entryCounts: entryCounts // For debugging
        };
    } catch (error) {
        console.error('Error detecting sheet structure:', error);
        return null;
    }
}

// ==================== END GOOGLE SHEETS INTEGRATION ====================

// ==================== GOOGLE DRIVE API INTEGRATION ====================

// Google OAuth 2.0 Configuration
const GOOGLE_DRIVE_CONFIG = {
    clientId: '', // Set this to your OAuth Client ID from Google Cloud Console
    scopes: [
        'https://www.googleapis.com/auth/drive.readonly', // Read access to Drive files (to get folder info)
        'https://www.googleapis.com/auth/drive.file', // Create and manage files/folders in Drive
        'https://www.googleapis.com/auth/spreadsheets.readonly', // Read-only access to Sheets
        'https://www.googleapis.com/auth/forms.body', // Create and edit Google Forms
        'https://www.googleapis.com/auth/contacts.readonly' // Read contacts to extract emails from contact tags
    ],
    discoveryDocs: ['https://www.googleapis.com/discovery/v1/apis/drive/v3/rest']
};

// OAuth 2.0 Token Management
let accessToken = null;
let tokenClient = null;

// Initialize Google API Client (using new Google Identity Services)
function initGoogleDriveAPI(clientId) {
    if (!clientId) {
        console.warn('Google Drive Client ID not provided');
        return Promise.reject(new Error('Client ID not provided'));
    }

    GOOGLE_DRIVE_CONFIG.clientId = clientId;

    // Load Google Identity Services (new library)
    return new Promise((resolve, reject) => {
        if (window.google && window.google.accounts) {
            // Already loaded
            resolve();
        } else {
            // Load the Google Identity Services script
            const script = document.createElement('script');
            script.src = 'https://accounts.google.com/gsi/client';
            script.async = true;
            script.defer = true;
            script.onload = () => {
                // Wait a bit for it to initialize
                setTimeout(() => {
                    if (window.google && window.google.accounts) {
                        resolve();
                    } else {
                        reject(new Error('Google Identity Services failed to initialize'));
                    }
                }, 500);
            };
            script.onerror = () => reject(new Error('Failed to load Google Identity Services script'));
            document.head.appendChild(script);
        }
    });
}

// Initialize GAPI Client (no longer needed for auth, but kept for API calls)
async function initializeGapiClient() {
    // This is now only used for loading the API client for making API calls
    // Authentication is handled separately via Google Identity Services
    if (typeof gapi === 'undefined') {
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://apis.google.com/js/api.js';
            script.onload = () => {
                gapi.load('client', () => {
                    gapi.client.init({
                        discoveryDocs: GOOGLE_DRIVE_CONFIG.discoveryDocs
                    }).then(resolve).catch(reject);
                });
            };
            script.onerror = () => reject(new Error('Failed to load Google API script'));
            document.head.appendChild(script);
        });
    }
    return Promise.resolve();
}

// Sign in with Google (using new Google Identity Services)
// promptMode: 'consent' = full auth dialog, 'select_account' = account picker only, 'none' = silent (fails if needs interaction), '' = minimal
async function signInWithGoogle(promptMode = 'consent') {
    return new Promise((resolve, reject) => {
        if (!window.google || !window.google.accounts) {
            reject(new Error('Google Identity Services not loaded. Please initialize first.'));
            return;
        }

        try {
            // Determine prompt value
            let promptValue;
            if (promptMode === true) promptValue = 'consent'; // Legacy: true = consent
            else if (promptMode === false) promptValue = ''; // Legacy: false = silent
            else promptValue = promptMode; // Use string directly
            
            const tokenClient = google.accounts.oauth2.initTokenClient({
                client_id: GOOGLE_DRIVE_CONFIG.clientId,
                scope: GOOGLE_DRIVE_CONFIG.scopes.join(' '),
                callback: (response) => {
                    if (response.error) {
                        reject(new Error(response.error + (response.error_description ? ': ' + response.error_description : '')));
                    } else {
                        accessToken = response.access_token;
                        resolve(accessToken);
                    }
                },
            });

            // Request access token with specified prompt mode
            tokenClient.requestAccessToken({ prompt: promptValue });
        } catch (error) {
            console.error('Error signing in:', error);
            reject(error);
        }
    });
}

// Get current access token (for use by other scripts)
function getAccessToken() {
    return accessToken;
}

// Make functions globally available
window.getAccessToken = getAccessToken;
window.isSignedIn = isSignedIn;
window.signInWithGoogle = signInWithGoogle;

// Sign out from Google
function signOutFromGoogle() {
    if (accessToken && window.google && window.google.accounts) {
        google.accounts.oauth2.revoke(accessToken, () => {
            console.log('Access token revoked');
        });
    }
    accessToken = null;
}

// Check if user is signed in
function isSignedIn() {
    return accessToken !== null;
}

// Extract File ID from various Google Drive URL formats
function extractFileIdFromDriveUrl(url) {
    if (!url) return null;

    // Format 1: https://drive.google.com/file/d/FILE_ID/view
    let match = url.match(/\/file\/d\/([a-zA-Z0-9-_]+)/);
    if (match) return match[1];

    // Format 2: https://drive.google.com/open?id=FILE_ID
    match = url.match(/[?&]id=([a-zA-Z0-9-_]+)/);
    if (match) return match[1];

    // Format 3: https://docs.google.com/spreadsheets/d/FILE_ID/edit (already handled by extractSpreadsheetId)
    match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (match) return match[1];

    // Format 4: Direct file ID
    if (/^[a-zA-Z0-9-_]+$/.test(url.trim())) {
        return url.trim();
    }

    return null;
}

// Get file metadata from Google Drive
async function getDriveFileMetadata(fileId) {
    if (!accessToken) {
        throw new Error('Not authenticated. Please sign in to Google.');
    }

    try {
        const response = await gapi.client.drive.files.get({
            fileId: fileId,
            fields: 'id,name,mimeType,size,modifiedTime'
        });

        return response.result;
    } catch (error) {
        console.error('Error getting file metadata:', error);
        throw new Error(`Failed to get file metadata: ${error.result?.error?.message || error.message}`);
    }
}

// Download file content from Google Drive
async function downloadDriveFile(fileId, mimeType) {
    if (!accessToken) {
        throw new Error('Not authenticated. Please sign in to Google.');
    }

    try {
        let url;
        
        // For Google Sheets, use the export API
        if (mimeType === 'application/vnd.google-apps.spreadsheet') {
            // Export as CSV
            url = `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=text/csv`;
        } else {
            // For other files, use the download API
            url = `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`;
        }

        const response = await fetch(url, {
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });

        if (!response.ok) {
            const errorData = await response.json().catch(() => ({}));
            throw new Error(errorData.error?.message || `Failed to download file: ${response.statusText}`);
        }

        return await response.text();
    } catch (error) {
        console.error('Error downloading file:', error);
        throw error;
    }
}

// Parse CSV content into array of arrays
function parseCSV(csvText) {
    const lines = csvText.split('\n').filter(line => line.trim());
    return lines.map(line => {
        // Simple CSV parser (handles basic cases)
        const result = [];
        let current = '';
        let inQuotes = false;

        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            if (char === '"') {
                inQuotes = !inQuotes;
            } else if (char === ',' && !inQuotes) {
                result.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        result.push(current.trim());
        return result;
    });
}

// Parse JSON content
function parseJSON(jsonText) {
    try {
        return JSON.parse(jsonText);
    } catch (error) {
        throw new Error('Invalid JSON format');
    }
}

// Load data from Google Drive file
async function loadDataFromGoogleDrive(fileUrl, fileType = 'auto') {
    const fileId = extractFileIdFromDriveUrl(fileUrl);
    if (!fileId) {
        throw new Error('Invalid Google Drive URL. Could not extract file ID.');
    }

    // Get file metadata to determine file type
    const metadata = await getDriveFileMetadata(fileId);
    const mimeType = metadata.mimeType;

    // Download file content
    const fileContent = await downloadDriveFile(fileId, mimeType);

    // Parse based on file type
    let parsedData;

    if (mimeType === 'application/vnd.google-apps.spreadsheet' || 
        mimeType === 'text/csv' || 
        metadata.name.endsWith('.csv')) {
        // CSV or Google Sheet exported as CSV
        parsedData = parseCSV(fileContent);
    } else if (mimeType === 'application/json' || metadata.name.endsWith('.json')) {
        // JSON file
        parsedData = parseJSON(fileContent);
    } else if (mimeType === 'text/plain' || metadata.name.endsWith('.txt')) {
        // Text file - split by lines
        parsedData = fileContent.split('\n').filter(line => line.trim()).map(line => [line.trim()]);
    } else {
        throw new Error(`Unsupported file type: ${mimeType}. Supported types: CSV, JSON, TXT, Google Sheets`);
    }

    return {
        data: parsedData,
        metadata: metadata
    };
}

// Load pots from Google Drive file (wrapper function)
async function loadPotsFromGoogleDrive(fileUrl, clientId) {
    // Initialize API if not already done
    if (!isSignedIn()) {
        if (!clientId) {
            throw new Error('Google Drive Client ID is required. Please provide it in the setup.');
        }
        
        await initGoogleDriveAPI(clientId);
        await signInWithGoogle();
    }

    // Load and parse the file
    const result = await loadDataFromGoogleDrive(fileUrl);
    
    // Parse into pots format (assuming same structure as Google Sheets)
    return parseSheetDataToPots(result.data);
}

// ==================== END GOOGLE DRIVE API INTEGRATION ====================

// ==================== CONFIGURATION STATE ====================
let config = {
    eventTitle: 'Tastewise Hacktivate',
    numGroups: 8,
    numPots: 4,
    groupNames: [],
    pots: [], // Array of { name: string, entries: string[] }
    groupRooms: {}, // Map of group name to room (e.g., { "Group A": "Room 101" })
    animationDuration: 0.8 // Duration in seconds between draws in animated mode
};

// ==================== POT COLORS ====================
// Colors chosen for good contrast with white/yellow text and distinct from each other
const POT_COLORS = [
    '#E74C3C', // Red - good contrast
    '#1ABC9C', // Turquoise - good contrast
    '#3498DB', // Blue - good contrast
    '#2ECC71', // Green - good contrast
    '#9B59B6', // Purple - good contrast
    '#E67E22', // Orange - good contrast
    '#F1C40F', // Yellow - good contrast (changed from orange/yellow)
    '#34495E'  // Dark Blue-Gray - good contrast (changed from dark red)
];

// Convert hex color to RGB for Google Sheets API
function hexToRgb(hex) {
    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? {
        red: parseInt(result[1], 16) / 255,
        green: parseInt(result[2], 16) / 255,
        blue: parseInt(result[3], 16) / 255
    } : null;
}

// ==================== DRAW STATE ====================
let drawState = {
    pots: [], // Deep copy of config.pots for draw
    groups: {}, // { groupName: [{entry: string, potIndex: number}, ...] }
    isDrawing: false,
    drawComplete: false,
    abortRequested: false
};

// ==================== DOM ELEMENTS ====================
const setupScreen = document.getElementById('setupScreen');
const drawScreen = document.getElementById('drawScreen');

// Step elements
const step1 = document.getElementById('step1');
const step2 = document.getElementById('step2');
const step3 = document.getElementById('step3');

// ==================== SETUP SCREEN LOGIC ====================

// Initialize default group names (A, B, C, ...)
function generateDefaultGroupNames(count) {
    const names = [];
    for (let i = 0; i < count; i++) {
        if (i < 26) {
            names.push(String.fromCharCode(65 + i)); // A, B, C, ...
        } else {
            names.push(`Group ${i + 1}`);
        }
    }
    return names;
}

// Generate default pot names
function generateDefaultPotNames(count) {
    const pots = [];
    for (let i = 0; i < count; i++) {
        pots.push({
            name: `Pot ${i + 1}`,
            entries: []
        });
    }
    return pots;
}

// Render Step 2: Group Names
function renderGroupsConfig() {
    const container = document.getElementById('groupsConfig');
    container.innerHTML = '';

    config.groupNames.forEach((name, index) => {
        const field = document.createElement('div');
        field.className = 'group-name-field';
        field.innerHTML = `
            <label>Group ${index + 1}:</label>
            <input type="text"
                   class="group-name-input"
                   data-index="${index}"
                   value="${name}"
                   placeholder="Enter name">
        `;
        container.appendChild(field);
    });

    // Add event listeners
    container.querySelectorAll('.group-name-input').forEach(input => {
        input.addEventListener('input', (e) => {
            const index = parseInt(e.target.dataset.index);
            config.groupNames[index] = e.target.value || `Group ${index + 1}`;
        });
    });
}

// Render Step 3: Pots Configuration
function renderPotsConfig() {
    const container = document.getElementById('potsConfig');
    container.innerHTML = '';

    // Debug: Check if pots have entries
    console.log('Rendering pots config:', config.pots);

    config.pots.forEach((pot, potIndex) => {
        // Ensure entries array exists
        if (!pot.entries) {
            pot.entries = [];
        }
        const potDiv = document.createElement('div');
        potDiv.className = 'pot-config';
        potDiv.dataset.potIndex = potIndex;

        potDiv.innerHTML = `
            <div class="pot-config-header">
                <h4>Pot ${potIndex + 1}</h4>
                <input type="text"
                       class="pot-name-input"
                       data-pot-index="${potIndex}"
                       value="${pot.name}"
                       placeholder="Pot name">
            </div>
            <div class="pot-entries" data-pot-index="${potIndex}">
                ${pot.entries && pot.entries.length > 0 ? pot.entries.map((entry, entryIndex) => `
                    <div class="entry-tag">
                        <span>${entry || ''}</span>
                        <button class="remove-entry" data-pot-index="${potIndex}" data-entry-index="${entryIndex}">&times;</button>
                    </div>
                `).join('') : '<div style="color: #aaa; font-style: italic; padding: 10px;">No entries loaded</div>'}
            </div>
            <div class="add-entry-row">
                <input type="text"
                       class="add-entry-input"
                       data-pot-index="${potIndex}"
                       placeholder="Enter name (e.g., Brazil, Team Alpha)">
                <button class="add-entry-btn" data-pot-index="${potIndex}">Add</button>
            </div>
            <div class="entry-count ${pot.entries.length >= config.numGroups ? 'valid' : 'invalid'}" data-pot-index="${potIndex}">
                ${pot.entries.length} / ${config.numGroups} entries ${pot.entries.length > config.numGroups ? `(extra entries will be ignored)` : pot.entries.length === config.numGroups ? '(exact)' : '(need more)'}
            </div>
        `;

        container.appendChild(potDiv);
    });

    // Add event listeners for pot names
    container.querySelectorAll('.pot-name-input').forEach(input => {
        input.addEventListener('input', (e) => {
            const potIndex = parseInt(e.target.dataset.potIndex);
            config.pots[potIndex].name = e.target.value || `Pot ${potIndex + 1}`;
        });
    });

    // Add event listeners for remove buttons
    container.querySelectorAll('.remove-entry').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const potIndex = parseInt(e.target.dataset.potIndex);
            const entryIndex = parseInt(e.target.dataset.entryIndex);
            config.pots[potIndex].entries.splice(entryIndex, 1);
            renderPotsConfig();
        });
    });

    // Add event listeners for add buttons
    container.querySelectorAll('.add-entry-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const potIndex = parseInt(e.target.dataset.potIndex);
            const input = container.querySelector(`.add-entry-input[data-pot-index="${potIndex}"]`);
            const value = input.value.trim();

            if (value) {
                config.pots[potIndex].entries.push(value);
                input.value = '';
                renderPotsConfig();
            }
        });
    });

    // Add enter key support for inputs
    container.querySelectorAll('.add-entry-input').forEach(input => {
        input.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                const potIndex = parseInt(e.target.dataset.potIndex);
                const btn = container.querySelector(`.add-entry-btn[data-pot-index="${potIndex}"]`);
                btn.click();
            }
        });
    });
}

// Validate configuration
function validateConfig() {
    const validationMsg = document.getElementById('validationMessage');
    const validationText = document.getElementById('validationText');

    // DO NOT truncate entries - keep ALL entries from ALL pots
    // We'll distribute all entries across groups, allowing imbalance
    // Some groups may have more entries than others if total doesn't divide evenly

    // Check group names are unique
    const uniqueNames = new Set(config.groupNames);
    if (uniqueNames.size !== config.groupNames.length) {
        validationMsg.classList.remove('hidden');
        validationText.textContent = 'Group names must be unique';
        return false;
    }

    validationMsg.classList.add('hidden');
    return true;
}

// Show validation error
function showValidation(message) {
    const validationMsg = document.getElementById('validationMessage');
    const validationText = document.getElementById('validationText');
    validationMsg.classList.remove('hidden');
    validationText.textContent = message;
}

function hideValidation() {
    document.getElementById('validationMessage').classList.add('hidden');
}

// ==================== NAVIGATION ====================

// Handle data source selection (Google Drive vs Google Sheets)
document.querySelectorAll('input[name="dataSource"]').forEach(radio => {
    radio.addEventListener('change', (e) => {
        const driveSection = document.getElementById('driveApiSection');
        const sheetsSection = document.getElementById('sheetsApiSection');
        
        if (e.target.value === 'drive') {
            driveSection.style.display = 'block';
            sheetsSection.style.display = 'none';
        } else {
            driveSection.style.display = 'none';
            sheetsSection.style.display = 'block';
        }
    });
});

// Handle Google Drive connection button
document.getElementById('connectGoogleDriveBtn').addEventListener('click', async () => {
    const clientIdInput = document.getElementById('googleDriveClientId');
    const clientId = clientIdInput.value.trim() || 
                     (typeof APP_CONFIG !== 'undefined' ? APP_CONFIG.googleDrive?.clientId || '' : '');
    const statusDiv = document.getElementById('googleDriveStatus');
    
    if (!clientId) {
        statusDiv.textContent = 'Please enter your Google OAuth Client ID first (or add it to config.js).';
        statusDiv.style.color = '#ff6b6b';
        return;
    }

    try {
        statusDiv.textContent = 'Loading Google API...';
        statusDiv.style.color = '#ffd700';
        
        // Initialize Google API
        await initGoogleDriveAPI(clientId);
        
        statusDiv.textContent = 'Signing in...';
        await signInWithGoogle();
        
        statusDiv.textContent = '✓ Connected to Google Drive!';
        statusDiv.style.color = '#4CAF50';
    } catch (error) {
        statusDiv.textContent = `Error: ${error.message}`;
        statusDiv.style.color = '#ff6b6b';
        console.error('Google Drive connection error:', error);
    }
});

// Step 1 -> Step 2
document.getElementById('continueToStep2').addEventListener('click', async () => {
    const numGroups = parseInt(document.getElementById('numGroups').value);
    const numPots = parseInt(document.getElementById('numPots').value);
    const eventTitle = document.getElementById('eventTitle').value.trim();
    const animationDuration = parseFloat(document.getElementById('animationDuration').value) || 0.8;
    const dataSource = document.querySelector('input[name="dataSource"]:checked')?.value || 'sheets';
    
    // Google Sheets API fields - check input field FIRST, then fallbacks
    const sheetUrlInput = document.getElementById('googleSheetUrl');
    // Try multiple ways to get the value
    const inputUrlValue = sheetUrlInput?.value || sheetUrlInput?.getAttribute('value') || '';
    const localStorageUrl = localStorage.getItem('lastGoogleSheetUrl') || '';
    const configUrl = (typeof APP_CONFIG !== 'undefined' ? APP_CONFIG.googleSheets?.lastSheetUrl || '' : '');
    
    // Use whichever URL is available (input field takes priority)
    let googleSheetUrl = inputUrlValue.trim() || localStorageUrl.trim() || configUrl.trim();
    
    // If we found a URL in the input but it's not in localStorage, save it
    if (inputUrlValue.trim() && inputUrlValue.trim() !== localStorageUrl) {
        localStorage.setItem('lastGoogleSheetUrl', inputUrlValue.trim());
        googleSheetUrl = inputUrlValue.trim();
    }
    
    const apiKeyInput = document.getElementById('googleApiKey');
    const inputApiKey = apiKeyInput?.value || '';
    const configApiKey = (typeof APP_CONFIG !== 'undefined' ? APP_CONFIG.googleSheets?.apiKey || '' : '');
    const googleApiKey = inputApiKey.trim() || configApiKey.trim();
    
    const sheetNameInput = document.getElementById('sheetName');
    const inputSheetName = sheetNameInput?.value || '';
    const configSheetName = (typeof APP_CONFIG !== 'undefined' ? APP_CONFIG.googleSheets?.defaultSheetName || 'Participants' : 'Participants');
    const sheetName = inputSheetName.trim() || configSheetName;
    
    console.log('=== STEP 1 -> STEP 2 DEBUG ===');
    console.log('Data source:', dataSource);
    console.log('Input field element exists:', !!sheetUrlInput);
    console.log('Input field value (raw):', sheetUrlInput?.value);
    console.log('Input field value (trimmed):', inputUrlValue.trim());
    console.log('Input field disabled?', sheetUrlInput?.disabled);
    console.log('Input field readonly?', sheetUrlInput?.readOnly);
    console.log('localStorage URL:', localStorageUrl);
    console.log('config URL:', configUrl);
    console.log('Google Sheet URL FINAL:', googleSheetUrl);
    console.log('API Key present:', !!googleApiKey);
    console.log('Sheet name:', sheetName);
    
    // Google Drive API fields (use config as fallback)
    const googleDriveClientId = document.getElementById('googleDriveClientId').value.trim() || 
                                 (typeof APP_CONFIG !== 'undefined' ? APP_CONFIG.googleDrive?.clientId || '' : '');
    const googleDriveFileUrl = document.getElementById('googleDriveFileUrl').value.trim();

    if (numGroups < 1 || numGroups > 16) {
        showValidation('Number of groups must be between 1 and 16');
        return;
    }
    if (numPots < 1 || numPots > 8) {
        showValidation('Number of pots must be between 1 and 8');
        return;
    }

    hideValidation();

    config.eventTitle = eventTitle || 'Custom Draw';
    config.numGroups = numGroups;
    config.numPots = numPots;
    config.animationDuration = animationDuration;
    config.groupNames = generateDefaultGroupNames(numGroups);
    
    // Save animation duration to localStorage
    localStorage.setItem('animationDuration', animationDuration.toString());

    const btn = document.getElementById('continueToStep2');
    btn.disabled = true;

    // Check which data source is selected
    if (dataSource === 'drive' && googleDriveFileUrl && googleDriveClientId) {
        try {
            // Show loading message
            showValidation('Loading data from Google Drive...');
            
            // Load pots from Google Drive
            const loadedPots = await loadPotsFromGoogleDrive(googleDriveFileUrl, googleDriveClientId);

            // Validate loaded data
            if (loadedPots.length !== numPots) {
                showValidation(`File has ${loadedPots.length} pots, but you specified ${numPots} pots. Please adjust.`);
                btn.disabled = false;
                return;
            }

            // DO NOT truncate entries - keep ALL entries from ALL pots
            // We'll distribute all entries across groups, allowing imbalance

            config.pots = loadedPots;
            
            // Try to get API key from input or config for fetching sheets
            const topicsApiKey = googleApiKey || (typeof APP_CONFIG !== 'undefined' ? APP_CONFIG.googleSheets?.apiKey || '' : '');
            
            // Load group names from Topics sheet (if it's a Google Sheets file)
            try {
                const topicsGroupNames = await loadGroupNamesFromTopicsSheet(topicsApiKey, googleDriveFileUrl, accessToken);
                if (topicsGroupNames && topicsGroupNames.length > 0) {
                    // Use group names from Topics sheet
                    if (topicsGroupNames.length >= numGroups) {
                        config.groupNames = topicsGroupNames.slice(0, numGroups);
                    } else {
                        // Use Topics names first, then pad with default names
                        const defaultNames = generateDefaultGroupNames(numGroups);
                        config.groupNames = [...topicsGroupNames];
                        for (let i = topicsGroupNames.length; i < numGroups; i++) {
                            config.groupNames.push(defaultNames[i]);
                        }
                    }
                    console.log('Group names loaded from Topics sheet:', config.groupNames);
                } else {
                    // Fall back to default names if Topics sheet not found or empty
                    config.groupNames = generateDefaultGroupNames(numGroups);
                    console.log('Using default group names (Topics sheet not found or empty)');
                }
                
                // Load rooms from Rooms sheet and randomly assign to groups
                try {
                    const rooms = await loadRoomsFromSheet(topicsApiKey, googleDriveFileUrl, accessToken);
                    if (rooms && Array.isArray(rooms) && rooms.length > 0) {
                        // Randomly assign rooms to groups
                        config.groupRooms = {};
                        const shuffledRooms = [...rooms].sort(() => Math.random() - 0.5); // Shuffle rooms
                        config.groupNames.forEach((groupName, index) => {
                            if (index < shuffledRooms.length) {
                                config.groupRooms[groupName] = shuffledRooms[index];
                            }
                        });
                        console.log('Loaded and randomly assigned rooms from Rooms sheet:', config.groupRooms);
                    } else {
                        config.groupRooms = {};
                        console.log('Rooms sheet not found or empty');
                    }
                } catch (error) {
                    console.warn('Could not load Rooms sheet:', error);
                    config.groupRooms = {};
                }
            } catch (error) {
                console.warn('Could not load Topics sheet, using default group names:', error);
                config.groupNames = generateDefaultGroupNames(numGroups);
            }
            
            // Load constraints from CannotBeWith and MustBeWith sheets
            try {
                console.log('Loading constraints from CannotBeWith and MustBeWith sheets...');
                const cannotBeWithPairs = await loadCannotBeWithConstraints(topicsApiKey, googleDriveFileUrl, accessToken);
                if (cannotBeWithPairs && cannotBeWithPairs.length > 0) {
                    CHEAT_CONSTRAINTS.cannotBeWith = cannotBeWithPairs;
                    console.log('✓ Loaded CannotBeWith constraints from sheet:', cannotBeWithPairs);
                } else {
                    // Clear constraints if sheet not found (instead of keeping old values)
                    CHEAT_CONSTRAINTS.cannotBeWith = [];
                    console.log('⚠ CannotBeWith sheet not found or empty, constraints cleared');
                }
                
                const mustBeWithPairs = await loadMustBeWithConstraints(topicsApiKey, googleDriveFileUrl, accessToken);
                if (mustBeWithPairs && mustBeWithPairs.length > 0) {
                    CHEAT_CONSTRAINTS.mustBeWith = mustBeWithPairs;
                    console.log('✓ Loaded MustBeWith constraints from sheet:', mustBeWithPairs);
                } else {
                    // Clear constraints if sheet not found (instead of keeping old values)
                    CHEAT_CONSTRAINTS.mustBeWith = [];
                    console.log('⚠ MustBeWith sheet not found or empty, constraints cleared');
                }
                
                const mustBeInTopicMappings = await loadMustBeInTopicConstraints(topicsApiKey, googleDriveFileUrl, accessToken);
                if (mustBeInTopicMappings && Object.keys(mustBeInTopicMappings).length > 0) {
                    CHEAT_CONSTRAINTS.mustBeInGroup = mustBeInTopicMappings;
                    console.log('✓ Loaded MustBeInTopic constraints from sheet:', mustBeInTopicMappings);
                } else {
                    // Clear constraints if sheet not found
                    CHEAT_CONSTRAINTS.mustBeInGroup = {};
                    console.log('⚠ MustBeInTopic sheet not found or empty, constraints cleared');
                }
            } catch (error) {
                console.error('Error loading constraint sheets:', error);
                // Clear constraints on error to avoid using stale data
                CHEAT_CONSTRAINTS.cannotBeWith = [];
                CHEAT_CONSTRAINTS.mustBeWith = [];
                CHEAT_CONSTRAINTS.mustBeInGroup = {};
            }
            
            btn.disabled = false;
            hideValidation();

        } catch (error) {
            showValidation(`Google Drive Error: ${error.message}`);
            btn.disabled = false;
            return;
        }
    } else if (dataSource === 'sheets') {
        // LAST CHANCE: Check the input field one more time right now
        const inputField = document.getElementById('googleSheetUrl');
        const inputValue = inputField?.value?.trim() || '';
        
        // If input field has a value, save it to localStorage immediately
        if (inputValue) {
            localStorage.setItem('lastGoogleSheetUrl', inputValue);
            console.log('Found URL in input field, saved to localStorage:', inputValue);
        }
        
        // Get URL from localStorage (most reliable)
        const localStorageUrl = localStorage.getItem('lastGoogleSheetUrl') || '';
        const actualSheetUrl = localStorageUrl || inputValue || googleSheetUrl || '';
        
        console.log('=== FINAL URL CHECK ===');
        console.log('Input field value RIGHT NOW:', inputValue);
        console.log('localStorage URL:', localStorageUrl);
        console.log('Using URL:', actualSheetUrl);
        
        if (!actualSheetUrl) {
            showValidation('Please enter a Google Sheets URL in the input field');
            btn.disabled = false;
            return;
        }
        // Check if we have API key (from input or config)
        const finalApiKey = googleApiKey || (typeof APP_CONFIG !== 'undefined' ? APP_CONFIG.googleSheets?.apiKey || '' : '');
        
        if (!finalApiKey) {
            showValidation('Please enter your Google Sheets API key (or add it to config.js)');
            btn.disabled = false;
            return;
        }
        
        // Use the API key we found
        const apiKeyToUse = finalApiKey;
        
        // Make absolutely sure we have a URL
        const urlToUse = googleSheetUrl || localStorage.getItem('lastGoogleSheetUrl') || '';
        
        if (!urlToUse) {
            showValidation('Please enter a Google Sheets URL');
            btn.disabled = false;
            return;
        }
        
        console.log('=== LOADING FROM SHEET ===');
        console.log('URL being used:', urlToUse);
        console.log('API Key being used:', apiKeyToUse ? apiKeyToUse.substring(0, 10) + '...' : 'MISSING');
        console.log('Sheet name:', sheetName);
        
        try {
            // Show loading message
            showValidation('Loading data from Google Sheets...');

            // Load pots from Google Sheets - use the actual URL we found
            const loadedPots = await loadPotsFromGoogleSheets(apiKeyToUse, actualSheetUrl, sheetName);
            console.log('Loaded pots BEFORE processing:', loadedPots);
            console.log('Loaded pots entries:', loadedPots.map(p => ({ name: p.name, entryCount: p.entries?.length || 0, entries: p.entries })));

            // Validate loaded data
            if (loadedPots.length !== numPots) {
                showValidation(`Sheet has ${loadedPots.length} pots, but you specified ${numPots} pots. Please adjust.`);
                btn.disabled = false;
                return;
            }

            // DO NOT truncate entries - keep ALL entries from ALL pots
            // Ensure entries arrays exist, but don't limit them
            for (const pot of loadedPots) {
                if (!pot.entries || !Array.isArray(pot.entries)) {
                    pot.entries = [];
                }
                // Keep ALL entries - we'll distribute all of them across groups
            }

            // Deep copy to ensure data is preserved - keep the pot names from the sheet
            config.pots = loadedPots.map(pot => ({
                name: pot.name || `Pot ${loadedPots.indexOf(pot) + 1}`, // Use name from sheet, fallback to default
                entries: Array.isArray(pot.entries) ? [...pot.entries] : [] // Copy entries array
            }));
            
            console.log('Pots loaded from Google Sheets:', config.pots);
            console.log('Pot names from sheet:', config.pots.map(p => p.name));
            
            // Load group names from Topics sheet
            try {
                const topicsGroupNames = await loadGroupNamesFromTopicsSheet(apiKeyToUse, actualSheetUrl);
                if (topicsGroupNames && topicsGroupNames.length > 0) {
                    // Use group names from Topics sheet
                    // If we have more names than groups, take the first numGroups
                    // If we have fewer names, pad with default names
                    if (topicsGroupNames.length >= numGroups) {
                        config.groupNames = topicsGroupNames.slice(0, numGroups);
                    } else {
                        // Use Topics names first, then pad with default names
                        const defaultNames = generateDefaultGroupNames(numGroups);
                        config.groupNames = [...topicsGroupNames];
                        for (let i = topicsGroupNames.length; i < numGroups; i++) {
                            config.groupNames.push(defaultNames[i]);
                        }
                    }
                    console.log('Group names loaded from Topics sheet:', config.groupNames);
                } else {
                    // Fall back to default names if Topics sheet not found or empty
                    config.groupNames = generateDefaultGroupNames(numGroups);
                    console.log('Using default group names (Topics sheet not found or empty)');
                }
                
                // Load rooms from Rooms sheet and randomly assign to groups
                try {
                    const rooms = await loadRoomsFromSheet(apiKeyToUse, actualSheetUrl);
                    if (rooms && Array.isArray(rooms) && rooms.length > 0) {
                        // Randomly assign rooms to groups
                        config.groupRooms = {};
                        const shuffledRooms = [...rooms].sort(() => Math.random() - 0.5); // Shuffle rooms
                        config.groupNames.forEach((groupName, index) => {
                            if (index < shuffledRooms.length) {
                                config.groupRooms[groupName] = shuffledRooms[index];
                            }
                        });
                        console.log('Loaded and randomly assigned rooms from Rooms sheet:', config.groupRooms);
                    } else {
                        config.groupRooms = {};
                        console.log('Rooms sheet not found or empty');
                    }
                } catch (error) {
                    console.warn('Could not load Rooms sheet:', error);
                    config.groupRooms = {};
                }
            } catch (error) {
                console.warn('Could not load Topics sheet, using default group names:', error);
                config.groupNames = generateDefaultGroupNames(numGroups);
            }
            
            // Load constraints from CannotBeWith and MustBeWith sheets
            try {
                console.log('Loading constraints from CannotBeWith and MustBeWith sheets...');
                const cannotBeWithPairs = await loadCannotBeWithConstraints(apiKeyToUse, actualSheetUrl);
                if (cannotBeWithPairs && cannotBeWithPairs.length > 0) {
                    CHEAT_CONSTRAINTS.cannotBeWith = cannotBeWithPairs;
                    console.log('✓ Loaded CannotBeWith constraints from sheet:', cannotBeWithPairs);
                } else {
                    // Clear constraints if sheet not found (instead of keeping old values)
                    CHEAT_CONSTRAINTS.cannotBeWith = [];
                    console.log('⚠ CannotBeWith sheet not found or empty, constraints cleared');
                }
                
                const mustBeWithPairs = await loadMustBeWithConstraints(apiKeyToUse, actualSheetUrl);
                if (mustBeWithPairs && mustBeWithPairs.length > 0) {
                    CHEAT_CONSTRAINTS.mustBeWith = mustBeWithPairs;
                    console.log('✓ Loaded MustBeWith constraints from sheet:', mustBeWithPairs);
                    console.log('MustBeWith constraints active:', CHEAT_CONSTRAINTS.mustBeWith);
                } else {
                    // Clear constraints if sheet not found (instead of keeping old values)
                    CHEAT_CONSTRAINTS.mustBeWith = [];
                    console.log('⚠ MustBeWith sheet not found or empty, constraints cleared');
                }
                
                const mustBeInTopicMappings = await loadMustBeInTopicConstraints(apiKeyToUse, actualSheetUrl);
                if (mustBeInTopicMappings && Object.keys(mustBeInTopicMappings).length > 0) {
                    CHEAT_CONSTRAINTS.mustBeInGroup = mustBeInTopicMappings;
                    console.log('✓ Loaded MustBeInTopic constraints from sheet:', mustBeInTopicMappings);
                } else {
                    // Clear constraints if sheet not found
                    CHEAT_CONSTRAINTS.mustBeInGroup = {};
                    console.log('⚠ MustBeInTopic sheet not found or empty, constraints cleared');
                }
            } catch (error) {
                console.error('Error loading constraint sheets:', error);
                // Clear constraints on error to avoid using stale data
                CHEAT_CONSTRAINTS.cannotBeWith = [];
                CHEAT_CONSTRAINTS.mustBeWith = [];
                CHEAT_CONSTRAINTS.mustBeInGroup = {};
            }
            
            // Save the sheet URL to localStorage for next time
            if (googleSheetUrl) {
                localStorage.setItem('lastGoogleSheetUrl', googleSheetUrl);
            }
            
            btn.disabled = false;
            hideValidation();

        } catch (error) {
            showValidation(`Google Sheets Error: ${error.message}`);
            btn.disabled = false;
            return;
        }
    } else {
        // No Google integration - use default empty pots
        config.pots = generateDefaultPotNames(numPots);
        btn.disabled = false;
    }

    renderGroupsConfig();

    step1.classList.add('hidden');
    step2.classList.remove('hidden');
    
    // If pots were loaded from Google Sheets/Drive, show option to skip Step 3
    const potsLoadedFromGoogle = (dataSource === 'sheets' && googleSheetUrl && googleApiKey) || 
                                  (dataSource === 'drive' && googleDriveFileUrl && googleDriveClientId);
    if (potsLoadedFromGoogle && config.pots.length > 0) {
        // Check if all pots have valid entries
        const allPotsValid = config.pots.every(pot => pot.entries.length >= config.numGroups);
        if (allPotsValid) {
            // Show a message that they can skip Step 3
            const skipMessage = document.createElement('div');
            skipMessage.id = 'skipStep3Message';
            skipMessage.style.cssText = 'margin-top: 15px; padding: 15px; background: rgba(76, 175, 80, 0.2); border: 1px solid rgba(76, 175, 80, 0.5); border-radius: 10px; text-align: center;';
            skipMessage.innerHTML = `
                <p style="color: #4CAF50; margin-bottom: 10px;">✓ Pots loaded from Google Sheet successfully!</p>
                <p style="color: #aaa; font-size: 0.9rem; margin-bottom: 10px;">You can review/edit pots in Step 3, or skip directly to the draw.</p>
                <button id="skipToDrawBtn" class="setup-btn primary" style="margin-top: 10px;">Skip to Draw</button>
            `;
            const step2Container = document.getElementById('step2');
            const existingSkipMsg = document.getElementById('skipStep3Message');
            if (existingSkipMsg) existingSkipMsg.remove();
            step2Container.appendChild(skipMessage);
            
            // Add event listener for skip button
            document.getElementById('skipToDrawBtn').addEventListener('click', () => {
                if (validateConfig()) {
                    startDrawScreen();
                }
            });
        }
    }
});

// Step 2 -> Step 1
document.getElementById('backToStep1').addEventListener('click', () => {
    step2.classList.add('hidden');
    step1.classList.remove('hidden');
    hideValidation();
});

// Step 2 -> Step 3
document.getElementById('continueToStep3').addEventListener('click', () => {
    // Validate group names are not empty
    const emptyNames = config.groupNames.some(name => !name.trim());
    if (emptyNames) {
        showValidation('All group names must be filled in');
        return;
    }

    // Debug: Log pots before rendering
    console.log('Going to Step 3, current config.pots:', config.pots);
    console.log('Pot entries:', config.pots.map(p => ({ name: p.name, entryCount: p.entries?.length || 0 })));

    hideValidation();
    renderPotsConfig();

    step2.classList.add('hidden');
    step3.classList.remove('hidden');
    
    // Remove skip message if it exists
    const skipMsg = document.getElementById('skipStep3Message');
    if (skipMsg) skipMsg.remove();
});

// Step 3 -> Step 2
document.getElementById('backToStep2').addEventListener('click', () => {
    step3.classList.add('hidden');
    step2.classList.remove('hidden');
    hideValidation();
});

// Start Draw
document.getElementById('startDraw').addEventListener('click', () => {
    if (!validateConfig()) {
        return;
    }

    startDrawScreen();
});

// Instant Draw from Setup Screen (Step 1)
document.getElementById('instantDrawBtnSetup').addEventListener('click', async () => {
    // Get values from Step 1
    const eventTitle = document.getElementById('eventTitle').value.trim();
    const numGroups = parseInt(document.getElementById('numGroups').value) || 8;
    const numPots = parseInt(document.getElementById('numPots').value) || 4;
    const animationDuration = parseFloat(document.getElementById('animationDuration').value) || 0.8;
    
    const dataSource = document.querySelector('input[name="dataSource"]:checked')?.value || 'sheets';
    const googleSheetUrl = document.getElementById('googleSheetUrl')?.value.trim() || '';
    const googleApiKey = document.getElementById('googleApiKey')?.value.trim() || '';
    const sheetName = document.getElementById('sheetName')?.value.trim() || 'Participants';
    const googleDriveFileUrl = document.getElementById('googleDriveFileUrl')?.value.trim() || '';
    const googleDriveClientId = document.getElementById('googleDriveClientId')?.value.trim() || '';
    
    // Validate basic inputs
    if (numGroups < 1 || numGroups > 16) {
        showValidation('Number of groups must be between 1 and 16');
        return;
    }
    if (numPots < 1 || numPots > 8) {
        showValidation('Number of pots must be between 1 and 8');
        return;
    }
    
    hideValidation();
    
    // Set config
    config.eventTitle = eventTitle || 'Custom Draw';
    config.numGroups = numGroups;
    config.numPots = numPots;
    config.animationDuration = animationDuration;
    config.groupNames = generateDefaultGroupNames(numGroups);
    
    // Load data if using Google Sheets/Drive
    if (dataSource === 'drive' && googleDriveFileUrl && googleDriveClientId) {
        try {
            showValidation('Loading data from Google Drive...');
            const loadedPots = await loadPotsFromGoogleDrive(googleDriveFileUrl, googleDriveClientId);
            if (loadedPots.length !== numPots) {
                showValidation(`File has ${loadedPots.length} pots, but you specified ${numPots} pots. Please adjust.`);
                return;
            }
            config.pots = loadedPots;
            
            // Load group names, rooms, and constraints
            const topicsApiKey = googleApiKey || (typeof APP_CONFIG !== 'undefined' ? APP_CONFIG.googleSheets?.apiKey || '' : '');
            const topicsGroupNames = await loadGroupNamesFromTopicsSheet(topicsApiKey, googleDriveFileUrl, accessToken);
            if (topicsGroupNames && topicsGroupNames.length > 0) {
                if (topicsGroupNames.length >= numGroups) {
                    config.groupNames = topicsGroupNames.slice(0, numGroups);
                } else {
                    const defaultNames = generateDefaultGroupNames(numGroups);
                    config.groupNames = [...topicsGroupNames];
                    for (let i = topicsGroupNames.length; i < numGroups; i++) {
                        config.groupNames.push(defaultNames[i]);
                    }
                }
            }
            
            // Load rooms
            const rooms = await loadRoomsFromSheet(topicsApiKey, googleDriveFileUrl, accessToken);
            if (rooms && Array.isArray(rooms) && rooms.length > 0) {
                config.groupRooms = {};
                const shuffledRooms = [...rooms].sort(() => Math.random() - 0.5);
                config.groupNames.forEach((groupName, index) => {
                    if (index < shuffledRooms.length) {
                        config.groupRooms[groupName] = shuffledRooms[index];
                    }
                });
            }
            
            // Load constraints
            const cannotBeWithPairs = await loadCannotBeWithConstraints(topicsApiKey, googleDriveFileUrl, accessToken);
            if (cannotBeWithPairs && cannotBeWithPairs.length > 0) {
                CHEAT_CONSTRAINTS.cannotBeWith = cannotBeWithPairs;
            }
            const mustBeWithPairs = await loadMustBeWithConstraints(topicsApiKey, googleDriveFileUrl, accessToken);
            if (mustBeWithPairs && mustBeWithPairs.length > 0) {
                CHEAT_CONSTRAINTS.mustBeWith = mustBeWithPairs;
            }
            
            const mustBeInTopicMappings = await loadMustBeInTopicConstraints(topicsApiKey, googleDriveFileUrl, accessToken);
            if (mustBeInTopicMappings && Object.keys(mustBeInTopicMappings).length > 0) {
                CHEAT_CONSTRAINTS.mustBeInGroup = mustBeInTopicMappings;
            }
            
            hideValidation();
        } catch (error) {
            showValidation(`Google Drive Error: ${error.message}`);
            return;
        }
    } else if (dataSource === 'sheets' && googleSheetUrl && googleApiKey) {
        try {
            showValidation('Loading data from Google Sheets...');
            const loadedPots = await loadPotsFromGoogleSheets(googleApiKey, googleSheetUrl, sheetName);
            if (loadedPots.length !== numPots) {
                showValidation(`Sheet has ${loadedPots.length} pots, but you specified ${numPots} pots. Please adjust.`);
                return;
            }
            config.pots = loadedPots.map(pot => ({
                name: pot.name || `Pot ${loadedPots.indexOf(pot) + 1}`,
                entries: Array.isArray(pot.entries) ? [...pot.entries] : []
            }));
            
            // Load group names, rooms, and constraints
            const topicsGroupNames = await loadGroupNamesFromTopicsSheet(googleApiKey, googleSheetUrl);
            if (topicsGroupNames && topicsGroupNames.length > 0) {
                if (topicsGroupNames.length >= numGroups) {
                    config.groupNames = topicsGroupNames.slice(0, numGroups);
                } else {
                    const defaultNames = generateDefaultGroupNames(numGroups);
                    config.groupNames = [...topicsGroupNames];
                    for (let i = topicsGroupNames.length; i < numGroups; i++) {
                        config.groupNames.push(defaultNames[i]);
                    }
                }
            }
            
            // Load rooms
            const rooms = await loadRoomsFromSheet(googleApiKey, googleSheetUrl);
            if (rooms && Array.isArray(rooms) && rooms.length > 0) {
                config.groupRooms = {};
                const shuffledRooms = [...rooms].sort(() => Math.random() - 0.5);
                config.groupNames.forEach((groupName, index) => {
                    if (index < shuffledRooms.length) {
                        config.groupRooms[groupName] = shuffledRooms[index];
                    }
                });
            }
            
            // Load constraints
            const cannotBeWithPairs = await loadCannotBeWithConstraints(googleApiKey, googleSheetUrl);
            if (cannotBeWithPairs && cannotBeWithPairs.length > 0) {
                CHEAT_CONSTRAINTS.cannotBeWith = cannotBeWithPairs;
            }
            const mustBeWithPairs = await loadMustBeWithConstraints(googleApiKey, googleSheetUrl);
            if (mustBeWithPairs && mustBeWithPairs.length > 0) {
                CHEAT_CONSTRAINTS.mustBeWith = mustBeWithPairs;
            }
            
            const mustBeInTopicMappings = await loadMustBeInTopicConstraints(googleApiKey, googleSheetUrl);
            if (mustBeInTopicMappings && Object.keys(mustBeInTopicMappings).length > 0) {
                CHEAT_CONSTRAINTS.mustBeInGroup = mustBeInTopicMappings;
            }
            
            hideValidation();
        } catch (error) {
            showValidation(`Google Sheets Error: ${error.message}`);
            return;
        }
    } else {
        // No Google integration - use default empty pots
        config.pots = generateDefaultPotNames(numPots);
    }
    
    // Validate and start draw
    if (!validateConfig()) {
        return;
    }
    
    startDrawScreen();
    // Small delay to ensure draw screen is rendered, then instant draw
    setTimeout(() => {
        instantDrawAll();
    }, 100);
});

// ==================== DRAW SCREEN LOGIC ====================

function startDrawScreen() {
    // Initialize draw state
    drawState = {
        pots: JSON.parse(JSON.stringify(config.pots)), // Deep copy
        groups: {},
        isDrawing: false,
        drawComplete: false
    };

    // Initialize groups as empty arrays (can grow to accommodate all entries)
    config.groupNames.forEach(name => {
        drawState.groups[name] = [];
    });

    // Update title
    document.getElementById('drawTitle').textContent = config.eventTitle;

    // Render pots and groups
    renderDrawPots();
    renderDrawGroups();

    // Switch screens
    setupScreen.classList.add('hidden');
    drawScreen.classList.remove('hidden');

    updateStatus('Click "DRAW NEXT" to begin');
    document.getElementById('drawBtn').disabled = false;
    document.getElementById('autoDrawBtn').disabled = false;
    const instantDrawBtn = document.getElementById('instantDrawBtn');
    if (instantDrawBtn) instantDrawBtn.disabled = false;
    
    // Add voting button if function exists
    // Also check if draw is already complete (in case we're returning to this screen)
    if (typeof addVotingButtonToUI === 'function') {
        setTimeout(() => {
            addVotingButtonToUI();
            // If draw is already complete, make sure button is there
            if (drawState && drawState.drawComplete) {
                setTimeout(() => addVotingButtonToUI(), 1000);
            }
        }, 500);
    }
}

function renderDrawPots() {
    const container = document.getElementById('potsContainer');
    container.innerHTML = '';

    drawState.pots.forEach((pot, index) => {
        const potColor = POT_COLORS[index % POT_COLORS.length];
        const potDiv = document.createElement('div');
        potDiv.className = 'pot';
        potDiv.id = `pot${index}`;
        potDiv.style.setProperty('--pot-color', potColor);

        potDiv.innerHTML = `
            <div class="pot-header" style="background: ${potColor}; border-left: 4px solid ${potColor}; color: white; font-weight: 600; text-shadow: 1px 1px 2px rgba(0,0,0,0.5);">${pot.name}</div>
            <div class="pot-teams">
                ${pot.entries.map((entry, entryIndex) => `
                    <div class="team-ball" data-pot="${index}" data-entry="${entryIndex}" style="background: ${potColor}; border-color: ${potColor}; box-shadow: 0 4px 15px ${potColor}80;">
                        <span class="name" style="color: white; font-weight: 600; text-shadow: 1px 1px 2px rgba(0,0,0,0.5);">${entry}</span>
                    </div>
                `).join('')}
            </div>
        `;

        container.appendChild(potDiv);
    });
}

function renderDrawGroups() {
    const container = document.getElementById('groupsContainer');
    container.innerHTML = '';

    config.groupNames.forEach((name, index) => {
        const groupDiv = document.createElement('div');
        groupDiv.className = 'group';
        groupDiv.id = `group${index}`;

        let slotsHtml = '';
        // Show all entries in the group (can be more than numPots due to imbalance)
        const entries = drawState.groups[name] || [];
        entries.forEach((entryData, index) => {
            if (entryData) {
                // Handle both old format (string) and new format ({entry, potIndex})
                const entryName = typeof entryData === 'string' ? entryData : entryData.entry;
                const potIndex = typeof entryData === 'string' ? -1 : entryData.potIndex;
                const potColor = potIndex >= 0 ? POT_COLORS[potIndex % POT_COLORS.length] : '#666';
                
                slotsHtml += `
                    <div class="slot filled" data-position="${index}" style="background: ${potColor}; border: 2px solid ${potColor};">
                        <div class="team-info team-placed" style="background: transparent; border: none;">
                            <span class="entry-name" style="color: white; font-weight: 600; text-shadow: 1px 1px 2px rgba(0,0,0,0.5);">${entryName}</span>
                        </div>
                    </div>
                `;
            } else {
                slotsHtml += `<div class="slot" data-position="${index}"></div>`;
            }
        });

        // Get room for this group if available
        const room = config.groupRooms && config.groupRooms[name] ? config.groupRooms[name] : null;
        const headerText = room ? `${name} (${room})` : name;
        
        groupDiv.innerHTML = `
            <div class="group-header">${headerText}</div>
            <div class="group-slots">${slotsHtml}</div>
        `;

        container.appendChild(groupDiv);
    });
}

// Get available groups for a specific pot position
// Now returns all groups since we allow imbalance
function getAvailableGroups(potIndex) {
    // Return groups that don't already have an entry from this pot
    // This enforces the constraint: at most one entry per pot per group
    return config.groupNames.filter(groupName => {
        const groupEntries = drawState.groups[groupName] || [];
        // Check if any entry in this group came from the same pot
        const hasPotEntry = groupEntries.some(entry => entry.potIndex === potIndex);
        return !hasPotEntry;
    });
}

// ==================== CONSTRAINT CHECKING FUNCTIONS ====================

// Build clusters of entries that must be together (transitive closure)
// If A must be with B, and B must be with C, then A, B, C must all be together
function buildMustBeWithClusters() {
    const clusters = [];
    const processed = new Set();
    
    for (const pair of CHEAT_CONSTRAINTS.mustBeWith) {
        if (!Array.isArray(pair) || pair.length !== 2) continue;
        
        const [entry1, entry2] = pair.map(e => normalizeName(e));
        
        // Find existing cluster that contains either entry
        let foundCluster = null;
        for (const cluster of clusters) {
            if (cluster.has(entry1) || cluster.has(entry2)) {
                foundCluster = cluster;
                break;
            }
        }
        
        if (foundCluster) {
            // Add both entries to existing cluster
            foundCluster.add(entry1);
            foundCluster.add(entry2);
        } else {
            // Create new cluster
            clusters.push(new Set([entry1, entry2]));
        }
    }
    
    // Merge clusters that share entries (transitive closure)
    let changed = true;
    while (changed) {
        changed = false;
        for (let i = 0; i < clusters.length; i++) {
            for (let j = i + 1; j < clusters.length; j++) {
                // Check if clusters share any entry
                const hasOverlap = Array.from(clusters[i]).some(entry => clusters[j].has(entry));
                if (hasOverlap) {
                    // Merge clusters
                    clusters[j].forEach(entry => clusters[i].add(entry));
                    clusters.splice(j, 1);
                    changed = true;
                    break;
                }
            }
            if (changed) break;
        }
    }
    
    return clusters;
}

// Get the cluster that contains an entry (if any)
function getMustBeWithCluster(entryName) {
    const normalizedName = normalizeName(entryName);
    const clusters = buildMustBeWithClusters();
    
    for (const cluster of clusters) {
        if (cluster.has(normalizedName)) {
            return cluster;
        }
    }
    return null;
}

// Get all entries currently in a group
function getEntriesInGroup(groupName) {
    if (drawState.groups[groupName]) {
        return drawState.groups[groupName]
            .filter(entryData => entryData !== null && entryData !== undefined)
            .map(entryData => typeof entryData === 'string' ? entryData : entryData.entry);
    }
    return [];
}

// Check if placing an entry in a group violates "cannotBeWith" constraints
function checkCannotBeWith(entryName, groupName) {
    if (!CHEAT_CONSTRAINTS.enabled) return true;

    const entriesInGroup = getEntriesInGroup(groupName);

    for (const pair of CHEAT_CONSTRAINTS.cannotBeWith) {
        if (!Array.isArray(pair)) continue; // Skip invalid entries
        if (pair.includes(entryName)) {
            const otherEntry = pair.find(e => e !== entryName);
            if (entriesInGroup.includes(otherEntry)) {
                return false; // Violation: would be with forbidden partner
            }
        }
    }
    return true;
}

// Helper function to normalize names for comparison
function normalizeName(name) {
    return (name || '').trim().toLowerCase();
}

// Check if placing an entry in a group satisfies "mustBeWith" constraints
function checkMustBeWith(entryName, groupName) {
    if (!CHEAT_CONSTRAINTS.enabled) return true;

    const cluster = getMustBeWithCluster(entryName);
    if (!cluster) return true; // No constraints for this entry
    
    // Check if ANY member of the cluster is already placed
    let clusterGroup = null;
    for (const gName of config.groupNames) {
        const entries = getEntriesInGroup(gName);
        // Check if any cluster member is in this group
        const clusterMemberInGroup = Array.from(cluster).some(clusterEntry => 
            entries.some(e => normalizeName(e) === clusterEntry)
        );
        if (clusterMemberInGroup) {
            clusterGroup = gName;
            break;
        }
    }
    
    if (clusterGroup !== null) {
        // At least one cluster member is already placed - we MUST go to that group
        if (clusterGroup !== groupName) {
            const clusterMembers = Array.from(cluster).join(', ');
            console.log(`MustBeWith violation: ${entryName} must be with cluster [${clusterMembers}] who are in ${clusterGroup}, not ${groupName}`);
            return false; // Violation: cluster members are in different group
        } else {
            const clusterMembers = Array.from(cluster).join(', ');
            console.log(`MustBeWith satisfied: ${entryName} can be placed with cluster [${clusterMembers}] in ${groupName}`);
        }
    }
    // If no cluster members are placed yet, allow placement (they will be restricted when drawn)
    
    return true;
}

// Check "mustBeInGroup" constraint
function checkMustBeInGroup(entryName, groupName) {
    if (!CHEAT_CONSTRAINTS.enabled) return true;

    // Check both exact match and case-insensitive match
    const normalizedEntryName = normalizeName(entryName);
    let forcedGroup = CHEAT_CONSTRAINTS.mustBeInGroup[entryName];
    
    // If not found, try case-insensitive lookup
    if (!forcedGroup) {
        for (const [key, value] of Object.entries(CHEAT_CONSTRAINTS.mustBeInGroup)) {
            if (normalizeName(key) === normalizedEntryName) {
                forcedGroup = value;
                break;
            }
        }
    }
    
    if (forcedGroup) {
        // Compare group names (case-insensitive)
        if (normalizeName(forcedGroup) !== normalizeName(groupName)) {
            console.log(`MustBeInGroup violation: ${entryName} must be in ${forcedGroup}, not ${groupName}`);
            return false; // Violation: must be in a specific different group
        } else {
            console.log(`MustBeInGroup satisfied: ${entryName} is correctly placed in ${groupName}`);
        }
    }
    return true;
}

// Check "cannotBeInGroup" constraint
function checkCannotBeInGroup(entryName, groupName) {
    if (!CHEAT_CONSTRAINTS.enabled) return true;

    const forbiddenGroups = CHEAT_CONSTRAINTS.cannotBeInGroup[entryName];
    if (forbiddenGroups && forbiddenGroups.includes(groupName)) {
        return false; // Violation: cannot be in this group
    }
    return true;
}

// Master function: Check ALL constraints for placing an entry in a group
function isValidPlacement(entryName, groupName) {
    if (!CHEAT_CONSTRAINTS.enabled) return true;

    return (
        checkCannotBeWith(entryName, groupName) &&
        checkMustBeWith(entryName, groupName) &&
        checkMustBeInGroup(entryName, groupName) &&
        checkCannotBeInGroup(entryName, groupName)
    );
}

// Get valid groups for an entry (respecting all constraints)
function getValidGroupsForEntry(entryName, potIndex) {
    const availableGroups = getAvailableGroups(potIndex);

    if (!CHEAT_CONSTRAINTS.enabled) {
        return availableGroups;
    }

    // First, check if this entry has a mustBeWith constraint with someone already placed
    const cluster = getMustBeWithCluster(entryName);
    if (cluster) {
        // Find where any cluster member is already placed
        for (const gName of config.groupNames) {
            const entries = getEntriesInGroup(gName);
            const clusterMemberInGroup = Array.from(cluster).some(clusterEntry => 
                entries.some(e => normalizeName(e) === clusterEntry)
            );
            if (clusterMemberInGroup) {
                // Cluster member is here - check if this group is still available for this pot
                if (availableGroups.includes(gName)) {
                    const clusterMembers = Array.from(cluster).join(', ');
                    console.log(`${entryName} must be with cluster [${clusterMembers}] in ${gName}`);
                    return [gName]; // This is the only valid group
                } else {
                    // Group already has entry from this pot - constraint conflict!
                    console.warn(`CONSTRAINT CONFLICT: ${entryName} must be with cluster but ${gName} already has entry from pot ${potIndex}`);
                    // Return empty - this entry cannot be placed validly
                    return [];
                }
            }
        }
    }

    // Check mustBeInGroup constraint (from MustBeInTopic sheet)
    const normalizedEntryName = normalizeName(entryName);
    let forcedGroup = CHEAT_CONSTRAINTS.mustBeInGroup[entryName];
    
    // If not found, try case-insensitive lookup
    if (!forcedGroup) {
        for (const [key, value] of Object.entries(CHEAT_CONSTRAINTS.mustBeInGroup)) {
            if (normalizeName(key) === normalizedEntryName) {
                forcedGroup = value;
                break;
            }
        }
    }
    
    if (forcedGroup) {
        // Find the matching group (case-insensitive)
        const matchingGroup = config.groupNames.find(gName => 
            normalizeName(gName) === normalizeName(forcedGroup)
        );
        if (matchingGroup && availableGroups.includes(matchingGroup)) {
            console.log(`${entryName} must be in ${matchingGroup} due to MustBeInTopic constraint`);
            return [matchingGroup];
        } else if (matchingGroup) {
            console.warn(`CONSTRAINT CONFLICT: ${entryName} must be in ${matchingGroup} but it already has entry from pot ${potIndex}`);
            return [];
        }
    }

    // Apply other constraints (cannotBeWith, cannotBeInGroup) to available groups
    const validGroups = availableGroups.filter(groupName => isValidPlacement(entryName, groupName));
    
    return validGroups;
}

// Find an entry from the pot that has valid placements (for smart selection)
function findDrawableEntry(potEntries, potIndex) {
    // Prioritize entries that MUST be placed with someone already placed
    // This ensures mustBeWith constraints are satisfied early
    const entriesWithMustBeWith = [];
    const entriesWithMustBeInGroup = [];
    const otherEntries = [];
    
    for (const entry of potEntries) {
        const cluster = getMustBeWithCluster(entry);
        if (cluster) {
            // Check if any cluster member is already placed
            let clusterMemberPlaced = false;
            for (const gName of config.groupNames) {
                const entries = getEntriesInGroup(gName);
                if (Array.from(cluster).some(clusterEntry => 
                    entries.some(e => normalizeName(e) === clusterEntry)
                )) {
                    clusterMemberPlaced = true;
                    break;
                }
            }
            if (clusterMemberPlaced) {
                entriesWithMustBeWith.push(entry);
                continue;
            }
        }
        
        // Check mustBeInGroup
        const normalizedName = normalizeName(entry);
        let hasMustBeInGroup = CHEAT_CONSTRAINTS.mustBeInGroup[entry];
        if (!hasMustBeInGroup) {
            for (const [key, value] of Object.entries(CHEAT_CONSTRAINTS.mustBeInGroup)) {
                if (normalizeName(key) === normalizedName) {
                    hasMustBeInGroup = true;
                    break;
                }
            }
        }
        if (hasMustBeInGroup) {
            entriesWithMustBeInGroup.push(entry);
        } else {
            otherEntries.push(entry);
        }
    }
    
    // Shuffle each category
    const shuffleMustBeWith = entriesWithMustBeWith.sort(() => Math.random() - 0.5);
    const shuffleMustBeInGroup = entriesWithMustBeInGroup.sort(() => Math.random() - 0.5);
    const shuffleOther = otherEntries.sort(() => Math.random() - 0.5);
    
    // Try entries in priority order: mustBeWith first, then mustBeInGroup, then others
    const orderedEntries = [...shuffleMustBeWith, ...shuffleMustBeInGroup, ...shuffleOther];
    
    for (const entry of orderedEntries) {
        const validGroups = getValidGroupsForEntry(entry, potIndex);
        if (validGroups.length > 0) {
            return { entry, validGroups };
        }
    }

    // No valid entry found - constraints might be impossible
    // Fall back to random selection from available groups only
    console.warn('CHEAT WARNING: Could not satisfy all constraints. Falling back to available groups.');
    const randomEntry = potEntries[Math.floor(Math.random() * potEntries.length)];
    const availableGroups = getAvailableGroups(potIndex);
    if (availableGroups.length > 0) {
        return { entry: randomEntry, validGroups: availableGroups };
    }
    
    // Absolute fallback - no groups available (all have entry from this pot)
    console.error('CRITICAL: No available groups for pot', potIndex);
    return { entry: randomEntry, validGroups: config.groupNames };
}

// Mark entry as drawn
function markEntryAsDrawn(potIndex, entryIndex) {
    const ball = document.querySelector(`.team-ball[data-pot="${potIndex}"][data-entry="${entryIndex}"]`);
    if (ball) {
        ball.classList.add('drawn');
    }
}

// Sleep utility
function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// Update status message
function updateStatus(message) {
    const statusElement = document.getElementById('drawStatus');
    // Check if message contains HTML (like links)
    if (message.includes('<a ') || message.includes('<')) {
        statusElement.innerHTML = message;
    } else {
        statusElement.textContent = message;
    }
}

// Show draw animation
async function showDrawAnimation(entryName) {
    const overlay = document.getElementById('drawOverlay');
    const ballText = overlay.querySelector('.ball-text');

    ballText.innerHTML = `<span class="entry-name">${entryName}</span>`;

    overlay.classList.add('active');
    await sleep(1500);
    overlay.classList.remove('active');
}

// Highlight group
function highlightGroup(groupIndex, highlight = true) {
    const group = document.getElementById(`group${groupIndex}`);
    if (highlight) {
        group.classList.add('active');
    } else {
        group.classList.remove('active');
    }
}

// Highlight slot
function highlightSlot(groupIndex, slotIndex, highlight = true) {
    const group = document.getElementById(`group${groupIndex}`);
    const slot = group.querySelectorAll('.slot')[slotIndex];
    if (highlight) {
        slot.classList.add('highlight');
    } else {
        slot.classList.remove('highlight');
    }
}

// Draw single entry
async function drawEntry() {
    if (drawState.isDrawing || drawState.drawComplete) return;

    drawState.isDrawing = true;
    const drawBtn = document.getElementById('drawBtn');
    const autoDrawBtn = document.getElementById('autoDrawBtn');
    drawBtn.disabled = true;
    autoDrawBtn.disabled = true;

    // Find ANY pot with remaining entries - check ALL pots
    let currentPot = null;
    let potIndex = -1;

    for (let i = 0; i < drawState.pots.length; i++) {
        if (drawState.pots[i].entries.length > 0) {
            currentPot = drawState.pots[i];
            potIndex = i;
            break;
        }
    }

    // Only mark complete if ALL pots are empty
    if (!currentPot || currentPot.entries.length === 0) {
        // Double-check: verify ALL pots are actually empty
        let totalRemaining = 0;
        drawState.pots.forEach(pot => {
            totalRemaining += pot.entries.length;
        });
        
        if (totalRemaining === 0) {
            drawState.drawComplete = true;
            updateStatus('DRAW COMPLETE!');
            // Add voting button when draw completes
            if (typeof addVotingButtonToUI === 'function') {
                setTimeout(() => addVotingButtonToUI(), 500);
            }
            createConfetti();
        } else {
            updateStatus(`Error: ${totalRemaining} entries still remaining but no pot found!`);
        }
        drawState.isDrawing = false;
        return;
    }

    updateStatus(`Drawing from ${currentPot.name}...`);

    // SECRET CHEAT: Use constraint-aware selection
    // This finds an entry that CAN be placed somewhere valid, then picks a valid group
    const { entry: selectedEntry, validGroups } = findDrawableEntry(currentPot.entries, potIndex);

    // Find the index of the selected entry in the current pot
    const randomIndex = currentPot.entries.indexOf(selectedEntry);

    // Show draw animation
    await showDrawAnimation(selectedEntry);

    // SECRET CHEAT: Select from valid groups only (respecting constraints)
    // But distribute evenly - find group with fewest entries
    let selectedGroupName;
    if (validGroups.length > 0) {
        // Find group with fewest entries for even distribution
        let minEntries = Infinity;
        const groupsWithMinEntries = [];
        
        validGroups.forEach(groupName => {
            const entryCount = drawState.groups[groupName].length;
            if (entryCount < minEntries) {
                minEntries = entryCount;
                groupsWithMinEntries.length = 0;
                groupsWithMinEntries.push(groupName);
            } else if (entryCount === minEntries) {
                groupsWithMinEntries.push(groupName);
            }
        });
        
        // Randomly select from groups with minimum entries
        const randomGroupIndex = Math.floor(Math.random() * groupsWithMinEntries.length);
        selectedGroupName = groupsWithMinEntries[randomGroupIndex];
    } else {
        // Fallback if no valid groups (shouldn't happen)
        selectedGroupName = config.groupNames[Math.floor(Math.random() * config.groupNames.length)];
    }
    
    const groupIndex = config.groupNames.indexOf(selectedGroupName);

    // Highlight the group
    highlightGroup(groupIndex, true);
    // Don't highlight slot since we're not using potIndex-based slots anymore

    updateStatus(`${selectedEntry} → ${selectedGroupName}`);

    await sleep(500);

    // Place entry in group (append to array, allowing imbalance)
    // Store with pot index for color coding
    drawState.groups[selectedGroupName].push({entry: selectedEntry, potIndex: potIndex});

    // Find the original entry index before removal for marking
    const originalEntryIndex = config.pots[potIndex].entries.indexOf(selectedEntry);

    // Remove entry from pot
    drawState.pots[potIndex].entries = currentPot.entries.filter((_, index) => index !== randomIndex);

    // Mark as drawn
    markEntryAsDrawn(potIndex, originalEntryIndex);

    // Update UI
    renderDrawGroups();

    await sleep(500);

    // Remove highlights
    highlightGroup(groupIndex, false);

    // Check if draw is complete - ALL pots must be empty
    let totalRemaining = 0;
    drawState.pots.forEach(pot => {
        totalRemaining += pot.entries.length;
    });

    if (totalRemaining === 0) {
        drawState.drawComplete = true;
        updateStatus('DRAW COMPLETE!');
        createConfetti();
    } else {
        updateStatus(`Click "DRAW NEXT" to continue (${totalRemaining} remaining)`);
    }

    drawState.isDrawing = false;
    drawBtn.disabled = false;
    autoDrawBtn.disabled = false;
}

// Auto draw all (with animation)
async function autoDrawAll() {
    if (drawState.isDrawing || drawState.drawComplete) return;

    const drawBtn = document.getElementById('drawBtn');
    const autoDrawBtn = document.getElementById('autoDrawBtn');
    const instantDrawBtn = document.getElementById('instantDrawBtn');
    const abortBtn = document.getElementById('abortBtn');
    
    drawBtn.disabled = true;
    autoDrawBtn.disabled = true;
    if (instantDrawBtn) instantDrawBtn.disabled = true;
    if (abortBtn) abortBtn.style.display = 'inline-block';
    
    drawState.abortRequested = false;

    // Get animation duration from config (convert seconds to milliseconds)
    const animationDelay = (config.animationDuration || 0.8) * 1000;

    while (!drawState.drawComplete && !drawState.abortRequested) {
        await drawEntry();
        if (!drawState.drawComplete && !drawState.abortRequested) {
            await sleep(animationDelay);
        }
    }
    
    // Hide abort button and re-enable controls
    if (abortBtn) abortBtn.style.display = 'none';
    
    if (drawState.abortRequested) {
        drawState.abortRequested = false;
        drawBtn.disabled = false;
        autoDrawBtn.disabled = false;
        if (instantDrawBtn) instantDrawBtn.disabled = false;
        updateStatus('Draw aborted. Click "DRAW NEXT" to continue.');
    }
}

// Abort the current auto draw
function abortDraw() {
    drawState.abortRequested = true;
}

// Instant draw all (no animation, one fell swoop)
// Can be used to continue a draw that's already been started
function instantDrawAll() {
    if (drawState.isDrawing) return; // Don't allow if already drawing
    if (drawState.drawComplete) {
        updateStatus('Draw already complete!');
        return;
    }

    const drawBtn = document.getElementById('drawBtn');
    const autoDrawBtn = document.getElementById('autoDrawBtn');
    const instantDrawBtn = document.getElementById('instantDrawBtn');
    drawBtn.disabled = true;
    autoDrawBtn.disabled = true;
    if (instantDrawBtn) instantDrawBtn.disabled = true;

    drawState.isDrawing = true;

    // Draw all entries instantly - continue until ALL entries from ALL pots are distributed
    let iterations = 0;
    const maxIterations = 1000; // Safety limit
    
    while (iterations < maxIterations) {
        // Count total remaining entries across ALL pots
        let totalRemaining = 0;
        drawState.pots.forEach(pot => {
            totalRemaining += pot.entries.length;
        });
        
        // If no entries remain, we're done
        if (totalRemaining === 0) {
            drawState.drawComplete = true;
            break;
        }
        
        // Find ANY pot with remaining entries
        let currentPot = null;
        let potIndex = -1;

        for (let i = 0; i < drawState.pots.length; i++) {
            if (drawState.pots[i].entries.length > 0) {
                currentPot = drawState.pots[i];
                potIndex = i;
                break;
            }
        }

        // If we can't find a pot but there are still entries, something is wrong
        if (!currentPot || currentPot.entries.length === 0) {
            console.error(`Error: ${totalRemaining} entries remaining but no pot found!`);
            break;
        }

        // Get entry and valid groups
        const { entry: selectedEntry, validGroups } = findDrawableEntry(currentPot.entries, potIndex);
        
        // Distribute evenly - find group with fewest entries
        let selectedGroupName;
        if (validGroups.length > 0) {
            let minEntries = Infinity;
            const groupsWithMinEntries = [];
            
            validGroups.forEach(groupName => {
                const entryCount = drawState.groups[groupName].length;
                if (entryCount < minEntries) {
                    minEntries = entryCount;
                    groupsWithMinEntries.length = 0;
                    groupsWithMinEntries.push(groupName);
                } else if (entryCount === minEntries) {
                    groupsWithMinEntries.push(groupName);
                }
            });
            
            // Randomly select from groups with minimum entries
            const randomGroupIndex = Math.floor(Math.random() * groupsWithMinEntries.length);
            selectedGroupName = groupsWithMinEntries[randomGroupIndex];
        } else {
            // Fallback - just pick any group
            selectedGroupName = config.groupNames[Math.floor(Math.random() * config.groupNames.length)];
        }

        // Place entry in group (append to array, allowing imbalance)
        // Store with pot index for color coding
        drawState.groups[selectedGroupName].push({entry: selectedEntry, potIndex: potIndex});

        // Remove entry from pot
        const entryIndex = currentPot.entries.indexOf(selectedEntry);
        if (entryIndex !== -1) {
            drawState.pots[potIndex].entries.splice(entryIndex, 1);
        }

        // Mark as drawn
        const originalEntryIndex = config.pots[potIndex].entries.indexOf(selectedEntry);
        if (originalEntryIndex !== -1) {
            markEntryAsDrawn(potIndex, originalEntryIndex);
        }
        
        iterations++;
    }
    
    // Final check
    let finalTotal = 0;
    drawState.pots.forEach(pot => {
        finalTotal += pot.entries.length;
    });
    
    if (finalTotal > 0) {
        console.error(`Warning: Draw stopped with ${finalTotal} entries still remaining after ${iterations} iterations`);
    }

    // Update UI once at the end
    renderDrawGroups();
    renderDrawPots();

    // Check if draw is complete
    let totalRemaining = 0;
    drawState.pots.forEach(pot => {
        totalRemaining += pot.entries.length;
    });

    if (totalRemaining === 0) {
        drawState.drawComplete = true;
        updateStatus('DRAW COMPLETE!');
        createConfetti();
    } else {
        updateStatus('DRAW COMPLETE!');
    }

    drawState.isDrawing = false;
    drawBtn.disabled = false;
    autoDrawBtn.disabled = false;
    if (instantDrawBtn) instantDrawBtn.disabled = false;
    
    // Add voting button when draw completes
    if (typeof addVotingButtonToUI === 'function') {
        setTimeout(() => addVotingButtonToUI(), 500);
    }
}

// Reset draw
function resetDraw() {
    // Reset draw state with fresh copy from config
    drawState = {
        pots: JSON.parse(JSON.stringify(config.pots)),
        groups: {},
        isDrawing: false,
        drawComplete: false
    };

    // Initialize groups as empty arrays (can grow to accommodate all entries)
    config.groupNames.forEach(name => {
        drawState.groups[name] = [];
    });

    renderDrawPots();
    renderDrawGroups();

    document.getElementById('drawBtn').disabled = false;
    document.getElementById('autoDrawBtn').disabled = false;
    const instantDrawBtn = document.getElementById('instantDrawBtn');
    if (instantDrawBtn) instantDrawBtn.disabled = false;

    updateStatus('Click "DRAW NEXT" to begin');
}

// Go back to configuration
function reconfigure() {
    drawScreen.classList.add('hidden');
    setupScreen.classList.remove('hidden');
    hideValidation();
}

// Create confetti effect
function createConfetti() {
    const colors = ['#ffd700', '#ff6b6b', '#4ecdc4', '#45b7d1', '#96e6a1', '#dda0dd'];

    for (let i = 0; i < 100; i++) {
        const confetti = document.createElement('div');
        confetti.className = 'confetti';
        confetti.style.left = Math.random() * 100 + 'vw';
        confetti.style.background = colors[Math.floor(Math.random() * colors.length)];
        confetti.style.animationDelay = Math.random() * 2 + 's';
        confetti.style.animationDuration = (Math.random() * 2 + 2) + 's';
        document.body.appendChild(confetti);

        setTimeout(() => {
            confetti.remove();
        }, 5000);
    }
}

// ==================== CONFIG INITIALIZATION ====================

// Load configuration from config.js and populate form fields
function initializeConfig() {
    // Check if config file is loaded
    if (typeof APP_CONFIG === 'undefined') {
        console.warn('config.js not found or APP_CONFIG not defined. Using defaults.');
        return;
    }

    // Load Google Sheets API key
    if (APP_CONFIG.googleSheets?.apiKey) {
        const apiKeyInput = document.getElementById('googleApiKey');
        if (apiKeyInput) {
            apiKeyInput.value = APP_CONFIG.googleSheets.apiKey;
            apiKeyInput.placeholder = 'Loaded from config.js';
            // Make sure it's not disabled
            apiKeyInput.disabled = false;
            // Force a visual update
            apiKeyInput.style.color = 'white';
            apiKeyInput.style.opacity = '1';
            console.log('API Key loaded from config.js:', APP_CONFIG.googleSheets.apiKey.substring(0, 10) + '...');
        } else {
            console.warn('googleApiKey input field not found');
        }
    } else {
        console.warn('No API key found in config.js');
    }

    // Load default sheet name
    if (APP_CONFIG.googleSheets?.defaultSheetName) {
        const sheetNameInput = document.getElementById('sheetName');
        if (sheetNameInput && !sheetNameInput.value) {
            sheetNameInput.value = APP_CONFIG.googleSheets.defaultSheetName;
        }
    }

    // Load last Google Sheet URL (from localStorage or config)
    const lastSheetUrl = localStorage.getItem('lastGoogleSheetUrl') || APP_CONFIG.googleSheets?.lastSheetUrl || '';
    if (lastSheetUrl) {
        const sheetUrlInput = document.getElementById('googleSheetUrl');
        if (sheetUrlInput) {
            sheetUrlInput.value = lastSheetUrl;
            sheetUrlInput.placeholder = 'Loaded from last session';
            // Also ensure it's saved to localStorage as backup
            localStorage.setItem('lastGoogleSheetUrl', lastSheetUrl);
            console.log('Sheet URL loaded and saved to localStorage:', lastSheetUrl);
            
            // Auto-detect structure for the loaded URL
            setTimeout(async () => {
                const apiKey = document.getElementById('googleApiKey')?.value.trim() || 
                              (typeof APP_CONFIG !== 'undefined' ? APP_CONFIG.googleSheets?.apiKey || '' : '');
                const sheetName = document.getElementById('sheetName')?.value.trim() || 'Participants';
                
                if (apiKey) {
                    try {
                        const structure = await detectSheetStructure(apiKey, lastSheetUrl, sheetName);
                        if (structure && structure.numPots > 0 && structure.numGroups > 0) {
                            const numPotsInput = document.getElementById('numPots');
                            const numGroupsInput = document.getElementById('numGroups');
                            
                            if (numPotsInput) {
                                numPotsInput.value = structure.numPots;
                            }
                            if (numGroupsInput) {
                                numGroupsInput.value = structure.numGroups;
                            }
                            
                            const urlSaveStatus = document.getElementById('urlSaveStatus');
                            if (urlSaveStatus) {
                                urlSaveStatus.textContent = `✓ Auto-detected: ${structure.numPots} pots, ${structure.numGroups} groups`;
                                urlSaveStatus.style.color = '#4CAF50';
                                setTimeout(() => {
                                    if (urlSaveStatus) urlSaveStatus.textContent = '';
                                }, 3000);
                            }
                            console.log('Auto-detected structure on page load:', structure);
                        }
                    } catch (error) {
                        console.log('Auto-detection on page load failed:', error.message);
                    }
                }
            }, 500);
        }
    }
    
    // Add event listener to save URL to localStorage whenever it's typed/pasted
    const sheetUrlInput = document.getElementById('googleSheetUrl');
    const detectStructureBtn = document.getElementById('detectStructureBtn');
    const urlSaveStatus = document.getElementById('urlSaveStatus');
    
    // Function to detect structure (reusable)
    async function triggerStructureDetection() {
        const url = sheetUrlInput?.value.trim() || localStorage.getItem('lastGoogleSheetUrl') || '';
        if (!url) {
            if (urlSaveStatus) {
                urlSaveStatus.textContent = 'Please enter a Google Sheets URL first';
                urlSaveStatus.style.color = '#ff6b6b';
                setTimeout(() => { if (urlSaveStatus) urlSaveStatus.textContent = ''; }, 2000);
            }
            return;
        }

        // Save URL to localStorage
        localStorage.setItem('lastGoogleSheetUrl', url);
        
        if (urlSaveStatus) {
            urlSaveStatus.textContent = 'Detecting structure...';
            urlSaveStatus.style.color = '#ffd700';
        }

        const apiKey = document.getElementById('googleApiKey')?.value.trim() || 
                      (typeof APP_CONFIG !== 'undefined' ? APP_CONFIG.googleSheets?.apiKey || '' : '');
        const sheetName = document.getElementById('sheetName')?.value.trim() || 'Participants';

        if (!apiKey) {
            if (urlSaveStatus) {
                urlSaveStatus.textContent = 'Please enter your API key first';
                urlSaveStatus.style.color = '#ff6b6b';
                setTimeout(() => { if (urlSaveStatus) urlSaveStatus.textContent = ''; }, 2000);
            }
            return;
        }

        try {
            const structure = await detectSheetStructure(apiKey, url, sheetName);
            if (structure && structure.numPots > 0 && structure.numGroups > 0) {
                // Auto-populate the fields
                const numPotsInput = document.getElementById('numPots');
                const numGroupsInput = document.getElementById('numGroups');
                
                if (numPotsInput) {
                    numPotsInput.value = structure.numPots;
                }
                if (numGroupsInput) {
                    numGroupsInput.value = structure.numGroups;
                }

                if (urlSaveStatus) {
                    urlSaveStatus.textContent = `✓ Detected: ${structure.numPots} pots, ${structure.numGroups} groups`;
                    urlSaveStatus.style.color = '#4CAF50';
                }
                console.log('Detected structure:', structure);
            } else {
                if (urlSaveStatus) {
                    urlSaveStatus.textContent = 'Could not detect structure (empty sheet?)';
                    urlSaveStatus.style.color = '#ff6b6b';
                }
            }
        } catch (error) {
            if (urlSaveStatus) {
                urlSaveStatus.textContent = 'Error: ' + error.message;
                urlSaveStatus.style.color = '#ff6b6b';
            }
            console.error('Detection error:', error);
        }
    }
    
    if (sheetUrlInput) {
        // Manual detect button
        if (detectStructureBtn) {
            detectStructureBtn.addEventListener('click', triggerStructureDetection);
        }
        
        // Auto-detect structure when URL is entered
        async function autoDetectStructure(url) {
            if (!url) return;
            
            const apiKey = document.getElementById('googleApiKey')?.value.trim() || 
                          (typeof APP_CONFIG !== 'undefined' ? APP_CONFIG.googleSheets?.apiKey || '' : '');
            const sheetName = document.getElementById('sheetName')?.value.trim() || 'Participants';

            if (!apiKey) return; // Can't detect without API key

            try {
                const structure = await detectSheetStructure(apiKey, url, sheetName);
                if (structure && structure.numPots > 0 && structure.numGroups > 0) {
                    // Auto-populate the fields
                    const numPotsInput = document.getElementById('numPots');
                    const numGroupsInput = document.getElementById('numGroups');
                    
                    if (numPotsInput && !numPotsInput.value) {
                        numPotsInput.value = structure.numPots;
                    }
                    if (numGroupsInput && !numGroupsInput.value) {
                        numGroupsInput.value = structure.numGroups;
                    }

                    if (urlSaveStatus) {
                        urlSaveStatus.textContent = `✓ Detected: ${structure.numPots} pots, ${structure.numGroups} groups`;
                        urlSaveStatus.style.color = '#4CAF50';
                        setTimeout(() => {
                            if (urlSaveStatus) urlSaveStatus.textContent = '';
                        }, 3000);
                    }
                    console.log('Auto-detected structure:', structure);
                }
            } catch (error) {
                console.log('Auto-detection failed:', error.message);
            }
        }
        
        // Save on input
        sheetUrlInput.addEventListener('input', async (e) => {
            const url = e.target.value.trim();
            if (url) {
                localStorage.setItem('lastGoogleSheetUrl', url);
                console.log('Sheet URL saved to localStorage (input):', url);
                // Auto-detect after a short delay
                setTimeout(() => autoDetectStructure(url), 500);
            }
        });
        
        // Save on paste
        sheetUrlInput.addEventListener('paste', async (e) => {
            setTimeout(() => {
                const url = e.target.value.trim();
                if (url) {
                    localStorage.setItem('lastGoogleSheetUrl', url);
                    console.log('Sheet URL pasted and saved to localStorage:', url);
                    // Auto-detect
                    autoDetectStructure(url);
                }
            }, 100);
        });
    }

    // Load Google Drive Client ID (for both Drive section and export field)
    // Check config.js first, then localStorage, then use empty
    const configClientId = APP_CONFIG.googleDrive?.clientId || '';
    const savedClientId = localStorage.getItem('googleDriveClientId') || '';
    const clientIdToUse = configClientId || savedClientId;
    
    if (clientIdToUse) {
        const clientIdInput = document.getElementById('googleDriveClientId');
        const exportClientIdInput = document.getElementById('googleDriveClientIdForExport');
        
        if (clientIdInput) {
            clientIdInput.value = clientIdToUse;
            clientIdInput.placeholder = configClientId ? 'Loaded from config.js' : 'Loaded from last session';
            // Save to localStorage as backup
            localStorage.setItem('googleDriveClientId', clientIdToUse);
        }
        if (exportClientIdInput) {
            exportClientIdInput.value = clientIdToUse;
            exportClientIdInput.placeholder = configClientId ? 'Loaded from config.js' : 'Loaded from last session';
            // Save to localStorage as backup
            localStorage.setItem('googleDriveClientId', clientIdToUse);
        }
    }
    
    // Add event listeners to save OAuth Client ID to localStorage when changed
    const clientIdInput = document.getElementById('googleDriveClientId');
    const exportClientIdInput = document.getElementById('googleDriveClientIdForExport');
    
    function saveClientIdToStorage(value) {
        if (value && value.trim()) {
            localStorage.setItem('googleDriveClientId', value.trim());
            console.log('OAuth Client ID saved to localStorage');
        }
    }
    
    if (clientIdInput) {
        clientIdInput.addEventListener('input', (e) => {
            saveClientIdToStorage(e.target.value);
            // Also update export field if it exists
            if (exportClientIdInput) {
                exportClientIdInput.value = e.target.value;
            }
        });
        clientIdInput.addEventListener('paste', (e) => {
            setTimeout(() => {
                saveClientIdToStorage(e.target.value);
                if (exportClientIdInput) {
                    exportClientIdInput.value = e.target.value;
                }
            }, 10);
        });
    }
    
    if (exportClientIdInput) {
        exportClientIdInput.addEventListener('input', (e) => {
            saveClientIdToStorage(e.target.value);
            // Also update Drive field if it exists
            if (clientIdInput) {
                clientIdInput.value = e.target.value;
            }
        });
        exportClientIdInput.addEventListener('paste', (e) => {
            setTimeout(() => {
                saveClientIdToStorage(e.target.value);
                if (clientIdInput) {
                    clientIdInput.value = e.target.value;
                }
            }, 10);
        });
    }

    // Load default event settings
    if (APP_CONFIG.defaults) {
        if (APP_CONFIG.defaults.eventTitle) {
            const eventTitleInput = document.getElementById('eventTitle');
            if (eventTitleInput && !eventTitleInput.value) {
                eventTitleInput.value = APP_CONFIG.defaults.eventTitle;
            }
        }
        if (APP_CONFIG.defaults.numGroups) {
            const numGroupsInput = document.getElementById('numGroups');
            if (numGroupsInput && !numGroupsInput.value) {
                numGroupsInput.value = APP_CONFIG.defaults.numGroups;
            }
        }
        if (APP_CONFIG.defaults.numPots) {
            const numPotsInput = document.getElementById('numPots');
            if (numPotsInput && !numPotsInput.value) {
                numPotsInput.value = APP_CONFIG.defaults.numPots;
            }
        }
        if (APP_CONFIG.defaults.animationDuration) {
            const animationDurationInput = document.getElementById('animationDuration');
            if (animationDurationInput) {
                animationDurationInput.value = APP_CONFIG.defaults.animationDuration;
            }
            config.animationDuration = APP_CONFIG.defaults.animationDuration;
        }
    }
    
    // Load animation duration from localStorage (always restore to input)
    const savedAnimationDuration = localStorage.getItem('animationDuration');
    if (savedAnimationDuration) {
        const animationDurationInput = document.getElementById('animationDuration');
        if (animationDurationInput) {
            animationDurationInput.value = savedAnimationDuration;
        }
        // Also set in config
        config.animationDuration = parseFloat(savedAnimationDuration) || 0.8;
    }
    
    // Add event listener to save animation duration when changed
    const animationDurationInput = document.getElementById('animationDuration');
    if (animationDurationInput) {
        animationDurationInput.addEventListener('input', (e) => {
            const value = parseFloat(e.target.value) || 0.8;
            localStorage.setItem('animationDuration', value.toString());
            config.animationDuration = value;
        });
    }
}

// Initialize config when DOM is ready
function initConfigWhenReady() {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => {
            initializeConfig();
            // Also try after a short delay in case sections are hidden
            setTimeout(initializeConfig, 100);
        });
    } else {
        // DOM already loaded
        initializeConfig();
        // Also try after a short delay in case sections are hidden
        setTimeout(initializeConfig, 100);
    }
}

initConfigWhenReady();

// ==================== EVENT LISTENERS ====================
document.getElementById('drawBtn').addEventListener('click', drawEntry);
document.getElementById('autoDrawBtn').addEventListener('click', autoDrawAll);
document.getElementById('instantDrawBtn').addEventListener('click', instantDrawAll);
document.getElementById('abortBtn').addEventListener('click', abortDraw);
document.getElementById('exportBtn').addEventListener('click', exportToGoogleSheet);
document.getElementById('resetBtn').addEventListener('click', resetDraw);
document.getElementById('reconfigureBtn').addEventListener('click', reconfigure);
