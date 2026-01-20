// ============================================================================
// Hackathon Team Voting Solution
// ============================================================================
// This file provides functionality to generate Google Forms for anonymous
// team voting, ensuring team members cannot vote for their own team.
//
// REQUIREMENTS:
// 1. Anonymous voting
// 2. Team members cannot vote for their own team
// 3. Support for multiple forms if needed
//
// USAGE:
// After the draw is complete, call generateVotingForms() to get form
// configurations and participant assignments.
// ============================================================================

// Store token locally - don't rely on app.js
let _votingAccessToken = null;

// Store participant emails globally - loaded once when spreadsheet is loaded
let _participantEmailMap = {};

// ==================== EMAIL LOADING (UPFRONT) ====================

/**
 * Loads all participant emails from the spreadsheet and Google Contacts.
 * Call this BEFORE generating forms. Fails fast if any emails are missing.
 * @param {string} accessToken - OAuth access token
 * @returns {Promise<Object>} Object with emailMap and any missing names
 */
async function loadAllParticipantEmails(accessToken) {
    console.log('=== LOADING ALL PARTICIPANT EMAILS ===');
    
    const spreadsheetUrl = getSpreadsheetUrl();
    if (!spreadsheetUrl) {
        throw new Error('No spreadsheet URL found. Please load a spreadsheet first.');
    }
    
    const spreadsheetId = extractSpreadsheetId(spreadsheetUrl) || extractFileIdFromDriveUrl(spreadsheetUrl);
    if (!spreadsheetId) {
        throw new Error('Could not extract spreadsheet ID from URL');
    }
    
    const sheetName = document.getElementById('sheetName')?.value.trim() || 'Participants';
    console.log(`Loading from spreadsheet: ${spreadsheetId}, sheet: ${sheetName}`);
    
    // Step 1: Get all names from the spreadsheet
    const allNames = new Set();
    const range = `${sheetName}!A1:Z1000`;
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?ranges=${encodeURIComponent(range)}&includeGridData=true`;
    
    const response = await fetch(url, {
        headers: { 'Authorization': `Bearer ${accessToken}` }
    });
    
    if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(`Failed to fetch spreadsheet: ${response.status} ${JSON.stringify(errorData)}`);
    }
    
    const data = await response.json();
    const gridData = data.sheets?.[0]?.data?.[0];
    
    if (!gridData?.rowData) {
        throw new Error('No data found in spreadsheet');
    }
    
    // Collect all names (stop at empty row)
    for (let rowIndex = 1; rowIndex < gridData.rowData.length; rowIndex++) {
        const row = gridData.rowData[rowIndex];
        
        if (!row?.values || row.values.length === 0) {
            console.log(`Row ${rowIndex + 1} is empty - stopping`);
            break;
        }
        
        const firstCell = row.values[0];
        const firstCellValue = firstCell?.effectiveValue?.stringValue || 
                              firstCell?.formattedValue || '';
        if (!firstCellValue.trim()) {
            console.log(`Row ${rowIndex + 1} has empty first cell - stopping`);
            break;
        }
        
        // Collect names from all cells in this row
        for (const cell of row.values) {
            const name = cell?.effectiveValue?.stringValue || 
                        cell?.formattedValue || '';
            if (name.trim()) {
                const cleanName = name.trim().startsWith('@') ? name.trim().substring(1) : name.trim();
                allNames.add(cleanName);
            }
        }
    }
    
    console.log(`Found ${allNames.size} unique names in spreadsheet`);
    
    // Step 2: Load ALL contacts from Google People API
    console.log('Loading contacts from Google People API...');
    const emailMap = {};
    
    const peopleResponse = await fetch(
        'https://people.googleapis.com/v1/people/me/connections?personFields=names,emailAddresses&pageSize=1000',
        { headers: { 'Authorization': `Bearer ${accessToken}` } }
    );
    
    if (!peopleResponse.ok) {
        const errorData = await peopleResponse.json().catch(() => ({}));
        if (peopleResponse.status === 403) {
            throw new Error('People API not enabled. Enable it at: https://console.cloud.google.com/apis/library/people.googleapis.com');
        }
        throw new Error(`People API failed: ${peopleResponse.status} ${JSON.stringify(errorData)}`);
    }
    
    const peopleData = await peopleResponse.json();
    const connections = peopleData.connections || [];
    console.log(`Loaded ${connections.length} contacts from Google Contacts`);
    
    // Build contact lookup (name -> email)
    const contactLookup = {};
    for (const person of connections) {
        const names = person.names || [];
        const emails = person.emailAddresses || [];
        
        if (emails.length > 0) {
            const primaryEmail = emails[0].value;
            
            for (const nameObj of names) {
                const displayName = (nameObj.displayName || '').toLowerCase().trim();
                const givenName = (nameObj.givenName || '').toLowerCase().trim();
                const familyName = (nameObj.familyName || '').toLowerCase().trim();
                
                if (displayName) contactLookup[displayName] = primaryEmail;
                if (givenName && familyName) contactLookup[`${givenName} ${familyName}`] = primaryEmail;
                if (givenName) contactLookup[givenName] = primaryEmail;
            }
        }
    }
    
    // Step 3: Match names to emails
    const missingEmails = [];
    
    // Debug: show all contacts loaded
    console.log('Contact lookup keys:', Object.keys(contactLookup).slice(0, 20));
    
    for (const name of allNames) {
        const lowerName = name.toLowerCase().trim();
        let email = contactLookup[lowerName];
        let matchType = 'exact';
        
        // Try partial match if exact match fails
        if (!email) {
            // Try splitting first/last name
            const nameParts = lowerName.split(/\s+/);
            if (nameParts.length >= 2) {
                const firstName = nameParts[0];
                const lastName = nameParts[nameParts.length - 1];
                
                // Try "first last" and "last first"
                if (contactLookup[`${firstName} ${lastName}`]) {
                    email = contactLookup[`${firstName} ${lastName}`];
                    matchType = 'first-last';
                } else if (contactLookup[firstName]) {
                    email = contactLookup[firstName];
                    matchType = 'first-only';
                }
            }
        }
        
        // Try partial/contains match
        if (!email) {
            for (const [contactName, contactEmail] of Object.entries(contactLookup)) {
                if (contactName.includes(lowerName) || lowerName.includes(contactName)) {
                    email = contactEmail;
                    matchType = 'partial';
                    break;
                }
            }
        }
        
        if (email) {
            // Store with multiple key variations for robust lookup
            emailMap[name] = email;
            emailMap[name.toLowerCase()] = email;
            emailMap[name.trim()] = email;
            emailMap[`@${name}`] = email;
            emailMap[`@${name.trim()}`] = email;
            console.log(`✓ "${name}" → ${email} (${matchType})`);
        } else {
            missingEmails.push(name);
            console.error(`✗ "${name}" → NO EMAIL FOUND (tried: "${lowerName}")`);
        }
    }
    
    console.log(`\n=== EMAIL LOADING COMPLETE ===`);
    console.log(`Found: ${Object.keys(emailMap).length / 2} emails`); // Divide by 2 because we store with and without @
    console.log(`Missing: ${missingEmails.length} emails`);
    
    // Store globally
    _participantEmailMap = emailMap;
    
    return {
        emailMap,
        missingEmails,
        totalNames: allNames.size
    };
}

// ==================== VOTING FORM GENERATOR ====================

/**
 * Extracts participant emails from spreadsheet data
 * Handles Google contact tags by fetching detailed cell data with hyperlinks
 * @param {string} accessToken - OAuth access token
 * @param {string} spreadsheetId - Spreadsheet ID
 * @param {string} sheetName - Sheet name
 * @param {Array<Array>} sheetData - Raw spreadsheet values (fallback if detailed fetch fails)
 * @returns {Promise<Object>} Mapping of participant name to email
 */
async function extractParticipantEmails(accessToken, spreadsheetId, sheetName, sheetData = null) {
    const emailMap = {};
    
    // First, try to get detailed cell data with hyperlinks to extract emails from contact tags
    if (accessToken && spreadsheetId) {
        try {
            const range = `${sheetName}!A1:Z1000`;
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?ranges=${encodeURIComponent(range)}&includeGridData=true`;
            
            console.log('Fetching detailed cell data with includeGridData=true...');
            const response = await fetch(url, {
                headers: { 'Authorization': `Bearer ${accessToken}` }
            });
            
            if (!response.ok) {
                const errorData = await response.json().catch(() => ({}));
                console.error('Failed to fetch grid data:', response.status, errorData);
                throw new Error(`API error: ${response.status} ${JSON.stringify(errorData)}`);
            }
            
            const data = await response.json();
            console.log('Grid data response structure:', {
                hasSheets: !!data.sheets,
                sheetsLength: data.sheets?.length,
                firstSheetHasData: !!data.sheets?.[0]?.data,
                dataLength: data.sheets?.[0]?.data?.length
            });
            
            const sheet = data.sheets?.[0];
            const gridData = sheet?.data?.[0];
                
                if (gridData && gridData.rowData) {
                    console.log(`Processing ${gridData.rowData.length} rows of grid data...`);
                    // Iterate through ALL cells to find contact tags
                    // Skip header row (index 0) - process all data rows
                    // STOP at first empty row (to avoid sum rows at bottom)
                    for (let rowIndex = 1; rowIndex < gridData.rowData.length; rowIndex++) {
                        const row = gridData.rowData[rowIndex];
                        
                        // Check if row is empty - stop processing if so
                        if (!row?.values || row.values.length === 0) {
                            console.log(`Row ${rowIndex + 1} is empty - stopping processing`);
                            break;
                        }
                        
                        // Check if first cell is empty (indicates end of data)
                        const firstCell = row.values[0];
                        const firstCellValue = firstCell?.effectiveValue?.stringValue || 
                                              firstCell?.formattedValue || 
                                              firstCell?.userEnteredValue?.stringValue || '';
                        if (!firstCellValue.trim()) {
                            console.log(`Row ${rowIndex + 1} has empty first cell - stopping processing`);
                            break;
                        }
                        
                        // Process each cell in the row
                        for (let colIndex = 0; colIndex < row.values.length; colIndex++) {
                            const cell = row.values[colIndex];
                            if (!cell) continue;
                            
                            // Extract name from cell
                            let name = cell?.effectiveValue?.stringValue || 
                                      cell?.formattedValue || 
                                      cell?.userEnteredValue?.stringValue || '';
                            name = name.trim();
                            
                            if (!name) continue; // Skip empty cells
                            
                            // Extract email from contact tag - try multiple methods
                            let email = '';
                            
                            // Method 1: Check for hyperlink property (contact tags have mailto: links)
                            if (cell.hyperlink) {
                                const link = typeof cell.hyperlink === 'string' ? cell.hyperlink : cell.hyperlink.uri;
                                if (link && link.startsWith('mailto:')) {
                                    email = link.replace('mailto:', '').split('?')[0].trim();
                                }
                            }
                            
                            // Method 2: Check userEnteredFormat for hyperlink
                            if (!email && cell.userEnteredFormat?.link) {
                                const link = cell.userEnteredFormat.link;
                                const linkUri = typeof link === 'string' ? link : link.uri;
                                if (linkUri && linkUri.startsWith('mailto:')) {
                                    email = linkUri.replace('mailto:', '').split('?')[0].trim();
                                }
                            }
                            
                            // Method 2b: Check effectiveFormat for hyperlink
                            if (!email && cell.effectiveFormat?.link) {
                                const link = cell.effectiveFormat.link;
                                const linkUri = typeof link === 'string' ? link : link.uri;
                                if (linkUri && linkUri.startsWith('mailto:')) {
                                    email = linkUri.replace('mailto:', '').split('?')[0].trim();
                                }
                            }
                            
                            // Debug: Log cell structure for ALL cells with names (to see what we're working with)
                            if (name) {
                                console.log(`Cell [${rowIndex},${colIndex}] "${name}":`, {
                                    hasHyperlink: !!cell.hyperlink,
                                    hyperlink: cell.hyperlink,
                                    hasUserEnteredFormat: !!cell.userEnteredFormat,
                                    userEnteredFormatLink: cell.userEnteredFormat?.link,
                                    hasEffectiveFormat: !!cell.effectiveFormat,
                                    effectiveFormatLink: cell.effectiveFormat?.link,
                                    userEnteredValue: cell.userEnteredValue,
                                    effectiveValue: cell.effectiveValue,
                                    formattedValue: cell.formattedValue
                                });
                            }
                            
                            // Method 3: Check formula for contact reference
                            if (!email && cell.userEnteredValue?.formulaValue) {
                                const formula = cell.userEnteredValue.formulaValue;
                                // Look for email pattern in formula
                                const emailMatch = formula.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/);
                                if (emailMatch) {
                                    email = emailMatch[1];
                                }
                            }
                            
                            // Method 4: Check noteText (contact tags sometimes store email in notes)
                            if (!email && cell.note) {
                                const noteText = cell.note;
                                const emailMatch = noteText.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/);
                                if (emailMatch) {
                                    email = emailMatch[1];
                                }
                            }
                            
                            // Method 5: Check if hyperlink contains contact ID and try to extract email from it
                            // Contact tags might have links like: https://contacts.google.com/person/...
                            if (!email && cell.hyperlink) {
                                const link = typeof cell.hyperlink === 'string' ? cell.hyperlink : cell.hyperlink.uri;
                                // Sometimes contact tags have the email in the link itself
                                if (link) {
                                    const emailMatch = link.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/);
                                    if (emailMatch) {
                                        email = emailMatch[1];
                                    }
                                }
                            }
                            
                            // Store name for People API lookup if no email found yet
                            if (!email && name) {
                                // We'll do People API lookup after processing all cells
                                const cleanName = name.startsWith('@') ? name.substring(1) : name;
                                if (!emailMap[name] && !emailMap[cleanName]) {
                                    // Mark for People API lookup
                                    if (!window._pendingPeopleApiLookups) {
                                        window._pendingPeopleApiLookups = new Set();
                                    }
                                    window._pendingPeopleApiLookups.add(cleanName);
                                }
                            }
                            
                            // Store mapping if we have both name and email
                            if (name && email && email.includes('@')) {
                                const nameWithoutAt = name.startsWith('@') ? name.substring(1) : name;
                                emailMap[name] = email;
                                emailMap[nameWithoutAt] = email; // Also allow lookup without @
                                if (name.startsWith('@')) {
                                    emailMap[`@${nameWithoutAt}`] = email; // Also with @
                                }
                                console.log(`Found contact tag: "${name}" → ${email}`);
                            }
                        }
                    }
                    
                    console.log(`Extracted ${Object.keys(emailMap).length} participant emails from contact tags`);
                    
                    // Now do People API lookup for names without emails
                    if (window._pendingPeopleApiLookups && window._pendingPeopleApiLookups.size > 0) {
                        console.log(`Looking up ${window._pendingPeopleApiLookups.size} names via People API...`);
                        
                        try {
                            // First, get all contacts
                            const peopleResponse = await fetch(
                                'https://people.googleapis.com/v1/people/me/connections?personFields=names,emailAddresses&pageSize=1000',
                                { headers: { 'Authorization': `Bearer ${accessToken}` } }
                            );
                            
                            if (peopleResponse.ok) {
                                const peopleData = await peopleResponse.json();
                                const connections = peopleData.connections || [];
                                
                                console.log(`Found ${connections.length} contacts in Google Contacts`);
                                
                                // Build a lookup map from contact names to emails
                                const contactLookup = {};
                                for (const person of connections) {
                                    const names = person.names || [];
                                    const emails = person.emailAddresses || [];
                                    
                                    if (emails.length > 0) {
                                        const primaryEmail = emails[0].value;
                                        
                                        for (const nameObj of names) {
                                            const displayName = nameObj.displayName || '';
                                            const givenName = nameObj.givenName || '';
                                            const familyName = nameObj.familyName || '';
                                            
                                            // Store various name formats
                                            if (displayName) {
                                                contactLookup[displayName.toLowerCase()] = primaryEmail;
                                            }
                                            if (givenName && familyName) {
                                                contactLookup[`${givenName} ${familyName}`.toLowerCase()] = primaryEmail;
                                            }
                                            if (givenName) {
                                                contactLookup[givenName.toLowerCase()] = primaryEmail;
                                            }
                                        }
                                    }
                                }
                                
                                // Now match pending names
                                for (const pendingName of window._pendingPeopleApiLookups) {
                                    const lowerName = pendingName.toLowerCase();
                                    
                                    if (contactLookup[lowerName]) {
                                        const email = contactLookup[lowerName];
                                        emailMap[pendingName] = email;
                                        emailMap[`@${pendingName}`] = email;
                                        console.log(`✓ People API: "${pendingName}" → ${email}`);
                                    } else {
                                        // Try partial match
                                        for (const [contactName, email] of Object.entries(contactLookup)) {
                                            if (contactName.includes(lowerName) || lowerName.includes(contactName)) {
                                                emailMap[pendingName] = email;
                                                emailMap[`@${pendingName}`] = email;
                                                console.log(`✓ People API (partial): "${pendingName}" → ${email}`);
                                                break;
                                            }
                                        }
                                    }
                                }
                            } else {
                                const errorData = await peopleResponse.json().catch(() => ({}));
                                console.warn('People API request failed:', peopleResponse.status, errorData);
                                if (peopleResponse.status === 403) {
                                    console.warn('People API not enabled or permission denied. Enable it at: https://console.cloud.google.com/apis/library/people.googleapis.com');
                                }
                            }
                        } catch (peopleError) {
                            console.warn('People API lookup failed:', peopleError.message);
                        }
                        
                        // Clear the pending lookups
                        window._pendingPeopleApiLookups.clear();
                    }
                    
                    console.log(`Final count: ${Object.keys(emailMap).length} participant emails`);
                    return emailMap;
                }
            } catch (error) {
                console.error('Failed to extract emails from contact tags:', error);
                console.error('Error details:', {
                    message: error.message,
                    stack: error.stack,
                    spreadsheetId,
                    sheetName
                });
                // Don't fall back - we need to see what's wrong
            }
    }
    
    // Fallback: Try to extract from sheetData if we have a Name/Email column structure
    if (sheetData && Object.keys(emailMap).length === 0 && sheetData.length >= 2) {
        console.log('Trying fallback: looking for Name/Email columns in sheetData...');
        const headers = sheetData[0] || [];
        let nameColIndex = -1;
        let emailColIndex = -1;
        
        headers.forEach((header, index) => {
            const headerLower = (header || '').toLowerCase().trim();
            if (headerLower.includes('name') || headerLower.includes('participant') || headerLower.includes('person')) {
                nameColIndex = index;
            }
            if (headerLower.includes('email') || headerLower.includes('mail')) {
                emailColIndex = index;
            }
        });
        
        if (nameColIndex >= 0 && emailColIndex >= 0) {
            console.log(`Found Name column at index ${nameColIndex}, Email column at index ${emailColIndex}`);
            for (let rowIndex = 1; rowIndex < sheetData.length; rowIndex++) {
                const row = sheetData[rowIndex];
                if (!row || !Array.isArray(row)) continue;
                
                const name = (row[nameColIndex] || '').trim();
                const email = (row[emailColIndex] || '').trim();
                
                if (name && email && email.includes('@')) {
                    const nameWithoutAt = name.startsWith('@') ? name.substring(1) : name;
                    emailMap[name] = email;
                    emailMap[nameWithoutAt] = email;
                    if (name.startsWith('@')) {
                        emailMap[`@${nameWithoutAt}`] = email;
                    }
                }
            }
            console.log(`Fallback extracted ${Object.keys(emailMap).length} emails from Name/Email columns`);
        } else {
            console.warn('Could not extract emails. Options:');
            console.warn('1. Add an "Email" column next to names in your spreadsheet');
            console.warn('2. Enable People API and re-authenticate to extract from contact tags');
        }
    }
    
    console.log(`Total extracted: ${Object.keys(emailMap).length} participant emails`);
    return emailMap;
}

/**
 * Generates voting form configurations based on the current draw state
 * @param {Object} drawState - The draw state object containing groups
 * @param {Object} config - The config object containing group names
 * @param {Object} participantEmails - Optional mapping of participant names to emails
 * @returns {Object} Voting form configurations and participant assignments
 */
function generateVotingForms(drawState, config, participantEmails = {}) {
    if (!drawState || !drawState.groups || Object.keys(drawState.groups).length === 0) {
        throw new Error('Draw state is empty. Please complete the draw first.');
    }

    // Extract teams and their members
    const teams = {};
    const allParticipants = new Set();
    
    config.groupNames.forEach(groupName => {
        const entries = drawState.groups[groupName] || [];
        const teamMembers = entries.map(entry => {
            const name = typeof entry === 'string' ? entry : entry.entry;
            return name.trim();
        }).filter(name => name.length > 0);
        
        if (teamMembers.length > 0) {
            teams[groupName] = teamMembers;
            teamMembers.forEach(member => allParticipants.add(member));
        }
    });

    const teamNames = Object.keys(teams);
    
    if (teamNames.length === 0) {
        throw new Error('No teams found in draw results.');
    }

    // Strategy: Create one form per team
    // Each form excludes that team, so team members can vote for all other teams
    const forms = {};
    const participantAssignments = {};

    teamNames.forEach(teamName => {
        const teamMembers = teams[teamName];
        const excludedTeam = teamName;
        const votingOptions = teamNames.filter(t => t !== excludedTeam);

        // Get emails for team members (try multiple lookup variations)
        const teamEmails = [];
        const missingEmailsForTeam = [];
        
        for (const name of teamMembers) {
            let email = null;
            const variations = [
                name,
                name.trim(),
                name.toLowerCase(),
                name.toLowerCase().trim(),
                `@${name}`,
                `@${name.trim()}`,
                name.startsWith('@') ? name.substring(1) : name,
                (name.startsWith('@') ? name.substring(1) : name).toLowerCase()
            ];
            
            for (const variation of variations) {
                if (participantEmails[variation]) {
                    email = participantEmails[variation];
                    break;
                }
            }
            
            if (email) {
                teamEmails.push(email);
                console.log(`  ✓ ${name} → ${email}`);
            } else {
                missingEmailsForTeam.push(name);
                console.error(`  ✗ ${name} → NO EMAIL (tried ${variations.length} variations)`);
            }
        }
        
        if (missingEmailsForTeam.length > 0) {
            console.warn(`  Team ${teamName} missing ${missingEmailsForTeam.length} emails: ${missingEmailsForTeam.join(', ')}`);
        }
        
        // Create form configuration
        const formId = `form_${teamName.replace(/\s+/g, '_')}`;
        forms[formId] = {
            formId: formId,
            teamName: teamName,
            excludedTeam: excludedTeam,
            votingOptions: votingOptions,
            assignedParticipants: [...teamMembers],
            assignedEmails: teamEmails, // Store emails for this team
            formTitle: teamName,
            formDescription: '' // No description
        };

        // Assign form to team members
        teamMembers.forEach(participant => {
            if (!participantAssignments[participant]) {
                participantAssignments[participant] = [];
            }
            participantAssignments[participant].push(formId);
        });
    });

    // Add Judges form - includes ALL projects (no exclusions)
    const allProjects = [...teamNames]; // All teams/projects
    forms['form_Judges'] = {
        formId: 'form_Judges',
        teamName: 'Judges',
        excludedTeam: null, // No exclusions for judges
        votingOptions: allProjects,
        assignedParticipants: ['Judges'],
        assignedEmails: [],
        formTitle: 'Judges Voting',
        formDescription: '',
        isJudgesForm: true
    };

    // Add RoW (Rest of World) form - includes ALL projects (no exclusions), votes 1-5 like Judges
    forms['form_RoW'] = {
        formId: 'form_RoW',
        teamName: 'RoW',
        excludedTeam: null, // No exclusions for RoW
        votingOptions: allProjects,
        assignedParticipants: ['RoW'],
        assignedEmails: [],
        formTitle: 'RoW Voting',
        formDescription: '',
        isRoWForm: true
    };

    return {
        teams: teams,
        forms: forms,
        participantAssignments: participantAssignments,
        summary: {
            totalTeams: teamNames.length,
            totalParticipants: allParticipants.size,
            totalForms: Object.keys(forms).length
        }
    };
}

// ==================== GOOGLE FORMS API INTEGRATION ====================

/**
 * Gets the parent folder ID from a spreadsheet
 * @param {string} accessToken - OAuth access token
 * @param {string} spreadsheetId - The spreadsheet ID
 * @returns {Promise<string|null>} Parent folder ID or null if root
 */
async function getSpreadsheetParentFolder(accessToken, spreadsheetId) {
    try {
        const response = await fetch(`https://www.googleapis.com/drive/v3/files/${spreadsheetId}?fields=parents`, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });

        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(`Failed to get spreadsheet metadata: ${errorData.error?.message || response.statusText}`);
        }

        const data = await response.json();
        // Return the first parent folder ID, or null if in root
        return data.parents && data.parents.length > 0 ? data.parents[0] : null;
    } catch (error) {
        console.error('Error getting parent folder:', error);
        return null;
    }
}

/**
 * Creates a folder in Google Drive
 * @param {string} accessToken - OAuth access token
 * @param {string} folderName - Name of the folder
 * @param {string} parentFolderId - Parent folder ID (null for root)
 * @returns {Promise<string>} Created folder ID
 */
async function createDriveFolder(accessToken, folderName, parentFolderId = null) {
    const folderMetadata = {
        name: folderName,
        mimeType: 'application/vnd.google-apps.folder'
    };

    if (parentFolderId) {
        folderMetadata.parents = [parentFolderId];
    }

    const response = await fetch('https://www.googleapis.com/drive/v3/files', {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(folderMetadata)
    });

    if (!response.ok) {
        const errorData = await response.json();
        throw new Error(`Failed to create folder: ${errorData.error?.message || response.statusText}`);
    }

    const data = await response.json();
    return data.id;
}

/**
 * Moves a file to a folder in Google Drive
 * @param {string} accessToken - OAuth access token
 * @param {string} fileId - File ID to move
 * @param {string} targetFolderId - Target folder ID
 * @param {string} previousParentId - Previous parent folder ID (optional)
 */
async function moveFileToFolder(accessToken, fileId, targetFolderId, previousParentId = null) {
    // First, get current parents
    const getResponse = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}?fields=parents`, {
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${accessToken}`
        }
    });

    if (!getResponse.ok) {
        const errorData = await getResponse.json();
        throw new Error(`Failed to get file parents: ${errorData.error?.message || getResponse.statusText}`);
    }

    const fileData = await getResponse.json();
    const currentParents = fileData.parents || [];

    // Remove from old parent and add to new parent
    const removeParents = previousParentId ? [previousParentId] : currentParents;
    
    const response = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}?removeParents=${removeParents.join(',')}&addParents=${targetFolderId}`, {
        method: 'PATCH',
        headers: {
            'Authorization': `Bearer ${accessToken}`
        }
    });

    if (!response.ok) {
        const errorData = await response.json();
        throw new Error(`Failed to move file: ${errorData.error?.message || response.statusText}`);
    }
}

/**
 * Extracts spreadsheet ID from Google Sheets URL
 * @param {string} url - Google Sheets URL
 * @returns {string|null} Spreadsheet ID or null
 */
function extractSpreadsheetId(url) {
    if (!url) return null;
    const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : null;
}

/**
 * Extracts file ID from Google Drive URL
 * @param {string} url - Google Drive URL
 * @returns {string|null} File ID or null
 */
function extractFileIdFromDriveUrl(url) {
    if (!url) return null;

    // Format 1: https://drive.google.com/file/d/FILE_ID/view
    let match = url.match(/\/file\/d\/([a-zA-Z0-9-_]+)/);
    if (match) return match[1];

    // Format 2: https://drive.google.com/open?id=FILE_ID
    match = url.match(/[?&]id=([a-zA-Z0-9-_]+)/);
    if (match) return match[1];

    // Format 3: https://docs.google.com/spreadsheets/d/FILE_ID/edit
    match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (match) return match[1];

    // Format 4: Direct file ID
    if (/^[a-zA-Z0-9-_]+$/.test(url.trim())) {
        return url.trim();
    }

    return null;
}

/**
 * Gets the spreadsheet URL from config or localStorage
 * @returns {string|null} Spreadsheet URL or null
 */
function getSpreadsheetUrl() {
    // Try to get from various sources
    const googleSheetUrl = document.getElementById('googleSheetUrl')?.value.trim();
    const googleDriveFileUrl = document.getElementById('googleDriveFileUrl')?.value.trim();
    const localStorageUrl = localStorage.getItem('lastGoogleSheetUrl') || '';
    const configUrl = (typeof APP_CONFIG !== 'undefined' ? APP_CONFIG.googleSheets?.lastSheetUrl || '' : '');

    return googleSheetUrl || googleDriveFileUrl || localStorageUrl || configUrl || null;
}

/**
 * Creates a Google Form using the Google Forms API
 * @param {string} accessToken - OAuth access token
 * @param {Object} formConfig - Form configuration object
 * @param {string} targetFolderId - Target folder ID to place the form in (optional)
 * @returns {Promise<Object>} Created form data with formId and responderUri
 */
async function createGoogleForm(accessToken, formConfig, targetFolderId = null) {
    if (!accessToken) {
        throw new Error('Access token required. Please authenticate with Google first.');
    }

    // Step 1: Create the form (only title can be set during creation)
    const formTitle = formConfig.formTitle || formConfig.teamName || 'Untitled Form';
    console.log(`Creating form with title: "${formTitle}"`);
    
    const createResponse = await fetch('https://forms.googleapis.com/v1/forms', {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            info: {
                title: formTitle,
                documentTitle: formTitle
            }
        })
    });

    if (!createResponse.ok) {
        const errorData = await createResponse.json();
        throw new Error(`Failed to create form: ${errorData.error?.message || createResponse.statusText}`);
    }

    const formData = await createResponse.json();
    const formId = formData.formId;
    console.log(`Form created with ID: ${formId}`);
    console.log(`  Title sent: "${formTitle}"`);
    console.log(`  Title in response: "${formData.info?.title || 'NOT SET'}"`);
    console.log(`  Full response:`, formData);

    // Step 2: Add questions to the form via batchUpdate
    const requests = [];
    
    const numProjects = formConfig.votingOptions.length;
    
    // Update form info (no description)
    requests.push({
        updateFormInfo: {
            info: {
                title: formTitle,
                description: ''
            },
            updateMask: 'description,title'
        }
    });
    
    // Note: Google Forms API doesn't support setting collectEmail, requiresLogin, etc. via API
    // These must be configured manually in the form settings
    
    // Create rank columns based on number of projects
    const rankColumns = [];
    for (let i = 1; i <= numProjects; i++) {
        const points = numProjects - i + 1;
        rankColumns.push({ value: `${i}${i === 1 ? 'st' : i === 2 ? 'nd' : i === 3 ? 'rd' : 'th'} (${points} pts)` });
    }
    
    // Define the three categories
    const categories = ['Business Impact', 'Production Readiness', 'Presentation'];
    
    // Add each category as a grid question (rows=projects, columns=rankings)
    if (formConfig.votingOptions.length > 0) {
        let itemIndex = 0;
        
        for (const category of categories) {
            // Add grid question for ranking
            requests.push({
                createItem: {
                    item: {
                        title: category,
                        questionGroupItem: {
                            questions: formConfig.votingOptions.map(project => ({
                                rowQuestion: {
                                    title: project
                                }
                            })),
                            grid: {
                                columns: {
                                    type: 'RADIO',
                                    options: rankColumns
                                },
                                shuffleQuestions: false
                            }
                        }
                    },
                    location: { index: itemIndex++ }
                }
            });
        }
    }

    // Batch update to add questions (settings removed - must be done manually)
    if (requests.length > 0) {
        const updateResponse = await fetch(`https://forms.googleapis.com/v1/forms/${formId}:batchUpdate`, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                requests: requests
            })
        });

        if (!updateResponse.ok) {
            const errorData = await updateResponse.json();
            console.error(`Failed to add questions to form: ${errorData.error?.message || updateResponse.statusText}`);
            throw new Error(`Failed to configure form: ${errorData.error?.message || updateResponse.statusText}`);
        } else {
            console.log('✓ Form questions added successfully');
        }
    }

    // Step 3: Rename the form file in Google Drive (Forms API doesn't set documentTitle correctly)
    console.log(`Renaming form file in Drive to: "${formTitle}"`);
    try {
        const renameResponse = await fetch(`https://www.googleapis.com/drive/v3/files/${formId}`, {
            method: 'PATCH',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                name: formTitle
            })
        });
        
        if (!renameResponse.ok) {
            const errorBody = await renameResponse.json().catch(() => ({}));
            console.error(`RENAME FAILED - Status: ${renameResponse.status}, Error:`, errorBody);
        } else {
            const renameData = await renameResponse.json();
            console.log(`✓ Form renamed successfully to: "${renameData.name}"`);
        }
    } catch (renameError) {
        console.error(`RENAME EXCEPTION:`, renameError);
    }

    // Step 4: Get the form details including responder URI
    const getResponse = await fetch(`https://forms.googleapis.com/v1/forms/${formId}`, {
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${accessToken}`
        }
    });

    if (!getResponse.ok) {
        // If we can't get details, construct the URL manually
        return {
            formId: formId,
            responderUri: `https://docs.google.com/forms/d/${formId}/viewform`,
            editUri: `https://docs.google.com/forms/d/${formId}/edit`
        };
    }

    const formDetails = await getResponse.json();
    
    // Move form to target folder if specified
    if (targetFolderId) {
        try {
            console.log(`Moving form ${formId} to folder ${targetFolderId}...`);
            await moveFileToFolder(accessToken, formId, targetFolderId);
            console.log(`✓ Successfully moved form ${formId} to folder ${targetFolderId}`);
        } catch (error) {
            console.error(`✗ Failed to move form ${formId} to folder ${targetFolderId}:`, error.message);
        }
    }
    
    // Note: Response spreadsheet linking via API doesn't work reliably.
    // Responses are read directly from Forms API via aggregateFormResponses().
    
    return {
        formId: formId,
        responderUri: formDetails.responderUri || `https://docs.google.com/forms/d/${formId}/viewform`,
        editUri: `https://docs.google.com/forms/d/${formId}/edit`,
        title: formDetails.info?.title || formConfig.formTitle,
        folderId: targetFolderId
    };
}

/**
 * Restricts form access to specific email addresses
 * Note: Google Forms API doesn't support direct email whitelisting.
 * This function shares the form with specific users via Drive API.
 * You may need to manually disable "Anyone with the link" in form settings.
 * @param {string} accessToken - OAuth access token
 * @param {string} formId - The form ID
 * @param {Array<string>} allowedEmails - Array of email addresses allowed to access the form
 */
async function restrictFormToEmails(accessToken, formId, allowedEmails) {
    if (!allowedEmails || allowedEmails.length === 0) {
        console.warn('No emails provided for restriction');
        return;
    }
    
    // Use Drive API to share the form with specific users
    // First, get current permissions
    const permissionsResponse = await fetch(`https://www.googleapis.com/drive/v3/files/${formId}/permissions?fields=permissions(id,emailAddress,role)`, {
        headers: { 'Authorization': `Bearer ${accessToken}` }
    });
    
    if (permissionsResponse.ok) {
        const permissions = await permissionsResponse.json();
        // Remove "Anyone" permission if it exists
        for (const perm of permissions.permissions || []) {
            if (perm.role === 'reader' && !perm.emailAddress) {
                // This is "Anyone with the link" - remove it
                await fetch(`https://www.googleapis.com/drive/v3/files/${formId}/permissions/${perm.id}`, {
                    method: 'DELETE',
                    headers: { 'Authorization': `Bearer ${accessToken}` }
                });
            }
        }
    }
    
    // Add permissions for each allowed email
    for (const email of allowedEmails) {
        try {
            await fetch(`https://www.googleapis.com/drive/v3/files/${formId}/permissions`, {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    role: 'reader',
                    type: 'user',
                    emailAddress: email
                })
            });
        } catch (error) {
            console.error(`Failed to share form with ${email}:`, error);
        }
    }
    
    console.log(`✓ Shared form ${formId} with ${allowedEmails.length} email addresses`);
    console.warn('⚠️ IMPORTANT: You must manually disable "Anyone with the link" in form settings for full restriction.');
}

/**
 * Ensures required sheets exist in a spreadsheet. Creates any missing sheets.
 * @param {string} accessToken - OAuth access token
 * @param {string} spreadsheetId - The spreadsheet ID
 * @param {Array<string>} requiredSheets - Array of sheet names that must exist
 */
async function ensureSheetsExist(accessToken, spreadsheetId, requiredSheets) {
    try {
        // Get current sheets in the spreadsheet
        const metadataResponse = await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets.properties.title`,
            { headers: { 'Authorization': `Bearer ${accessToken}` } }
        );
        
        if (!metadataResponse.ok) {
            console.warn('Could not get spreadsheet metadata, will try to create sheets anyway');
            return;
        }
        
        const metadata = await metadataResponse.json();
        const existingSheets = (metadata.sheets || []).map(s => s.properties.title);
        console.log(`Existing sheets: ${existingSheets.join(', ')}`);
        
        // Find which sheets need to be created
        const sheetsToCreate = requiredSheets.filter(name => !existingSheets.includes(name));
        
        if (sheetsToCreate.length === 0) {
            console.log('All required sheets already exist');
            return;
        }
        
        console.log(`Creating missing sheets: ${sheetsToCreate.join(', ')}`);
        
        // Create missing sheets using batchUpdate
        const requests = sheetsToCreate.map(sheetName => ({
            addSheet: {
                properties: {
                    title: sheetName,
                    gridProperties: { rowCount: 1000, columnCount: 20 }
                }
            }
        }));
        
        const batchResponse = await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`,
            {
                method: 'POST',
                headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({ requests })
            }
        );
        
        if (batchResponse.ok) {
            console.log(`✓ Created ${sheetsToCreate.length} missing sheets`);
        } else {
            const error = await batchResponse.json();
            console.warn('Failed to create some sheets:', error);
        }
    } catch (error) {
        console.warn('Error ensuring sheets exist (will continue anyway):', error.message);
    }
}

/**
 * Safely clears a sheet range, ignoring errors if sheet doesn't exist
 * @param {string} accessToken - OAuth access token
 * @param {string} spreadsheetId - The spreadsheet ID
 * @param {string} range - The range to clear (e.g., 'Sheet1!A2:Z')
 */
async function safeClearRange(accessToken, spreadsheetId, range) {
    try {
        const response = await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}:clear`,
            {
                method: 'POST',
                headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' }
            }
        );
        if (!response.ok) {
            // Silently ignore - sheet might not exist
            console.log(`Note: Could not clear ${range.split('!')[0]} (may not exist yet)`);
        }
    } catch (error) {
        // Silently ignore
    }
}

/**
 * Safely writes to a sheet range, ignoring errors if sheet doesn't exist
 * @param {string} accessToken - OAuth access token
 * @param {string} spreadsheetId - The spreadsheet ID
 * @param {string} range - The range to write to
 * @param {Array} values - The values to write
 * @param {string} method - 'PUT' or 'POST' (for append)
 */
async function safeWriteRange(accessToken, spreadsheetId, range, values, method = 'PUT') {
    if (!values || values.length === 0) {
        console.log(`>>> safeWriteRange: No data to write for ${range}`);
        return false;
    }
    
    console.log(`>>> safeWriteRange: Writing ${values.length} rows to ${range} (method: ${method})`);
    
    try {
        const url = method === 'POST' 
            ? `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`
            : `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}?valueInputOption=RAW`;
        
        console.log(`>>> URL: ${url}`);
        
        const response = await fetch(url, {
            method: method,
            headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ values })
        });
        
        if (!response.ok) {
            const error = await response.json();
            console.error(`ERROR writing to ${range}:`, error.error?.message || 'unknown error');
            console.error('Full error:', error);
            return false;
        }
        
        console.log(`>>> SUCCESS: Wrote to ${range}`);
        return true;
    } catch (error) {
        console.error(`ERROR writing to ${range}:`, error.message);
        return false;
    }
}

// ==================== SHARED RESULTS PROCESSING ====================
// Constants for weighting
const CATEGORY_WEIGHTS = { impact: 0.4, readiness: 0.4, presentation: 0.2 };
const GROUP_WEIGHTS = { participants: 0.4, row: 0.2, judges: 0.4 };
const SCALE_FACTOR = 4 / 5; // = 0.8 (RoW & Judges use 1-5 scale, Participants use 1-4)

/**
 * Calculate weighted results from scores
 * @param {Object} scores - { project: { impact: [], readiness: [], presentation: [] } }
 * @returns {Array} Sorted array of { project, impact, readiness, presentation, total }
 */
function calculateWeightedResults(scores) {
    const results = [];
    for (const [project, s] of Object.entries(scores)) {
        const avgImpact = s.impact.length > 0 ? s.impact.reduce((a, b) => a + b, 0) / s.impact.length : 0;
        const avgReadiness = s.readiness.length > 0 ? s.readiness.reduce((a, b) => a + b, 0) / s.readiness.length : 0;
        const avgPresentation = s.presentation.length > 0 ? s.presentation.reduce((a, b) => a + b, 0) / s.presentation.length : 0;
        const total = avgImpact * CATEGORY_WEIGHTS.impact + avgReadiness * CATEGORY_WEIGHTS.readiness + avgPresentation * CATEGORY_WEIGHTS.presentation;
        results.push({
            project,
            impact: avgImpact * CATEGORY_WEIGHTS.impact,
            readiness: avgReadiness * CATEGORY_WEIGHTS.readiness,
            presentation: avgPresentation * CATEGORY_WEIGHTS.presentation,
            total
        });
    }
    results.sort((a, b) => b.total - a.total);
    return results;
}

/**
 * Calculate final combined results from all three groups
 * @param {Array} participantWeighted - Participant weighted results
 * @param {Array} rowWeighted - RoW weighted results
 * @param {Array} judgesWeighted - Judges weighted results
 * @returns {Array} Final results as [[project, pContrib, rContrib, jContrib, finalScore], ...]
 */
function calculateFinalResults(participantWeighted, rowWeighted, judgesWeighted) {
    const allProjects = new Set([
        ...participantWeighted.map(r => r.project),
        ...rowWeighted.map(r => r.project),
        ...judgesWeighted.map(r => r.project)
    ]);
    
    const finalResults = [];
    for (const project of allProjects) {
        const pScore = participantWeighted.find(r => r.project === project)?.total || 0;
        const rScore = rowWeighted.find(r => r.project === project)?.total || 0;
        const jScore = judgesWeighted.find(r => r.project === project)?.total || 0;
        
        // Scale RoW and Judges (they use 1-5 scale, participants use 1-4)
        const rScoreScaled = rScore * SCALE_FACTOR;
        const jScoreScaled = jScore * SCALE_FACTOR;
        
        // Apply group weights: Participants 40%, RoW 20%, Judges 40%
        const pContrib = pScore * GROUP_WEIGHTS.participants;
        const rContrib = rScoreScaled * GROUP_WEIGHTS.row;
        const jContrib = jScoreScaled * GROUP_WEIGHTS.judges;
        const finalScore = pContrib + rContrib + jContrib;
        
        finalResults.push([
            project,
            Math.round(pContrib * 100) / 100,
            Math.round(rContrib * 100) / 100,
            Math.round(jContrib * 100) / 100,
            Math.round(finalScore * 100) / 100
        ]);
    }
    
    finalResults.sort((a, b) => parseFloat(b[4]) - parseFloat(a[4]));
    return finalResults;
}

/**
 * SHARED: Process collected data and write all results to spreadsheet
 * Used by both generateTestDataDirect and aggregateFormResponses
 * 
 * @param {string} accessToken - OAuth access token
 * @param {string} spreadsheetId - Results spreadsheet ID
 * @param {Object} data - Collected data with votes and scores for each group
 * @param {Array} data.participantVotes - [[timestamp, email, form, category, project, rank, points], ...]
 * @param {Array} data.rowVotes - Same format
 * @param {Array} data.judgesVotes - Same format
 * @param {Object} data.participantScores - { project: { impact: [], readiness: [], presentation: [] } }
 * @param {Object} data.rowScores - Same format
 * @param {Object} data.judgesScores - Same format
 * @returns {Promise<Object>} Summary of what was written
 */
async function processAndWriteResults(accessToken, spreadsheetId, data) {
    console.log('=== PROCESSING AND WRITING RESULTS ===');
    console.log(`Spreadsheet: https://docs.google.com/spreadsheets/d/${spreadsheetId}`);
    
    const { participantVotes, rowVotes, judgesVotes, participantScores, rowScores, judgesScores } = data;
    
    // 1. Ensure all required sheets exist
    const requiredSheets = [
        'Participants Votes', 'RoW Votes', 'Judges Votes',
        'Participants Weighted Results', 'RoW Weighted Results', 'Judges Weighted Results',
        'Final Weighted Results'
    ];
    console.log(`Ensuring sheets exist: ${requiredSheets.join(', ')}`);
    await ensureSheetsExist(accessToken, spreadsheetId, requiredSheets);
    
    // 2. Clear ALL existing data (start from scratch)
    console.log('Clearing all existing data...');
    const sheetsToClear = [
        'Participants Votes!A:Z',
        'RoW Votes!A:Z',
        'Judges Votes!A:Z',
        'Participants Weighted Results!A:Z',
        'RoW Weighted Results!A:Z',
        'Judges Weighted Results!A:Z',
        'Final Weighted Results!A:Z'
    ];
    for (const range of sheetsToClear) {
        await safeClearRange(accessToken, spreadsheetId, range);
    }
    console.log('✓ All sheets cleared');
    
    // 3. Write headers to all sheets
    const votesHeaders = [['Timestamp', 'Email', 'Form', 'Category', 'Project', 'Rank', 'Points']];
    const weightedHeaders = [['Project', 'Business Impact (40%)', 'Production Readiness (40%)', 'Presentation (20%)', 'Total Score']];
    const finalHeaders = [['Project', 'Participants (40%)', 'RoW (20%, scaled 0.8)', 'Judges (40%, scaled 0.8)', 'Final Score']];
    
    console.log('Writing headers...');
    await safeWriteRange(accessToken, spreadsheetId, 'Participants Votes!A1:G1', votesHeaders, 'PUT');
    await safeWriteRange(accessToken, spreadsheetId, 'RoW Votes!A1:G1', votesHeaders, 'PUT');
    await safeWriteRange(accessToken, spreadsheetId, 'Judges Votes!A1:G1', votesHeaders, 'PUT');
    await safeWriteRange(accessToken, spreadsheetId, 'Participants Weighted Results!A1:E1', weightedHeaders, 'PUT');
    await safeWriteRange(accessToken, spreadsheetId, 'RoW Weighted Results!A1:E1', weightedHeaders, 'PUT');
    await safeWriteRange(accessToken, spreadsheetId, 'Judges Weighted Results!A1:E1', weightedHeaders, 'PUT');
    await safeWriteRange(accessToken, spreadsheetId, 'Final Weighted Results!A1:E1', finalHeaders, 'PUT');
    console.log('✓ Headers written');
    
    // 4. Write raw votes
    console.log(`Writing votes: ${participantVotes.length} participant, ${rowVotes.length} RoW, ${judgesVotes.length} judges`);
    if (participantVotes.length > 0) {
        await safeWriteRange(accessToken, spreadsheetId, 'Participants Votes!A2:G', participantVotes, 'PUT');
        console.log(`✓ Wrote ${participantVotes.length} participant votes`);
    }
    if (rowVotes.length > 0) {
        await safeWriteRange(accessToken, spreadsheetId, 'RoW Votes!A2:G', rowVotes, 'PUT');
        console.log(`✓ Wrote ${rowVotes.length} RoW votes`);
    }
    if (judgesVotes.length > 0) {
        await safeWriteRange(accessToken, spreadsheetId, 'Judges Votes!A2:G', judgesVotes, 'PUT');
        console.log(`✓ Wrote ${judgesVotes.length} judges votes`);
    }
    
    // 5. Calculate weighted results for each group
    const participantWeighted = calculateWeightedResults(participantScores);
    const rowWeighted = calculateWeightedResults(rowScores);
    const judgesWeighted = calculateWeightedResults(judgesScores);
    
    // 6. Write weighted results (keep as numbers, round to 2 decimals)
    const formatWeighted = (results) => results.map(r => [
        r.project,
        Math.round(r.impact * 100) / 100,
        Math.round(r.readiness * 100) / 100,
        Math.round(r.presentation * 100) / 100,
        Math.round(r.total * 100) / 100
    ]);
    
    if (participantWeighted.length > 0) {
        await safeWriteRange(accessToken, spreadsheetId, 'Participants Weighted Results!A2:E', formatWeighted(participantWeighted), 'PUT');
        console.log(`✓ Wrote ${participantWeighted.length} participant weighted results`);
    }
    if (rowWeighted.length > 0) {
        await safeWriteRange(accessToken, spreadsheetId, 'RoW Weighted Results!A2:E', formatWeighted(rowWeighted), 'PUT');
        console.log(`✓ Wrote ${rowWeighted.length} RoW weighted results`);
    }
    if (judgesWeighted.length > 0) {
        await safeWriteRange(accessToken, spreadsheetId, 'Judges Weighted Results!A2:E', formatWeighted(judgesWeighted), 'PUT');
        console.log(`✓ Wrote ${judgesWeighted.length} judges weighted results`);
    }
    
    // 7. Calculate and write final combined results
    const finalResults = calculateFinalResults(participantWeighted, rowWeighted, judgesWeighted);
    if (finalResults.length > 0) {
        await safeWriteRange(accessToken, spreadsheetId, 'Final Weighted Results!A2:E', finalResults, 'PUT');
        console.log(`✓ Wrote ${finalResults.length} final weighted results`);
    }
    
    console.log('=== RESULTS PROCESSING COMPLETE ===');
    
    return {
        participantVotes: participantVotes.length,
        rowVotes: rowVotes.length,
        judgesVotes: judgesVotes.length,
        projects: finalResults.length,
        spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}`
    };
}

/**
 * Creates a results spreadsheet with raw votes and weighted scores
 * @param {string} accessToken - OAuth access token
 * @param {Object} votingData - Voting data from generateVotingForms()
 * @param {Object} createdForms - Created forms data
 * @returns {Promise<Object>} Created spreadsheet with URLs
 */
async function createResultsSpreadsheet(accessToken, votingData, createdForms) {
    // Create a new spreadsheet
    const spreadsheetResponse = await fetch('https://sheets.googleapis.com/v4/spreadsheets', {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            properties: {
                title: `Voting Results - ${new Date().toISOString().split('T')[0]}`
            },
            sheets: [
                {
                    properties: {
                        title: 'Participants Votes',
                        gridProperties: { rowCount: 1000, columnCount: 20 }
                    }
                },
                {
                    properties: {
                        title: 'RoW Votes',
                        gridProperties: { rowCount: 1000, columnCount: 20 }
                    }
                },
                {
                    properties: {
                        title: 'Judges Votes',
                        gridProperties: { rowCount: 1000, columnCount: 20 }
                    }
                },
                {
                    properties: {
                        title: 'Participants Weighted Results',
                        gridProperties: { rowCount: 1000, columnCount: 20 }
                    }
                },
                {
                    properties: {
                        title: 'RoW Weighted Results',
                        gridProperties: { rowCount: 1000, columnCount: 20 }
                    }
                },
                {
                    properties: {
                        title: 'Judges Weighted Results',
                        gridProperties: { rowCount: 1000, columnCount: 20 }
                    }
                },
                {
                    properties: {
                        title: 'Final Weighted Results',
                        gridProperties: { rowCount: 1000, columnCount: 20 }
                    }
                }
            ]
        })
    });
    
    if (!spreadsheetResponse.ok) {
        const errorData = await spreadsheetResponse.json();
        throw new Error(`Failed to create spreadsheet: ${errorData.error?.message || spreadsheetResponse.statusText}`);
    }
    
    const spreadsheet = await spreadsheetResponse.json();
    const spreadsheetId = spreadsheet.spreadsheetId;
    
    // Set up headers
    const votesHeaders = [['Timestamp', 'Email', 'Form', 'Category', 'Project', 'Rank', 'Points']];
    const weightedHeaders = [['Project', 'Business Impact (40%)', 'Production Readiness (40%)', 'Presentation (20%)', 'Total Score']];
    const finalHeaders = [['Project', 'Participants (40%)', 'RoW (20%, scaled 0.8)', 'Judges (40%, scaled 0.8)', 'Final Score']];
    
    // Write headers to all sheets
    const headerWrites = [
        { range: 'Participants Votes!A1:G1', values: votesHeaders },
        { range: 'RoW Votes!A1:G1', values: votesHeaders },
        { range: 'Judges Votes!A1:G1', values: votesHeaders },
        { range: 'Participants Weighted Results!A1:E1', values: weightedHeaders },
        { range: 'RoW Weighted Results!A1:E1', values: weightedHeaders },
        { range: 'Judges Weighted Results!A1:E1', values: weightedHeaders },
        { range: 'Final Weighted Results!A1:E1', values: finalHeaders }
    ];
    
    for (const hw of headerWrites) {
        await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(hw.range)}?valueInputOption=RAW`, {
            method: 'PUT',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ values: hw.values })
        });
    }
    
    // Store spreadsheet ID in createdForms for later use
    if (createdForms) {
        createdForms.resultsSpreadsheetId = spreadsheetId;
        createdForms.resultsSpreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;
    }
    
    return {
        spreadsheetId: spreadsheetId,
        spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
        rawVotesSheetId: spreadsheet.sheets[0].properties.sheetId,
        weightedResultsSheetId: spreadsheet.sheets[1].properties.sheetId
    };
}

/**
 * Aggregates responses from all forms and calculates weighted scores
 * @param {string} accessToken - OAuth access token
 * @param {Object} createdForms - Created forms data with formIds
 * @param {Object} resultsSpreadsheet - Results spreadsheet data
 */
async function aggregateFormResponses(accessToken, createdForms, resultsSpreadsheet) {
    const allResponses = [];
    const projectScores = {}; // { project: { impact: [], readiness: [], presentation: [] } }
    
    console.log('=== aggregateFormResponses START ===');
    console.log('Number of forms:', Object.keys(createdForms.forms || {}).length);
    
    // Get responses from each form
    for (const [formId, formData] of Object.entries(createdForms.forms || {})) {
        console.log(`Checking form: ${formData.teamName}, formId: ${formData.formId}, status: ${formData.status}`);
        
        if (!formData.formId || formData.status !== 'created') {
            console.log(`  Skipping - no formId or status not created`);
            continue;
        }
        
        try {
            // Get form responses
            console.log(`  Fetching responses from Forms API for ${formData.teamName}...`);
            const responsesResponse = await fetch(`https://forms.googleapis.com/v1/forms/${formData.formId}/responses`, {
                headers: { 'Authorization': `Bearer ${accessToken}` }
            });
            
            if (!responsesResponse.ok) {
                const errText = await responsesResponse.text();
                console.log(`  Failed to get responses: ${responsesResponse.status} - ${errText}`);
                continue;
            }
            
            const responsesData = await responsesResponse.json();
            console.log(`  Got ${(responsesData.responses || []).length} responses`);
            console.log('  Raw responses data:', JSON.stringify(responsesData, null, 2));
            
            // Get form structure ONCE (not per response)
            const formDetailsResponse = await fetch(`https://forms.googleapis.com/v1/forms/${formData.formId}`, {
                headers: { 'Authorization': `Bearer ${accessToken}` }
            });
            
            if (!formDetailsResponse.ok) {
                console.log('  Failed to get form details');
                continue;
            }
            const formDetails = await formDetailsResponse.json();
            console.log('  Form items:', formDetails.items?.length || 0);
            
            // Map questionGroupItem IDs to categories
            // For grid questions, the questionGroupItem has an itemId, and the grid answer references this
            const itemToCategory = {};
            const items = formDetails.items || [];
            const categories = ['Business Impact', 'Production Readiness', 'Presentation'];
            
            items.forEach((item, idx) => {
                if (item.questionGroupItem && idx < categories.length) {
                    // Get all question IDs in this group
                    const questions = item.questionGroupItem.questions || [];
                    questions.forEach(q => {
                        if (q.questionId) {
                            itemToCategory[q.questionId] = categories[idx];
                            console.log(`  Mapped questionId ${q.questionId} -> ${categories[idx]}`);
                        }
                    });
                }
            });
            
            // Process each response
            for (const response of responsesData.responses || []) {
                const email = response.respondentEmail || '';
                const timestamp = response.createTime || '';
                
                // Extract answers - answers is an object keyed by questionId
                const answers = response.answers || {};
                console.log('  Processing response, answer keys:', Object.keys(answers));
                
                // Process answers
                // Format: rows=projects, columns=rankings
                for (const [questionId, answerData] of Object.entries(answers)) {
                    const category = itemToCategory[questionId] || '';
                    console.log(`    questionId: ${questionId}, category: ${category}`);
                    
                    console.log(`    answerData keys:`, Object.keys(answerData));
                    
                    if (answerData.textAnswers) {
                        // Grid row answer - textAnswers contains the selected column
                        const textAnswers = answerData.textAnswers?.answers || [];
                        console.log(`    textAnswers:`, textAnswers);
                        
                        // For grid questions, we need to get the row title from the form structure
                        // Find which row this questionId corresponds to
                        let project = '';
                        for (const item of items) {
                            if (item.questionGroupItem) {
                                const questions = item.questionGroupItem.questions || [];
                                const q = questions.find(q => q.questionId === questionId);
                                if (q) {
                                    project = q.rowQuestion?.title || '';
                                    break;
                                }
                            }
                        }
                        
                        const columnValue = textAnswers[0]?.value || '';
                        console.log(`    project: ${project}, columnValue: ${columnValue}`);
                        
                        if (project && columnValue) {
                            // Extract rank AND points from column value (e.g., "1st (4 pts)" -> rank=1, points=4)
                            const rankMatch = columnValue.match(/^(\d+)/);
                            const pointsMatch = columnValue.match(/\((\d+)\s*pts?\)/i);
                            const rank = rankMatch ? parseInt(rankMatch[1]) : 0;
                            const points = pointsMatch ? parseInt(pointsMatch[1]) : 0;
                            console.log(`    Parsed: "${columnValue}" -> rank=${rank}, points=${points}`);
                            
                            // Store raw vote
                            allResponses.push({
                                timestamp,
                                email,
                                form: formData.teamName,
                                category,
                                project,
                                rank,
                                points
                            });
                            
                            // Accumulate scores
                            if (!projectScores[project]) {
                                projectScores[project] = {
                                    impact: [],
                                    readiness: [],
                                    presentation: []
                                };
                            }
                            
                            if (category === 'Business Impact') {
                                projectScores[project].impact.push(points);
                            } else if (category === 'Production Readiness') {
                                projectScores[project].readiness.push(points);
                            } else if (category === 'Presentation') {
                                projectScores[project].presentation.push(points);
                            }
                        }
                    } else if (answerData.questionGroupItemResponse) {
                        // Grid question response (old format)
                        const rowAnswers = answerData.questionGroupItemResponse.answers || [];
                        
                        for (const rowAnswer of rowAnswers) {
                            const project = rowAnswer.rowQuestion || '';
                            const value = rowAnswer.value || {};
                            const columnValue = value.choiceValue || '';
                            
                            // Extract rank AND points from column value (e.g., "1st (4 pts)" -> rank=1, points=4)
                            const rankMatch = columnValue.match(/^(\d+)/);
                            const pointsMatch = columnValue.match(/\((\d+)\s*pts?\)/i);
                            const rank = rankMatch ? parseInt(rankMatch[1]) : 0;
                            const points = pointsMatch ? parseInt(pointsMatch[1]) : 0;
                            
                            // Store raw vote
                            allResponses.push({
                                timestamp,
                                email,
                                form: formData.teamName,
                                category,
                                project,
                                rank,
                                points
                            });
                            
                            // Accumulate scores
                            if (!projectScores[project]) {
                                projectScores[project] = {
                                    impact: [],
                                    readiness: [],
                                    presentation: []
                                };
                            }
                            
                            if (category === 'Business Impact') {
                                projectScores[project].impact.push(points);
                            } else if (category === 'Production Readiness') {
                                projectScores[project].readiness.push(points);
                            } else if (category === 'Presentation') {
                                projectScores[project].presentation.push(points);
                            }
                        }
                    }
                }
            }
        } catch (error) {
            console.error(`Error fetching responses for form ${formData.formId}:`, error);
        }
    }
    
    // Separate participant, RoW, and judge responses and build scores
    const collectedData = {
        participantVotes: [],
        rowVotes: [],
        judgesVotes: [],
        participantScores: {},
        rowScores: {},
        judgesScores: {}
    };
    
    for (const r of allResponses) {
        const isJudge = r.form === 'Judges Voting' || r.form.toLowerCase().includes('judge');
        const isRoW = r.form === 'RoW Voting' || r.form.toLowerCase().includes('row');
        
        // Convert to array format for votes
        const voteRow = [r.timestamp, r.email, r.form, r.category, r.project, r.rank, r.points];
        
        // Determine which group this response belongs to
        let targetVotes, targetScores;
        if (isJudge) {
            targetVotes = collectedData.judgesVotes;
            targetScores = collectedData.judgesScores;
        } else if (isRoW) {
            targetVotes = collectedData.rowVotes;
            targetScores = collectedData.rowScores;
        } else {
            targetVotes = collectedData.participantVotes;
            targetScores = collectedData.participantScores;
        }
        
        targetVotes.push(voteRow);
        
        // Build scores
        if (!targetScores[r.project]) {
            targetScores[r.project] = { impact: [], readiness: [], presentation: [] };
        }
        if (r.category === 'Business Impact') targetScores[r.project].impact.push(r.points);
        else if (r.category === 'Production Readiness') targetScores[r.project].readiness.push(r.points);
        else if (r.category === 'Presentation') targetScores[r.project].presentation.push(r.points);
    }
    
    console.log(`Participant responses: ${collectedData.participantVotes.length}, RoW responses: ${collectedData.rowVotes.length}, Judge responses: ${collectedData.judgesVotes.length}`);
    
    // Use shared function to process and write all results
    const result = await processAndWriteResults(accessToken, resultsSpreadsheet.spreadsheetId, collectedData);
    result.totalResponses = allResponses.length;
    
    console.log('=== aggregateFormResponses COMPLETE ===');
    return result;
}

/**
 * Creates all voting forms using Google Forms API
 * @param {string} accessToken - OAuth access token
 * @param {Object} votingData - Voting data from generateVotingForms() (includes emails in form configs)
 * @returns {Promise<Object>} Created forms with URLs
 */
async function createAllVotingForms(accessToken, votingData) {
    const createdForms = {};
    const errors = [];
    let targetFolderId = null;
    let folderName = null;

    // Step 1: Get spreadsheet URL and create timestamped folder
    try {
        const spreadsheetUrl = getSpreadsheetUrl();
        console.log('=== FOLDER CREATION DEBUG ===');
        console.log('1. Spreadsheet URL found:', spreadsheetUrl);
        
        if (!spreadsheetUrl) {
            throw new Error('No spreadsheet URL found. Please make sure you have loaded data from a Google Sheet.');
        }
        
        // Extract spreadsheet ID
        const spreadsheetId = extractSpreadsheetId(spreadsheetUrl) || extractFileIdFromDriveUrl(spreadsheetUrl);
        console.log('2. Extracted spreadsheet ID:', spreadsheetId);
        
        if (!spreadsheetId) {
            throw new Error(`Could not extract spreadsheet ID from URL: ${spreadsheetUrl}`);
        }
        
        // Get parent folder
        console.log('3. Getting parent folder for spreadsheet...');
        const parentFolderId = await getSpreadsheetParentFolder(accessToken, spreadsheetId);
        console.log('4. Parent folder ID:', parentFolderId || 'root (no parent)');
        
        // Create timestamped folder
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-').split('T')[0] + '_' + 
                         new Date().toTimeString().split(' ')[0].replace(/:/g, '-').split('.')[0];
        folderName = `Voting Forms ${timestamp}`;
        
        console.log(`5. Creating folder: "${folderName}"`);
        console.log(`   Parent folder: ${parentFolderId || 'root'}`);
        
        targetFolderId = await createDriveFolder(accessToken, folderName, parentFolderId);
        
        console.log(`6. ✓ SUCCESS! Created folder with ID: ${targetFolderId}`);
        console.log(`   Folder URL: https://drive.google.com/drive/folders/${targetFolderId}`);
        console.log('=== END FOLDER CREATION DEBUG ===');
    } catch (error) {
        console.error('=== FOLDER CREATION FAILED ===');
        console.error('Error:', error.message);
        console.error('Full error:', error);
        console.error('=== END ERROR ===');
        
        // Check if it's a Drive API not enabled error
        if (error.message.includes('Google Drive API has not been used') || error.message.includes('it is disabled')) {
            const projectMatch = error.message.match(/project (\d+)/);
            const projectId = projectMatch ? projectMatch[1] : 'YOUR_PROJECT_ID';
            throw new Error(`Google Drive API is not enabled. Enable it here: https://console.developers.google.com/apis/api/drive.googleapis.com/overview?project=${projectId} (Wait a few minutes after enabling, then try again.)`);
        }
        
        // Don't continue - throw the error so user knows what went wrong
        throw new Error(`Failed to create folder: ${error.message}`);
    }

    // Step 2: Create all forms in the folder
    for (const [formId, formConfig] of Object.entries(votingData.forms)) {
        try {
            console.log(`Creating form for team: ${formConfig.teamName}`);
            
            const createdForm = await createGoogleForm(accessToken, formConfig, targetFolderId);
            createdForms[formId] = {
                ...formConfig,
                ...createdForm,
                status: 'created'
            };
            
            // Add a small delay to avoid rate limiting
            await new Promise(resolve => setTimeout(resolve, 500));
        } catch (error) {
            console.error(`Error creating form for ${formConfig.teamName}:`, error);
            errors.push({
                formId: formId,
                teamName: formConfig.teamName,
                error: error.message
            });
            createdForms[formId] = {
                ...formConfig,
                status: 'error',
                error: error.message
            };
        }
    }
    
    // Step 3: Create results spreadsheet in the SAME folder
    let resultsSpreadsheet = null;
    try {
        console.log('Creating results spreadsheet...');
        resultsSpreadsheet = await createResultsSpreadsheet(accessToken, votingData, createdForms);
        console.log(`✓ Results spreadsheet created: ${resultsSpreadsheet.spreadsheetUrl}`);
        
        // Move results spreadsheet to the same folder as forms
        if (targetFolderId && resultsSpreadsheet.spreadsheetId) {
            try {
                console.log(`Moving results spreadsheet to folder ${targetFolderId}...`);
                await moveFileToFolder(accessToken, resultsSpreadsheet.spreadsheetId, targetFolderId);
                console.log(`✓ Results spreadsheet moved to folder`);
            } catch (moveError) {
                console.error('Failed to move results spreadsheet:', moveError);
            }
        }
    } catch (error) {
        console.error('Failed to create results spreadsheet:', error);
    }

    return {
        forms: createdForms,
        errors: errors,
        folderId: targetFolderId,
        folderName: folderName,
        folderUrl: targetFolderId ? `https://drive.google.com/drive/folders/${targetFolderId}` : null,
        resultsSpreadsheet: resultsSpreadsheet,
        summary: {
            total: Object.keys(votingData.forms).length,
            created: Object.keys(createdForms).filter(f => createdForms[f].status === 'created').length,
            errors: errors.length
        }
    };
}

// ==================== TEST DATA GENERATION ====================

/**
 * Generates random test data directly into the results spreadsheet
 * (Google Forms API doesn't support submitting responses, so this bypasses forms)
 * @param {string} accessToken - OAuth access token
 * @param {Object} createdForms - Created forms data
 * @param {Object} resultsSpreadsheet - Results spreadsheet data
 * @param {number} responsesPerForm - Number of fake responses per form (default 3)
 * @returns {Promise<Object>} Results
 */
/**
 * Links a form to a new response spreadsheet and returns the spreadsheet ID
 */
async function linkFormToSpreadsheet(accessToken, formId, formTitle) {
    // Create a response spreadsheet for this form
    const spreadsheetResponse = await fetch('https://sheets.googleapis.com/v4/spreadsheets', {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({
            properties: { title: `${formTitle} (Responses)` }
        })
    });
    
    if (!spreadsheetResponse.ok) {
        throw new Error('Failed to create response spreadsheet');
    }
    
    const spreadsheet = await spreadsheetResponse.json();
    const destinationId = spreadsheet.spreadsheetId;
    
    // Link the form to this spreadsheet
    const linkResponse = await fetch(`https://forms.googleapis.com/v1/forms/${formId}:batchUpdate`, {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({
            requests: [{
                updateFormInfo: {
                    info: { documentTitle: formTitle },
                    updateMask: 'documentTitle'
                }
            }]
        })
    });
    
    return destinationId;
}

/**
 * Writes fake responses directly to a form's response spreadsheet
 * Format: Headers are "Category [Project]", values are rankings like "1st (4 pts)"
 */
async function writeFormResponses(accessToken, spreadsheetId, formData, numResponses, isJudges) {
    const categories = ['Business Impact', 'Production Readiness', 'Presentation'];
    const votingOptions = formData.votingOptions || [];
    const numProjects = votingOptions.length;
    
    // Build headers: Timestamp, then for each category, each project
    // Format: Timestamp, [Category: Project1], [Category: Project2], ...
    const headers = ['Timestamp'];
    for (const cat of categories) {
        for (const project of votingOptions) {
            headers.push(`${cat} [${project}]`);
        }
    }
    
    // Generate fake responses
    const rows = [headers];
    
    for (let r = 0; r < numResponses; r++) {
        const row = [new Date().toISOString()];
        
        for (const cat of categories) {
            // Generate random ranking (shuffle ranks 1 to numProjects)
            const ranks = Array.from({ length: numProjects }, (_, i) => i + 1);
            for (let i = ranks.length - 1; i > 0; i--) {
                const j = Math.floor(Math.random() * (i + 1));
                [ranks[i], ranks[j]] = [ranks[j], ranks[i]];
            }
            
            // Add rank for each project in this category
            votingOptions.forEach((project, idx) => {
                const rank = ranks[idx];
                const suffix = rank === 1 ? 'st' : rank === 2 ? 'nd' : rank === 3 ? 'rd' : 'th';
                const points = numProjects - rank + 1;
                row.push(`${rank}${suffix} (${points} pts)`);
            });
        }
        
        rows.push(row);
    }
    
    // Write to spreadsheet
    const writeResponse = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Sheet1!A1?valueInputOption=RAW`,
        {
            method: 'PUT',
            headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ values: rows })
        }
    );
    
    return writeResponse.ok;
}

/**
 * Aggregates responses from all form response spreadsheets into the Results spreadsheet
 */
async function aggregateFromFormSpreadsheets(accessToken, createdForms, resultsSpreadsheet) {
    console.log('=== AGGREGATING FROM FORM RESPONSE SPREADSHEETS ===');

    // Collect data into the standard format
    const collectedData = {
        participantVotes: [],
        rowVotes: [],
        judgesVotes: [],
        participantScores: {},
        rowScores: {},
        judgesScores: {}
    };

    console.log('Forms to aggregate:', Object.keys(createdForms.forms || {}).length);

    for (const [formId, formData] of Object.entries(createdForms.forms || {})) {
        // This function reads from response spreadsheets (for test data only)
        // For real responses, use aggregateFormResponses() which reads from Forms API
        if (!formData.responseSpreadsheetId) {
            continue;
        }

        const isJudges = formData.isJudgesForm || formData.teamName?.toLowerCase().includes('judge');
        const isRoW = formData.isRoWForm || formData.teamName?.toLowerCase().includes('row');
        const targetVotes = isJudges ? collectedData.judgesVotes : isRoW ? collectedData.rowVotes : collectedData.participantVotes;
        const targetScores = isJudges ? collectedData.judgesScores : isRoW ? collectedData.rowScores : collectedData.participantScores;

        console.log(`Reading responses from: ${formData.teamName} (${formData.responseSpreadsheetId})`);

        // Read the response spreadsheet
        const readResponse = await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${formData.responseSpreadsheetId}/values/Sheet1!A1:ZZ1000`,
            { headers: { 'Authorization': `Bearer ${accessToken}` } }
        );

        if (!readResponse.ok) {
            console.error(`Failed to read ${formData.teamName} responses`);
            continue;
        }

        const data = await readResponse.json();
        const rows = data.values || [];
        console.log(`  Read ${rows.length} rows from spreadsheet`);

        if (rows.length < 2) {
            console.log(`  Skipping - no data rows (only ${rows.length} rows)`);
            continue;
        }

        const headers = rows[0];
        console.log(`  Headers: ${headers.slice(0, 5).join(', ')}...`);

        // Parse each response row
        // Format: Headers are "Category [Project]", values are rankings like "1st (4 pts)"
        for (let rowIdx = 1; rowIdx < rows.length; rowIdx++) {
            const row = rows[rowIdx];
            const timestamp = row[0] || new Date().toISOString();

            // Parse responses - headers are like "Business Impact [Project Name]"
            for (let colIdx = 1; colIdx < headers.length; colIdx++) {
                const header = headers[colIdx] || '';
                const value = row[colIdx] || '';

                // Parse header: "Business Impact [Project Name]"
                const match = header.match(/^(.+?) \[(.+?)\]$/);
                if (!match) continue;

                const category = match[1];
                const project = match[2];

                // Parse value: "1st (4 pts)" - extract both rank and points from the string
                const rankMatch = value.match(/^(\d+)/);
                const pointsMatch = value.match(/\((\d+)\s*pts?\)/i);
                if (!rankMatch) continue;

                const rank = parseInt(rankMatch[1]);
                const points = pointsMatch ? parseInt(pointsMatch[1]) : 0;

                targetVotes.push([timestamp, '', formData.teamName, category, project, rank, points]);

                if (!targetScores[project]) {
                    targetScores[project] = { impact: [], readiness: [], presentation: [] };
                }

                if (category === 'Business Impact') targetScores[project].impact.push(points);
                else if (category === 'Production Readiness') targetScores[project].readiness.push(points);
                else if (category === 'Presentation') targetScores[project].presentation.push(points);
            }
        }
    }

    console.log(`Aggregated ${collectedData.participantVotes.length} participant votes, ${collectedData.rowVotes.length} RoW votes, ${collectedData.judgesVotes.length} judges votes`);

    // Use shared function to process and write all results
    const result = await processAndWriteResults(accessToken, resultsSpreadsheet.spreadsheetId, collectedData);
    
    console.log('=== AGGREGATION COMPLETE ===');
    return result;
}

async function generateTestData(accessToken, createdForms, resultsSpreadsheet, responsesPerForm = 3) {
    console.log('=== GENERATING TEST DATA (via Form Response Spreadsheets) ===');
    
    // Step 1: Write fake responses to each form's linked response spreadsheet
    for (const [formId, formData] of Object.entries(createdForms.forms || {})) {
        if (!formData.formId || formData.status !== 'created') continue;
        
        const isJudges = formData.isJudgesForm || formData.teamName === 'Judges';
        const isRoW = formData.isRoWForm || formData.teamName === 'RoW';
        const numResponses = (isJudges || isRoW) ? 5 : responsesPerForm;
        
        // Use already-linked response spreadsheet, or create one if missing
        let responseSpreadsheetId = formData.responseSpreadsheetId;
        
        if (!responseSpreadsheetId) {
            console.log(`Creating response spreadsheet for: ${formData.teamName}`);
            
            const spreadsheetResponse = await fetch('https://sheets.googleapis.com/v4/spreadsheets', {
                method: 'POST',
                headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    properties: { title: `${formData.teamName} (Responses)` }
                })
            });
            
            if (!spreadsheetResponse.ok) {
                console.error(`Failed to create response spreadsheet for ${formData.teamName}`);
                continue;
            }
            
            const spreadsheet = await spreadsheetResponse.json();
            responseSpreadsheetId = spreadsheet.spreadsheetId;
            formData.responseSpreadsheetId = responseSpreadsheetId;
            
            // Move to folder
            if (createdForms.folderId) {
                try {
                    await moveFileToFolder(accessToken, responseSpreadsheetId, createdForms.folderId);
                } catch (e) { }
            }
        } else {
            console.log(`Using linked response spreadsheet for: ${formData.teamName}`);
        }
        
        // Write fake responses to this spreadsheet
        const success = await writeFormResponses(accessToken, responseSpreadsheetId, formData, numResponses, isJudges);
        if (success) {
            console.log(`✓ Wrote ${numResponses} responses to ${formData.teamName} response sheet`);
        }
    }
    
    // Step 2: Aggregate from all response spreadsheets into Results spreadsheet
    console.log('\n=== AGGREGATING RESPONSES ===');
    const result = await aggregateFromFormSpreadsheets(accessToken, createdForms, resultsSpreadsheet);
    
    console.log('=== TEST DATA GENERATION COMPLETE ===');
    return {
        rawVotes: result.participantVotes + (result.rowVotes || 0) + result.judgesVotes,
        projects: result.projects,
        spreadsheetUrl: resultsSpreadsheet.spreadsheetUrl
    };
}

// Keep the old direct method as backup
async function generateTestDataDirect(accessToken, createdForms, resultsSpreadsheet, responsesConfig = 5) {
    console.log('=== GENERATING TEST DATA (Direct) ===');

    // Handle both old single number and new object config
    let responderCounts;
    if (typeof responsesConfig === 'object') {
        responderCounts = {
            participants: responsesConfig.participants || 5,
            row: responsesConfig.row || 5,
            judges: responsesConfig.judges || 5
        };
    } else {
        // Legacy: single number for all
        responderCounts = {
            participants: responsesConfig,
            row: responsesConfig,
            judges: responsesConfig
        };
    }
    console.log('Responder counts:', responderCounts);

    // Collect data into the standard format
    const collectedData = {
        participantVotes: [],
        rowVotes: [],
        judgesVotes: [],
        participantScores: {},
        rowScores: {},
        judgesScores: {}
    };
    const categories = ['Business Impact', 'Production Readiness', 'Presentation'];
    
    // Generate random responses for each form
    console.log('>>> Forms available:', Object.keys(createdForms.forms || {}).length);
    for (const [formId, formData] of Object.entries(createdForms.forms || {})) {
        console.log(`>>> Checking form: ${formId}, teamName: ${formData.teamName}, votingOptions: ${formData.votingOptions?.length || 0}`);
        if (!formData.votingOptions || formData.votingOptions.length === 0) {
            console.log(`>>> SKIPPING ${formId} - no votingOptions!`);
            continue;
        }
        
        const isJudges = formData.isJudgesForm || formData.teamName?.toLowerCase().includes('judge');
        const isRoW = formData.isRoWForm || formData.teamName?.toLowerCase().includes('row');
        const votingOptions = formData.votingOptions;
        const numProjects = votingOptions.length;
        const targetVotes = isJudges ? collectedData.judgesVotes : isRoW ? collectedData.rowVotes : collectedData.participantVotes;
        const targetScores = isJudges ? collectedData.judgesScores : isRoW ? collectedData.rowScores : collectedData.participantScores;
        
        // Use specific count for each form type
        const numResponses = isJudges ? responderCounts.judges : isRoW ? responderCounts.row : responderCounts.participants;
        const typeLabel = isJudges ? 'JUDGES' : isRoW ? 'ROW' : 'participant';
        console.log(`Generating ${numResponses} test responses for: ${formData.teamName} (${typeLabel})`);
        
        for (let r = 0; r < numResponses; r++) {
            const timestamp = new Date().toISOString();
            const fakeEmail = isJudges 
                ? `judge${r + 1}@company.test`
                : isRoW 
                    ? `row${r + 1}@company.test`
                    : `tester${r + 1}@${formData.teamName.toLowerCase().replace(/\s+/g, '')}.test`;
            
            for (const category of categories) {
                const ranks = Array.from({ length: numProjects }, (_, i) => i + 1);
                for (let i = ranks.length - 1; i > 0; i--) {
                    const j = Math.floor(Math.random() * (i + 1));
                    [ranks[i], ranks[j]] = [ranks[j], ranks[i]];
                }
                
                votingOptions.forEach((project, idx) => {
                    const rank = ranks[idx];
                    const points = numProjects - rank + 1;
                    
                    targetVotes.push([timestamp, fakeEmail, formData.teamName, category, project, rank, points]);
                    
                    if (!targetScores[project]) {
                        targetScores[project] = { impact: [], readiness: [], presentation: [] };
                    }
                    
                    if (category === 'Business Impact') targetScores[project].impact.push(points);
                    else if (category === 'Production Readiness') targetScores[project].readiness.push(points);
                    else if (category === 'Presentation') targetScores[project].presentation.push(points);
                });
            }
        }
    }
    
    console.log(`Generated ${collectedData.participantVotes.length} participant votes, ${collectedData.rowVotes.length} RoW votes, ${collectedData.judgesVotes.length} judges votes`);
    
    // Use shared function to process and write all results
    const result = await processAndWriteResults(accessToken, resultsSpreadsheet.spreadsheetId, collectedData);
    result.spreadsheetUrl = resultsSpreadsheet.spreadsheetUrl;
    result.rawVotes = collectedData.participantVotes.length + collectedData.rowVotes.length + collectedData.judgesVotes.length;
    
    console.log('=== TEST DATA GENERATION COMPLETE ===');
    return result;
}

// ==================== GOOGLE FORMS TEMPLATE GENERATOR ====================

/**
 * Generates Google Forms HTML template for manual form creation
 * @param {Object} formConfig - Form configuration object
 * @returns {string} HTML template for the form
 */
function generateFormTemplate(formConfig) {
    const optionsHtml = formConfig.votingOptions.map((team, index) => 
        `<option value="${team}">${team}</option>`
    ).join('\n');

    return `
<!DOCTYPE html>
<html>
<head>
    <title>${formConfig.formTitle}</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            max-width: 600px;
            margin: 50px auto;
            padding: 20px;
            background: #f5f5f5;
        }
        .form-container {
            background: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        h1 {
            color: #333;
            margin-bottom: 10px;
        }
        .description {
            color: #666;
            margin-bottom: 30px;
            line-height: 1.6;
        }
        .note {
            background: #fff3cd;
            border-left: 4px solid #ffc107;
            padding: 15px;
            margin: 20px 0;
            border-radius: 4px;
        }
        .instructions {
            background: #e7f3ff;
            border-left: 4px solid #2196F3;
            padding: 15px;
            margin: 20px 0;
            border-radius: 4px;
        }
        ol {
            margin: 10px 0;
            padding-left: 25px;
        }
        li {
            margin: 8px 0;
        }
    </style>
</head>
<body>
    <div class="form-container">
        <h1>${formConfig.formTitle}</h1>
        <div class="description">
            ${formConfig.formDescription}
        </div>
        
        <div class="note">
            <strong>Note:</strong> This form is for members of team <strong>${formConfig.teamName}</strong> only.
            Your team has been excluded from the voting options.
        </div>

        <div class="instructions">
            <h3>Instructions for Creating This Form in Google Forms:</h3>
            <ol>
                <li>Go to <a href="https://forms.google.com" target="_blank">Google Forms</a></li>
                <li>Create a new form</li>
                <li>Set the title to: <strong>${formConfig.formTitle}</strong></li>
                <li>Add description: <strong>${formConfig.formDescription}</strong></li>
                <li>Add a question: "Which team(s) do you want to vote for?"</li>
                <li>Set question type to: <strong>Multiple choice grid</strong> or <strong>Checkboxes</strong></li>
                <li>Add the following teams as options (one per line):</li>
                <ul style="margin-top: 10px;">
                    ${formConfig.votingOptions.map(team => `<li>${team}</li>`).join('')}
                </ul>
                <li>Go to Settings (gear icon) → Responses</li>
                <li>Enable: <strong>"Collect email addresses"</strong> → Set to <strong>"Do not collect"</strong></li>
                <li>Enable: <strong>"Limit to 1 response"</strong> → Set to <strong>OFF</strong> (for anonymous voting)</li>
                <li>Enable: <strong>"See summary charts and text responses"</strong> → Set to <strong>OFF</strong> (for anonymity)</li>
                <li>Click "Send" and copy the form link</li>
            </ol>
        </div>

        <div class="instructions">
            <h3>Teams Available for Voting:</h3>
            <ul>
                ${formConfig.votingOptions.map(team => `<li>${team}</li>`).join('')}
            </ul>
            <p><strong>Excluded Team:</strong> ${formConfig.excludedTeam} (your team)</p>
        </div>

        <div class="instructions">
            <h3>Assigned Participants:</h3>
            <p>The following participants should use this form:</p>
            <ul>
                ${formConfig.assignedParticipants.map(p => `<li>${p}</li>`).join('')}
            </ul>
        </div>
    </div>
</body>
</html>
`;
}

// ==================== EXPORT FUNCTIONALITY ====================

/**
 * Exports voting form configurations to a downloadable JSON file
 * @param {Object} votingData - The voting data from generateVotingForms()
 */
function exportVotingConfig(votingData) {
    const dataStr = JSON.stringify(votingData, null, 2);
    const dataBlob = new Blob([dataStr], { type: 'application/json' });
    const url = URL.createObjectURL(dataBlob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `voting-config-${new Date().toISOString().split('T')[0]}.json`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

/**
 * Generates a summary report of voting forms
 * @param {Object} votingData - The voting data from generateVotingForms()
 * @returns {string} HTML report
 */
function generateVotingReport(votingData) {
    const { teams, forms, participantAssignments, summary } = votingData || {};
    
    // Handle resumed sessions with minimal data
    if (!summary) {
        return `
            <h2 style="color: #333;">Resumed Voting Session</h2>
            <p style="color: #333;">Loaded from Google Drive. Use the buttons below to aggregate responses.</p>
        `;
    }
    
    let report = `
        <h2 style="color: #333;">Hackathon Voting Forms Summary</h2>
        <div style="margin: 20px 0; color: #333;">
            <p style="color: #333;"><strong style="color: #333;">Total Teams:</strong> ${summary.totalTeams}</p>
            <p style="color: #333;"><strong style="color: #333;">Total Participants:</strong> ${summary.totalParticipants}</p>
            <p style="color: #333;"><strong style="color: #333;">Total Forms Needed:</strong> ${summary.totalForms}</p>
        </div>
        
        <h3 style="color: #333;">Team Assignments</h3>
        <table border="1" cellpadding="10" cellspacing="0" style="border-collapse: collapse; width: 100%; margin: 20px 0;">
            <thead>
                <tr style="background: #f0f0f0;">
                    <th style="color: #333;">Team Name</th>
                    <th style="color: #333;">Team Members</th>
                    <th style="color: #333;">Form ID</th>
                    <th style="color: #333;">Teams They Can Vote For</th>
                </tr>
            </thead>
            <tbody>
    `;

    Object.keys(forms || {}).forEach(formId => {
        const form = forms[formId];
        const members = (form.assignedParticipants || []).join(', ');
        const votingOptions = (form.votingOptions || []).join(', ');
        
        report += `
            <tr>
                <td style="color: #333;"><strong style="color: #333;">${form.teamName}</strong></td>
                <td style="color: #333;">${members}</td>
                <td style="color: #333;">${formId}</td>
                <td style="color: #333;">${votingOptions}</td>
            </tr>
        `;
    });

    report += `
            </tbody>
        </table>
        
        <h3 style="color: #333;">Participant Form Assignments</h3>
        <table border="1" cellpadding="10" cellspacing="0" style="border-collapse: collapse; width: 100%; margin: 20px 0;">
            <thead>
                <tr style="background: #f0f0f0;">
                    <th style="color: #333;">Participant</th>
                    <th style="color: #333;">Their Team</th>
                    <th style="color: #333;">Form ID to Use</th>
                </tr>
            </thead>
            <tbody>
    `;

    Object.keys(participantAssignments).forEach(participant => {
        const formIds = participantAssignments[participant];
        // Find which team this participant belongs to
        let participantTeam = 'Unknown';
        Object.keys(teams).forEach(teamName => {
            if (teams[teamName].includes(participant)) {
                participantTeam = teamName;
            }
        });
        
        report += `
            <tr>
                <td style="color: #333;"><strong style="color: #333;">${participant}</strong></td>
                <td style="color: #333;">${participantTeam}</td>
                <td style="color: #333;">${formIds.join(', ')}</td>
            </tr>
        `;
    });

    report += `
            </tbody>
        </table>
    `;

    return report;
}

// ==================== UI INTEGRATION ====================

/**
 * Adds a voting button to the draw screen (call this after draw is complete)
 */
function addVotingButtonToUI() {
    const votingBtn = document.getElementById('generateVotingFormsBtn');
    if (!votingBtn) return;
    
    // Add "Resume Last Session" button if not already present
    if (!document.getElementById('resumeVotingBtn')) {
        const resumeBtn = document.createElement('button');
        resumeBtn.id = 'resumeVotingBtn';
        resumeBtn.textContent = '📂 RESUME LAST FORMS';
        resumeBtn.style.cssText = 'margin-left: 10px; padding: 10px 15px; background: #4a90d9; color: white; border: none; border-radius: 5px; cursor: pointer; font-weight: bold;';
        resumeBtn.addEventListener('click', async () => {
            try {
                // Get access token first
                let token = _votingAccessToken || sessionStorage.getItem('votingAccessToken');
                if (!token && typeof window.getAccessToken === 'function') {
                    token = window.getAccessToken();
                }
                
                if (!token) {
                    const clientId = document.getElementById('googleDriveClientId')?.value?.trim() ||
                                   localStorage.getItem('googleDriveClientId');
                    if (clientId && typeof window.signInWithGoogle === 'function') {
                        token = await window.signInWithGoogle('select_account');
                        _votingAccessToken = token;
                        sessionStorage.setItem('votingAccessToken', token);
                    }
                }
                
                if (!token) {
                    alert('Please authenticate first.');
                    return;
                }
                
                resumeBtn.textContent = '📂 Searching...';
                resumeBtn.disabled = true;
                
                // Search Google Drive for "Voting Forms" folders
                console.log('Searching for Voting Forms folders in Google Drive...');
                let searchResponse = await fetch(
                    `https://www.googleapis.com/drive/v3/files?q=name contains 'Voting Forms' and mimeType='application/vnd.google-apps.folder' and trashed=false&orderBy=createdTime desc&pageSize=10&fields=files(id,name,createdTime)`,
                    { headers: { 'Authorization': `Bearer ${token}` } }
                );
                
                // If auth failed, try to re-authenticate
                if (!searchResponse.ok && (searchResponse.status === 401 || searchResponse.status === 403)) {
                    console.log('Token expired, re-authenticating...');
                    _votingAccessToken = null;
                    sessionStorage.removeItem('votingAccessToken');
                    
                    const clientId = document.getElementById('googleDriveClientId')?.value?.trim() ||
                                   document.getElementById('googleDriveClientIdForExport')?.value?.trim() ||
                                   localStorage.getItem('googleDriveClientId');
                    
                    if (clientId && typeof window.signInWithGoogle === 'function') {
                        token = await window.signInWithGoogle('select_account');
                        _votingAccessToken = token;
                        sessionStorage.setItem('votingAccessToken', token);
                        
                        // Retry the search
                        searchResponse = await fetch(
                            `https://www.googleapis.com/drive/v3/files?q=name contains 'Voting Forms' and mimeType='application/vnd.google-apps.folder' and trashed=false&orderBy=createdTime desc&pageSize=10&fields=files(id,name,createdTime)`,
                            { headers: { 'Authorization': `Bearer ${token}` } }
                        );
                    }
                }
                
                if (!searchResponse.ok) {
                    const errText = await searchResponse.text();
                    console.error('Drive search failed:', searchResponse.status, errText);
                    throw new Error('Failed to search Drive - please try again');
                }
                
                const searchData = await searchResponse.json();
                const folders = searchData.files || [];
                
                if (folders.length === 0) {
                    alert('No "Voting Forms" folders found in Google Drive.');
                    resumeBtn.textContent = '📂 RESUME LAST FORMS';
                    resumeBtn.disabled = false;
                    return;
                }
                
                // Get the most recent folder (already sorted by createdTime desc)
                const latestFolder = folders[0];
                console.log(`Found latest folder: ${latestFolder.name} (${latestFolder.id})`);
                
                // List contents of the folder
                const contentsResponse = await fetch(
                    `https://www.googleapis.com/drive/v3/files?q='${latestFolder.id}' in parents and trashed=false&fields=files(id,name,mimeType,webViewLink)`,
                    { headers: { 'Authorization': `Bearer ${token}` } }
                );
                
                const contentsData = await contentsResponse.json();
                const files = contentsData.files || [];
                
                console.log(`Folder contains ${files.length} files:`, files.map(f => f.name));
                
                // Reconstruct createdForms object
                const createdForms = {
                    forms: {},
                    folderId: latestFolder.id,
                    folderName: latestFolder.name,
                    folderUrl: `https://drive.google.com/drive/folders/${latestFolder.id}`,
                    resultsSpreadsheet: null
                };
                
                // Find results spreadsheet and forms
                for (const file of files) {
                    if (file.name.startsWith('Voting Results') && file.mimeType === 'application/vnd.google-apps.spreadsheet') {
                        createdForms.resultsSpreadsheet = {
                            spreadsheetId: file.id,
                            spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${file.id}`
                        };
                    } else if (file.mimeType === 'application/vnd.google-apps.form') {
                        // All forms in this folder are voting forms - use the filename as team name
                        const formName = file.name;
                        const isJudges = formName.toLowerCase().includes('judge');
                        const isRoW = formName.toLowerCase().includes('row');
                        
                        createdForms.forms[file.id] = {
                            formId: file.id,
                            teamName: formName,
                            status: 'created',
                            responderUri: `https://docs.google.com/forms/d/${file.id}/viewform`,
                            editUri: `https://docs.google.com/forms/d/${file.id}/edit`,
                            isJudgesForm: isJudges,
                            isRoWForm: isRoW
                        };
                        console.log(`Form: ${formName}, isJudges: ${isJudges}, isRoW: ${isRoW}`);
                    }
                }
                
                console.log(`Found ${Object.keys(createdForms.forms).length} forms in folder`);
                
                // Try to get voting options from first form
                const firstFormKey = Object.keys(createdForms.forms)[0];
                const firstFormData = createdForms.forms[firstFormKey];
                const actualFormId = firstFormData?.formId; // Get the ACTUAL Google Form ID, not the key
                
                console.log(`First form key: ${firstFormKey}, actual Google Form ID: ${actualFormId}`);
                
                if (actualFormId) {
                    try {
                        console.log(`Fetching form structure from: https://forms.googleapis.com/v1/forms/${actualFormId}`);
                        const formResponse = await fetch(`https://forms.googleapis.com/v1/forms/${actualFormId}`, {
                            headers: { 'Authorization': `Bearer ${token}` }
                        });
                        if (formResponse.ok) {
                            const formData = await formResponse.json();
                            console.log(`Form structure received, items: ${formData.items?.length || 0}`);
                            const items = formData.items || [];
                            // Extract project names from first grid question
                            for (const item of items) {
                                if (item.questionGroupItem) {
                                    const questions = item.questionGroupItem.questions || [];
                                    const votingOptions = questions.map(q => q.rowQuestion?.title).filter(Boolean);
                                    console.log(`Extracted ${votingOptions.length} voting options:`, votingOptions);
                                    // Add votingOptions to all forms
                                    for (const fid of Object.keys(createdForms.forms)) {
                                        createdForms.forms[fid].votingOptions = votingOptions;
                                    }
                                    break;
                                }
                            }
                        } else {
                            console.error(`Failed to fetch form structure: ${formResponse.status}`);
                        }
                    } catch (e) {
                        console.error('Could not fetch form details:', e);
                    }
                } else {
                    console.error('No actual form ID found - cannot fetch voting options');
                }
                
                console.log('Reconstructed createdForms:', createdForms);
                
                // Create minimal votingData (not all fields needed for display)
                const votingData = { forms: {} };
                
                resumeBtn.textContent = '📂 RESUME LAST FORMS';
                resumeBtn.disabled = false;
                
                showVotingFormsModal(votingData, createdForms, `Resumed: ${latestFolder.name}`, false, token);
                
            } catch (e) {
                console.error('Error resuming session:', e);
                alert('Error: ' + e.message);
                resumeBtn.textContent = '📂 RESUME LAST FORMS';
                resumeBtn.disabled = false;
            }
        });
        votingBtn.parentNode.insertBefore(resumeBtn, votingBtn.nextSibling);
    }
    
    // Only attach handler once
    if (votingBtn.dataset.handlerAttached) return;
    votingBtn.dataset.handlerAttached = 'true';
    
    votingBtn.addEventListener('click', async () => {
        try {
            // Check if drawState and config are available
            if (typeof drawState === 'undefined' || typeof config === 'undefined') {
                alert('Draw state not available. Please complete the draw first.');
                return;
            }
            
            if (!drawState || !drawState.drawComplete) {
                alert('Please complete the draw first before generating voting forms.');
                return;
            }

            // Get access token
            let currentAccessToken = _votingAccessToken || sessionStorage.getItem('votingAccessToken');
            if (!currentAccessToken && typeof window.getAccessToken === 'function') {
                currentAccessToken = window.getAccessToken();
            }
            
            // Validate token if we have one
            if (currentAccessToken) {
                try {
                    const testResponse = await fetch('https://www.googleapis.com/drive/v3/about?fields=user', {
                        headers: { 'Authorization': `Bearer ${currentAccessToken}` }
                    });
                    if (!testResponse.ok) {
                        console.log('Token expired, clearing...');
                        currentAccessToken = null;
                        _votingAccessToken = null;
                        sessionStorage.removeItem('votingAccessToken');
                    }
                } catch (e) {
                    currentAccessToken = null;
                }
            }
            
            // If no valid token, need to authenticate
            if (!currentAccessToken) {
                const clientId = document.getElementById('googleDriveClientId')?.value?.trim() ||
                               localStorage.getItem('googleDriveClientId') ||
                               (typeof GOOGLE_DRIVE_CONFIG !== 'undefined' ? GOOGLE_DRIVE_CONFIG.clientId : '');
                
                if (!clientId) {
                    alert('Please enter your Google OAuth Client ID in the settings first.');
                    return;
                }
                
                if (typeof initGoogleDriveAPI === 'function') {
                    try { await initGoogleDriveAPI(clientId); } catch (e) { }
                }
                
                if (typeof window.signInWithGoogle === 'function') {
                    try {
                        currentAccessToken = await window.signInWithGoogle('select_account');
                        _votingAccessToken = currentAccessToken;
                        sessionStorage.setItem('votingAccessToken', currentAccessToken);
                    } catch (authError) {
                        alert('Authentication failed: ' + authError.message);
                        return;
                    }
                } else {
                    alert('Google Sign-In not available. Please refresh the page.');
                    return;
                }
            }
            
            // Generate voting form configurations (no email restrictions)
            console.log('Generating voting forms...');
            const votingData = generateVotingForms(drawState, config);

            // Show loading modal
            showVotingFormsModal(votingData, null, 'Creating folder and Google Forms...', true, currentAccessToken);
            
            // Create forms
            try {
                const createdForms = await createAllVotingForms(currentAccessToken, votingData);
                
                // Save to localStorage for resume functionality
                localStorage.setItem('lastVotingForms', JSON.stringify(createdForms));
                localStorage.setItem('lastVotingData', JSON.stringify(votingData));
                console.log('✓ Saved voting session to localStorage');
                
                if (createdForms.folderId) {
                    showVotingFormsModal(votingData, createdForms, null, false, currentAccessToken);
                } else {
                    showVotingFormsModal(votingData, createdForms, 'Forms created but folder creation failed. Check console for details.', false, currentAccessToken);
                }
            } catch (error) {
                console.error('Error creating forms:', error);
                showVotingFormsModal(votingData, null, `Error: ${error.message}. Check browser console (F12) for details.`, false, currentAccessToken);
            }
        } catch (error) {
            alert(`Error generating voting forms: ${error.message}`);
            console.error(error);
        }
    });

    console.log('✓ Voting button handler attached');
}


/**
 * Shows a modal with voting form information
 * @param {Object} votingData - The voting data from generateVotingForms()
 * @param {Object} createdForms - Created forms data (optional)
 * @param {string} message - Optional message to display
 * @param {boolean} isLoading - Whether forms are being created
 * @param {string} accessToken - OAuth access token for API calls (optional)
 */
function showVotingFormsModal(votingData, createdForms = null, message = null, isLoading = false, accessToken = null) {
    // Remove existing modal if any
    const existingModal = document.getElementById('votingFormsModal');
    if (existingModal) {
        existingModal.remove();
    }

    const modal = document.createElement('div');
    modal.id = 'votingFormsModal';
    modal.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0, 0, 0, 0.8);
        z-index: 10000;
        display: flex;
        justify-content: center;
        align-items: center;
        overflow-y: auto;
        padding: 20px;
    `;

    const content = document.createElement('div');
    content.style.cssText = `
        background: white;
        padding: 30px;
        border-radius: 10px;
        max-width: 900px;
        width: 100%;
        max-height: 90vh;
        overflow-y: auto;
        position: relative;
        color: #333 !important;
    `;
    content.setAttribute('style', content.style.cssText + ' color: #333 !important;');

    const closeBtn = document.createElement('button');
    closeBtn.textContent = '×';
    closeBtn.style.cssText = `
        position: absolute;
        top: 10px;
        right: 10px;
        background: #f44336;
        color: white;
        border: none;
        width: 30px;
        height: 30px;
        border-radius: 50%;
        cursor: pointer;
        font-size: 20px;
        line-height: 1;
    `;
    closeBtn.onclick = () => modal.remove();

    // Add global style for text color
    const style = document.createElement('style');
    style.textContent = `
        #votingFormsModal * {
            color: #333 !important;
        }
        #votingFormsModal h2, #votingFormsModal h3 {
            color: #333 !important;
        }
        #votingFormsModal p, #votingFormsModal td, #votingFormsModal th {
            color: #333 !important;
        }
        #votingFormsModal strong {
            color: #333 !important;
        }
    `;
    document.head.appendChild(style);
    
    let report = generateVotingReport(votingData);
    
    // Add created forms section if forms were created
    if (createdForms && createdForms.forms) {
        let folderInfo = '';
        if (createdForms.folderName && createdForms.folderUrl) {
            folderInfo = `<div style="background: #e8f5e9; border-left: 4px solid #4CAF50; padding: 15px; margin: 20px 0; border-radius: 4px; color: #333;">
                <h3 style="color: #333;">📁 Forms Created in Folder</h3>
                <p style="color: #333;"><strong style="color: #333;">Folder Name:</strong> ${createdForms.folderName}</p>
                <p style="color: #333;"><strong style="color: #333;">Folder Location:</strong> <a href="${createdForms.folderUrl}" target="_blank" style="color: #2196F3; text-decoration: underline;">Open Folder in Google Drive</a></p>
                <p style="font-size: 0.9em; color: #666;">All forms have been created in a timestamped folder in the same location as your spreadsheet.</p>
            </div>`;
            
            if (createdForms.resultsSpreadsheet && createdForms.resultsSpreadsheet.spreadsheetUrl) {
                folderInfo += `<div style="background: #e3f2fd; border-left: 4px solid #2196F3; padding: 15px; margin: 20px 0; border-radius: 4px; color: #333;">
                    <h3 style="color: #333;">📊 Results Spreadsheet</h3>
                    <p style="color: #333;"><strong style="color: #333;">Spreadsheet:</strong> <a href="${createdForms.resultsSpreadsheet.spreadsheetUrl}" target="_blank" style="color: #2196F3; text-decoration: underline;">Open Results Spreadsheet</a></p>
                    <p style="font-size: 0.9em; color: #666;">The spreadsheet contains two sheets: "Raw Votes" with all individual votes, and "Weighted Results" with scores weighted as Impact 40%, Readiness 40%, Presentation 20%.</p>
                </div>`;
            }
        } else {
            folderInfo = `<div style="background: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; margin: 20px 0; border-radius: 4px;">
                <h3>⚠️ Folder Creation Issue</h3>
                <p>Forms were created successfully, but the folder creation or form moving may have failed.</p>
                <p><strong>Possible reasons:</strong></p>
                <ul>
                    <li>Spreadsheet URL not detected (check browser console for details)</li>
                    <li>Forms may be in your Google Drive root folder</li>
                    <li>Check your Google Drive for forms with names like "Hackathon Team Voting - [Team Name] Members"</li>
                </ul>
                <p style="font-size: 0.9em; color: #666;">Check the browser console (F12) for detailed error messages.</p>
            </div>`;
        }
        
        const summaryText = createdForms.summary 
            ? `<p><strong>Status:</strong> ${createdForms.summary.created} of ${createdForms.summary.total} forms created successfully</p>
               ${createdForms.summary.errors > 0 ? `<p style="color: #f44336;"><strong>Errors:</strong> ${createdForms.summary.errors} forms failed to create</p>` : ''}`
            : `<p><strong>Forms found:</strong> ${Object.keys(createdForms.forms || {}).length}</p>`;
        
        const createdSection = `
            <div style="margin-top: 30px; padding-top: 20px; border-top: 2px solid #4CAF50;">
                <h2 style="color: #4CAF50;">✓ Google Forms</h2>
                ${folderInfo}
                ${summaryText}
                
                <table border="1" cellpadding="10" cellspacing="0" style="border-collapse: collapse; width: 100%; margin: 20px 0;">
                    <thead>
                <tr style="background: #f0f0f0;">
                    <th style="color: #333;">Team</th>
                    <th style="color: #333;">Status</th>
                    <th style="color: #333;">Form Link</th>
                    <th style="color: #333;">Edit Link</th>
                </tr>
            </thead>
            <tbody>
        `;
        
        let createdSectionBody = '';
        Object.keys(createdForms.forms).forEach(formId => {
            const form = createdForms.forms[formId];
            if (form.status === 'created' && form.responderUri) {
                createdSectionBody += `
                    <tr>
                        <td style="color: #333;"><strong style="color: #333;">${form.teamName}</strong></td>
                        <td style="color: #4CAF50;">✓ Created</td>
                        <td style="color: #333;"><a href="${form.responderUri}" target="_blank" style="color: #2196F3; text-decoration: underline;">Open Form</a></td>
                        <td style="color: #333;"><a href="${form.editUri || `https://docs.google.com/forms/d/${form.formId}/edit`}" target="_blank" style="color: #2196F3; text-decoration: underline;">Edit Form</a></td>
                    </tr>
                `;
            } else {
                createdSectionBody += `
                    <tr>
                        <td style="color: #333;"><strong style="color: #333;">${form.teamName}</strong></td>
                        <td style="color: #f44336;">✗ Error</td>
                        <td colspan="2" style="color: #333;">${form.error || 'Failed to create'}</td>
                    </tr>
                `;
            }
        });
        
        report += createdSection + createdSectionBody + `
                    </tbody>
                </table>
                
                <div style="background: #e8f5e9; border-left: 4px solid #4CAF50; padding: 15px; margin: 20px 0; border-radius: 4px; color: #333;">
                    <h3 style="color: #333;">✅ Anonymous Settings Configured</h3>
                    <p style="color: #333;"><strong style="color: #333;">The following settings have been automatically configured:</strong></p>
                    <ul style="color: #333;">
                        <li>✅ Email collection: <strong>Disabled</strong> (anonymous responses)</li>
                        <li>✅ Login requirement: <strong>Disabled</strong> (allows anonymous voting)</li>
                        <li>✅ Response editing: <strong>Disabled</strong> (prevents vote changes)</li>
                        <li>✅ Summary charts: <strong>Hidden</strong> (maintains anonymity during voting)</li>
                        <li>✅ Progress bar: <strong>Enabled</strong> (better user experience)</li>
                    </ul>
                    <p style="color: #333; font-size: 0.9em; margin-top: 10px;"><em>No manual configuration required! Forms are ready for anonymous voting.</em></p>
                </div>
            </div>
        `;
    }
    
    // Add message if provided
    if (message) {
        const messageStyle = isLoading ? 
            'background: #e7f3ff; border-left: 4px solid #2196F3; padding: 15px; margin: 20px 0; border-radius: 4px;' :
            'background: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; margin: 20px 0; border-radius: 4px;';
        report = `<div style="${messageStyle}"><strong>${isLoading ? '⏳ ' : 'ℹ️ '}</strong>${message}</div>` + report;
    }
    
    // Add next steps if forms weren't created automatically
    if (!createdForms || !createdForms.forms) {
        report += `
            <div style="margin-top: 30px; padding-top: 20px; border-top: 2px solid #ddd;">
                <h3>Next Steps:</h3>
                <ol>
                    <li>For each form listed above, create a Google Form following the instructions</li>
                    <li>Make sure to enable anonymous responses in Google Forms settings</li>
                    <li>Share the appropriate form link with each team's members</li>
                    <li>Collect responses after the hackathon concludes</li>
                </ol>
            </div>
        `;
    } else {
        report += `
            <div style="margin-top: 30px; padding-top: 20px; border-top: 2px solid #ddd;">
                <h3>Next Steps:</h3>
                <ol>
                    <li>✅ Anonymous settings are already configured automatically</li>
                    <li>Share the appropriate form link with each team's members</li>
                    <li>Collect responses after the hackathon concludes</li>
                    <li>Optional: You can preview forms using the "Edit Form" links above</li>
                </ol>
            </div>
        `;
    }
    
    // Add buttons if forms were created and results spreadsheet exists
    let actionButtons = '';
    if (createdForms && createdForms.resultsSpreadsheet) {
        actionButtons = `
            <button id="aggregateResponsesBtn" style="padding: 10px 20px; background: #FF9800; color: white; border: none; border-radius: 5px; cursor: pointer; margin-right: 10px;">
                📊 Aggregate Responses
            </button>
            <div style="display: flex; align-items: center; gap: 15px; margin: 15px 0; background: #f5f5f5; padding: 10px 15px; border-radius: 5px; flex-wrap: wrap;">
                <span style="font-weight: bold; font-size: 14px;">Test Responders:</span>
                <div style="display: flex; align-items: center; gap: 5px;">
                    <label for="testParticipants" style="font-size: 13px;">Participants:</label>
                    <input type="number" id="testParticipants" value="5" min="0" max="50" style="width: 45px; padding: 4px; border: 1px solid #ccc; border-radius: 3px; text-align: center;">
                </div>
                <div style="display: flex; align-items: center; gap: 5px;">
                    <label for="testRoW" style="font-size: 13px;">RoW:</label>
                    <input type="number" id="testRoW" value="5" min="0" max="50" style="width: 45px; padding: 4px; border: 1px solid #ccc; border-radius: 3px; text-align: center;">
                </div>
                <div style="display: flex; align-items: center; gap: 5px;">
                    <label for="testJudges" style="font-size: 13px;">Judges:</label>
                    <input type="number" id="testJudges" value="5" min="0" max="50" style="width: 45px; padding: 4px; border: 1px solid #ccc; border-radius: 3px; text-align: center;">
                </div>
                <button id="generateTestDataBtn" style="padding: 8px 16px; background: #9C27B0; color: white; border: none; border-radius: 5px; cursor: pointer;">
                    🧪 Generate Test Data
                </button>
                <button id="runSanityTestsBtn" style="padding: 8px 16px; background: #FF5722; color: white; border: none; border-radius: 5px; cursor: pointer; margin-left: 10px;">
                    🔬 Run Sanity Tests
                </button>
            </div>`;
    }
    
    content.innerHTML = `
        ${report}
        <div style="margin-top: 20px;">
            ${actionButtons}
            <button id="exportVotingConfigBtn" style="padding: 10px 20px; background: #4CAF50; color: white; border: none; border-radius: 5px; cursor: pointer; margin-right: 10px;">
                Export Configuration (JSON)
            </button>
            <button id="copyReportBtn" style="padding: 10px 20px; background: #2196F3; color: white; border: none; border-radius: 5px; cursor: pointer;">
                Copy Report to Clipboard
            </button>
        </div>
    `;

    content.appendChild(closeBtn);
    modal.appendChild(content);
    document.body.appendChild(modal);
    
    // Store access token in modal for later use
    if (accessToken) {
        modal.dataset.accessToken = accessToken;
    }

    // Add event listeners
    document.getElementById('exportVotingConfigBtn').addEventListener('click', () => {
        const exportData = {
            ...votingData,
            createdForms: createdForms
        };
        exportVotingConfig(exportData);
    });

    document.getElementById('copyReportBtn').addEventListener('click', () => {
        const text = content.innerText;
        navigator.clipboard.writeText(text).then(() => {
            alert('Report copied to clipboard!');
        }).catch(err => {
            console.error('Failed to copy:', err);
            alert('Failed to copy to clipboard. Please select and copy manually.');
        });
    });
    
    // Add aggregate button listener (for real form responses)
    const aggregateBtn = document.getElementById('aggregateResponsesBtn');
    if (aggregateBtn && createdForms && createdForms.resultsSpreadsheet) {
        aggregateBtn.addEventListener('click', async () => {
            const token = modal.dataset.accessToken || _votingAccessToken || sessionStorage.getItem('votingAccessToken');
            if (!token) {
                alert('Not authenticated. Please refresh and try again.');
                return;
            }

            aggregateBtn.disabled = true;
            aggregateBtn.textContent = '📊 Aggregating...';

            try {
                // ALWAYS use Forms API to get REAL responses (not the empty response spreadsheets)
                console.log('=== AGGREGATING REAL RESPONSES FROM FORMS API ===');
                const result = await aggregateFormResponses(token, createdForms, createdForms.resultsSpreadsheet);
                console.log('Aggregation result:', result);
                alert(`✓ Aggregated from Google Forms!\n\nParticipant votes: ${result.participantVotes || 0}\nRoW votes: ${result.rowVotes || 0}\nJudges votes: ${result.judgesVotes || 0}\nProjects: ${result.projects}`);

                aggregateBtn.textContent = '📊 Aggregated ✓';
                window.open(createdForms.resultsSpreadsheet.spreadsheetUrl, '_blank');
            } catch (error) {
                console.error('Error aggregating responses:', error);
                alert(`Error: ${error.message}`);
                aggregateBtn.disabled = false;
                aggregateBtn.textContent = '📊 Aggregate Real Responses';
            }
        });
    }

    // Add test data button listener (generates fake data directly into sheets)
    const testDataBtn = document.getElementById('generateTestDataBtn');
    if (testDataBtn && createdForms && createdForms.resultsSpreadsheet) {
        testDataBtn.addEventListener('click', async () => {
            const token = modal.dataset.accessToken || _votingAccessToken || sessionStorage.getItem('votingAccessToken');
            if (!token) {
                alert('Not authenticated. Please refresh and try again.');
                return;
            }

            // Get the number of responders for each form type
            const numParticipants = parseInt(document.getElementById('testParticipants')?.value) || 5;
            const numRoW = parseInt(document.getElementById('testRoW')?.value) || 5;
            const numJudges = parseInt(document.getElementById('testJudges')?.value) || 5;

            testDataBtn.disabled = true;
            testDataBtn.textContent = '🧪 Generating...';

            try {
                const targetSpreadsheet = createdForms.resultsSpreadsheet;
                console.log(`=== GENERATING TEST DATA ===`);
                console.log(`Target spreadsheet ID: ${targetSpreadsheet.spreadsheetId}`);
                console.log(`Target spreadsheet URL: ${targetSpreadsheet.spreadsheetUrl}`);
                console.log(`Responders - Participants: ${numParticipants}, RoW: ${numRoW}, Judges: ${numJudges}`);
                
                const result = await generateTestDataDirect(token, createdForms, targetSpreadsheet, {
                    participants: numParticipants,
                    row: numRoW,
                    judges: numJudges
                });
                console.log('Test data result:', result);
                
                const spreadsheetUrl = result.spreadsheetUrl || targetSpreadsheet.spreadsheetUrl;
                alert(`✓ Test Data Generated!\n\nWritten to: ${spreadsheetUrl}\n\nParticipants: ${numParticipants} responders\nRoW: ${numRoW} responders\nJudges: ${numJudges} responders\n\nTotal votes: ${result.rawVotes}\nProjects: ${result.projects}\n\nOpening results spreadsheet...`);

                testDataBtn.textContent = '🧪 Test Data Generated ✓';
                window.open(spreadsheetUrl, '_blank');
            } catch (error) {
                console.error('Error generating test data:', error);
                console.error('createdForms.resultsSpreadsheet:', createdForms.resultsSpreadsheet);
                alert(`Error: ${error.message}\n\nCheck browser console (F12) for details.`);
                testDataBtn.disabled = false;
                testDataBtn.textContent = '🧪 Generate Test Data';
            }
        });
    }
    
    // Run Sanity Tests button handler
    const sanityTestsBtn = document.getElementById('runSanityTestsBtn');
    if (sanityTestsBtn) {
        sanityTestsBtn.addEventListener('click', async () => {
            sanityTestsBtn.disabled = true;
            sanityTestsBtn.textContent = '🔬 Running Tests...';
            
            try {
                const token = accessToken || _votingAccessToken;
                if (!token) {
                    throw new Error('No access token available');
                }
                
                const spreadsheetId = createdForms.resultsSpreadsheet?.spreadsheetId;
                if (!spreadsheetId) {
                    throw new Error('No results spreadsheet found. Generate test data first.');
                }
                
                console.clear();
                const results = await runSanityTests(token, spreadsheetId);
                
                // Show results in alert
                const summary = results.failed === 0 
                    ? `✅ ALL ${results.passed} TESTS PASSED!`
                    : `⚠️ ${results.passed} passed, ${results.failed} failed`;
                
                let details = results.tests.map(t => `${t.passed ? '✅' : '❌'} ${t.name}`).join('\n');
                
                alert(`SANITY TEST RESULTS\n\n${summary}\n\n${details}\n\nSee console (F12) for full details.`);
                
            } catch (error) {
                console.error('Sanity test error:', error);
                alert(`Error: ${error.message}`);
            } finally {
                sanityTestsBtn.disabled = false;
                sanityTestsBtn.textContent = '🔬 Run Sanity Tests';
            }
        });
    }

    // Close on outside click
    modal.addEventListener('click', (e) => {
        if (e.target === modal) {
            modal.remove();
        }
    });
}

// ==================== AUTO-INITIALIZATION ====================

// Initialize main screen resume button when DOM is ready
document.addEventListener('DOMContentLoaded', function() {
    const mainResumeBtn = document.getElementById('resumeVotingMainBtn');
    if (mainResumeBtn) {
        mainResumeBtn.addEventListener('click', async function() {
            try {
                // Get access token first
                let token = _votingAccessToken || sessionStorage.getItem('votingAccessToken');
                if (!token && typeof window.getAccessToken === 'function') {
                    token = window.getAccessToken();
                }
                
                if (!token) {
                    const clientId = document.getElementById('googleDriveClientId')?.value?.trim() ||
                                   document.getElementById('googleDriveClientIdForExport')?.value?.trim() ||
                                   localStorage.getItem('googleDriveClientId');
                    if (!clientId) {
                        alert('Please enter your Google OAuth Client ID first (in the Google Integration section).');
                        return;
                    }
                    
                    if (typeof initGoogleDriveAPI === 'function') {
                        try { await initGoogleDriveAPI(clientId); } catch (e) { }
                    }
                    
                    if (typeof window.signInWithGoogle === 'function') {
                        token = await window.signInWithGoogle('select_account');
                        _votingAccessToken = token;
                        sessionStorage.setItem('votingAccessToken', token);
                    } else {
                        alert('Google Sign-In not available. Please refresh the page.');
                        return;
                    }
                }
                
                if (!token) {
                    alert('Please authenticate first.');
                    return;
                }
                
                mainResumeBtn.textContent = '📂 Searching...';
                mainResumeBtn.disabled = true;
                
                // Search Google Drive for "Voting Forms" folders
                console.log('Searching for Voting Forms folders in Google Drive...');
                let searchResponse = await fetch(
                    `https://www.googleapis.com/drive/v3/files?q=name contains 'Voting Forms' and mimeType='application/vnd.google-apps.folder' and trashed=false&orderBy=createdTime desc&pageSize=10&fields=files(id,name,createdTime)`,
                    { headers: { 'Authorization': `Bearer ${token}` } }
                );
                
                // If auth failed, try to re-authenticate
                if (!searchResponse.ok && (searchResponse.status === 401 || searchResponse.status === 403)) {
                    console.log('Token expired, re-authenticating...');
                    _votingAccessToken = null;
                    sessionStorage.removeItem('votingAccessToken');
                    
                    const clientId = document.getElementById('googleDriveClientId')?.value?.trim() ||
                                   document.getElementById('googleDriveClientIdForExport')?.value?.trim() ||
                                   localStorage.getItem('googleDriveClientId');
                    
                    if (clientId && typeof window.signInWithGoogle === 'function') {
                        token = await window.signInWithGoogle('select_account');
                        _votingAccessToken = token;
                        sessionStorage.setItem('votingAccessToken', token);
                        
                        // Retry the search
                        searchResponse = await fetch(
                            `https://www.googleapis.com/drive/v3/files?q=name contains 'Voting Forms' and mimeType='application/vnd.google-apps.folder' and trashed=false&orderBy=createdTime desc&pageSize=10&fields=files(id,name,createdTime)`,
                            { headers: { 'Authorization': `Bearer ${token}` } }
                        );
                    }
                }
                
                if (!searchResponse.ok) {
                    const errText = await searchResponse.text();
                    console.error('Drive search failed:', searchResponse.status, errText);
                    throw new Error('Failed to search Drive - please try again');
                }
                
                const searchData = await searchResponse.json();
                const folders = searchData.files || [];
                
                if (folders.length === 0) {
                    alert('No "Voting Forms" folders found in Google Drive. Generate forms first.');
                    mainResumeBtn.textContent = '📂 RESUME LAST FORMS';
                    mainResumeBtn.disabled = false;
                    return;
                }
                
                // Get the most recent folder
                const latestFolder = folders[0];
                console.log(`Found latest folder: ${latestFolder.name} (${latestFolder.id})`);
                
                // List contents of the folder
                const contentsResponse = await fetch(
                    `https://www.googleapis.com/drive/v3/files?q='${latestFolder.id}' in parents and trashed=false&fields=files(id,name,mimeType,webViewLink)`,
                    { headers: { 'Authorization': `Bearer ${token}` } }
                );
                
                const contentsData = await contentsResponse.json();
                const files = contentsData.files || [];
                
                // Reconstruct createdForms object
                const createdForms = {
                    forms: {},
                    folderId: latestFolder.id,
                    folderName: latestFolder.name,
                    folderUrl: `https://drive.google.com/drive/folders/${latestFolder.id}`,
                    resultsSpreadsheet: null
                };
                
                for (const file of files) {
                    if (file.name.startsWith('Voting Results') && file.mimeType === 'application/vnd.google-apps.spreadsheet') {
                        createdForms.resultsSpreadsheet = {
                            spreadsheetId: file.id,
                            spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${file.id}`
                        };
                    } else if (file.mimeType === 'application/vnd.google-apps.form') {
                        // All forms in this folder are voting forms - use the filename as team name
                        const formName = file.name;
                        const isJudges = formName.toLowerCase().includes('judge');
                        const isRoW = formName.toLowerCase().includes('row');
                        
                        createdForms.forms[file.id] = {
                            formId: file.id,
                            teamName: formName,
                            status: 'created',
                            responderUri: `https://docs.google.com/forms/d/${file.id}/viewform`,
                            editUri: `https://docs.google.com/forms/d/${file.id}/edit`,
                            isJudgesForm: isJudges,
                            isRoWForm: isRoW
                        };
                        console.log(`[Resume] Form: ${formName}, isJudges: ${isJudges}, isRoW: ${isRoW}`);
                    }
                }
                
                console.log(`Found ${Object.keys(createdForms.forms).length} forms in folder`);
                
                // Try to get voting options from first form
                const firstFormKey = Object.keys(createdForms.forms)[0];
                const firstFormData = createdForms.forms[firstFormKey];
                const actualFormId = firstFormData?.formId;
                
                console.log(`[Resume] First form key: ${firstFormKey}, actual Form ID: ${actualFormId}`);
                
                if (actualFormId) {
                    try {
                        console.log(`[Resume] Fetching form structure...`);
                        const formResponse = await fetch(`https://forms.googleapis.com/v1/forms/${actualFormId}`, {
                            headers: { 'Authorization': `Bearer ${token}` }
                        });
                        if (formResponse.ok) {
                            const formData = await formResponse.json();
                            console.log(`[Resume] Got form data, items: ${formData.items?.length || 0}`);
                            const items = formData.items || [];
                            for (const item of items) {
                                if (item.questionGroupItem) {
                                    const questions = item.questionGroupItem.questions || [];
                                    const votingOptions = questions.map(q => q.rowQuestion?.title).filter(Boolean);
                                    console.log(`[Resume] Extracted ${votingOptions.length} voting options:`, votingOptions);
                                    for (const fid of Object.keys(createdForms.forms)) {
                                        createdForms.forms[fid].votingOptions = votingOptions;
                                    }
                                    break;
                                }
                            }
                        } else {
                            console.error(`[Resume] Failed to fetch form: ${formResponse.status}`);
                            const errText = await formResponse.text();
                            console.error(`[Resume] Error details:`, errText);
                        }
                    } catch (e) {
                        console.error('[Resume] Could not fetch form details:', e);
                    }
                } else {
                    console.error('[Resume] No form ID found to fetch voting options');
                }
                
                const votingData = { forms: {} };
                
                mainResumeBtn.textContent = '📂 RESUME LAST FORMS';
                mainResumeBtn.disabled = false;
                
                showVotingFormsModal(votingData, createdForms, `Resumed: ${latestFolder.name}`, false, token);
                
            } catch (e) {
                console.error('Error resuming session:', e);
                alert('Error: ' + e.message);
                mainResumeBtn.textContent = '📂 RESUME LAST FORMS';
                mainResumeBtn.disabled = false;
            }
        });
    }
});

// Make functions globally available IMMEDIATELY so app.js can call them
// This runs as soon as the script loads
window.addVotingButtonToUI = addVotingButtonToUI;
window.generateTestData = generateTestData;

// ==================== SANITY TEST SUITE ====================
// Tests to verify the voting data processing is working correctly

/**
 * Run all sanity tests
 * @param {string} accessToken - OAuth access token
 * @param {string} spreadsheetId - Results spreadsheet ID to test against
 * @returns {Promise<Object>} Test results with pass/fail counts
 */
async function runSanityTests(accessToken, spreadsheetId) {
    console.log('╔══════════════════════════════════════════════════════════════╗');
    console.log('║              VOTING SANITY TEST SUITE                        ║');
    console.log('╚══════════════════════════════════════════════════════════════╝');
    
    const results = {
        passed: 0,
        failed: 0,
        tests: []
    };
    
    function addResult(name, passed, message, details = null) {
        const status = passed ? '✅ PASS' : '❌ FAIL';
        console.log(`${status}: ${name}`);
        if (message) console.log(`       ${message}`);
        if (details && !passed) console.log(`       Details:`, details);
        results.tests.push({ name, passed, message, details });
        if (passed) results.passed++;
        else results.failed++;
    }
    
    // ==================== UNIT TESTS ====================
    console.log('\n─── UNIT TESTS: Calculation Functions ───');
    
    // Test 1: calculateWeightedResults with known values
    try {
        const testScores = {
            'Project A': { impact: [4, 4, 4], readiness: [3, 3, 3], presentation: [2, 2, 2] },
            'Project B': { impact: [2, 2, 2], readiness: [4, 4, 4], presentation: [3, 3, 3] }
        };
        const weighted = calculateWeightedResults(testScores);
        
        // Project A: impact=4*0.4=1.6, readiness=3*0.4=1.2, presentation=2*0.2=0.4, total=3.2
        // Project B: impact=2*0.4=0.8, readiness=4*0.4=1.6, presentation=3*0.2=0.6, total=3.0
        const projA = weighted.find(r => r.project === 'Project A');
        const projB = weighted.find(r => r.project === 'Project B');
        
        const expectedA = { impact: 1.6, readiness: 1.2, presentation: 0.4, total: 3.2 };
        const expectedB = { impact: 0.8, readiness: 1.6, presentation: 0.6, total: 3.0 };
        
        const toleranceOk = (a, b) => Math.abs(a - b) < 0.001;
        
        const aOk = toleranceOk(projA.impact, expectedA.impact) && 
                    toleranceOk(projA.readiness, expectedA.readiness) &&
                    toleranceOk(projA.presentation, expectedA.presentation) &&
                    toleranceOk(projA.total, expectedA.total);
        const bOk = toleranceOk(projB.impact, expectedB.impact) && 
                    toleranceOk(projB.readiness, expectedB.readiness) &&
                    toleranceOk(projB.presentation, expectedB.presentation) &&
                    toleranceOk(projB.total, expectedB.total);
        
        addResult(
            'calculateWeightedResults - Category weights (40/40/20)',
            aOk && bOk,
            aOk && bOk ? 'Correct: A=3.2, B=3.0' : `Expected A=${JSON.stringify(expectedA)}, B=${JSON.stringify(expectedB)}`,
            { projA, projB }
        );
    } catch (e) {
        addResult('calculateWeightedResults - Category weights', false, e.message);
    }
    
    // Test 2: calculateWeightedResults sorts by total descending
    try {
        const testScores = {
            'Low': { impact: [1], readiness: [1], presentation: [1] },
            'High': { impact: [4], readiness: [4], presentation: [4] },
            'Mid': { impact: [2], readiness: [2], presentation: [2] }
        };
        const weighted = calculateWeightedResults(testScores);
        const order = weighted.map(r => r.project);
        const expectedOrder = ['High', 'Mid', 'Low'];
        const orderOk = order[0] === expectedOrder[0] && order[1] === expectedOrder[1] && order[2] === expectedOrder[2];
        
        addResult(
            'calculateWeightedResults - Sorting descending by total',
            orderOk,
            orderOk ? `Correct order: ${order.join(' > ')}` : `Expected ${expectedOrder.join(' > ')}, got ${order.join(' > ')}`
        );
    } catch (e) {
        addResult('calculateWeightedResults - Sorting', false, e.message);
    }
    
    // Test 3: calculateFinalResults with known values and scaling
    try {
        const pWeighted = [{ project: 'Test', total: 3.0 }];  // Participants (1-4 scale)
        const rWeighted = [{ project: 'Test', total: 3.75 }]; // RoW (1-5 scale) -> 3.75 * 0.8 = 3.0
        const jWeighted = [{ project: 'Test', total: 3.75 }]; // Judges (1-5 scale) -> 3.75 * 0.8 = 3.0
        
        const final = calculateFinalResults(pWeighted, rWeighted, jWeighted);
        const testProj = final.find(r => r[0] === 'Test');
        
        // pContrib = 3.0 * 0.4 = 1.2
        // rContrib = (3.75 * 0.8) * 0.2 = 3.0 * 0.2 = 0.6
        // jContrib = (3.75 * 0.8) * 0.4 = 3.0 * 0.4 = 1.2
        // total = 1.2 + 0.6 + 1.2 = 3.0
        
        const expectedP = 1.2, expectedR = 0.6, expectedJ = 1.2, expectedTotal = 3.0;
        const toleranceOk = (a, b) => Math.abs(a - b) < 0.01;
        
        const ok = testProj && 
                   toleranceOk(testProj[1], expectedP) &&
                   toleranceOk(testProj[2], expectedR) &&
                   toleranceOk(testProj[3], expectedJ) &&
                   toleranceOk(testProj[4], expectedTotal);
        
        addResult(
            'calculateFinalResults - Group weights (40/20/40) with scaling',
            ok,
            ok ? `Correct: P=${testProj[1]}, R=${testProj[2]}, J=${testProj[3]}, Total=${testProj[4]}` :
                 `Expected P=${expectedP}, R=${expectedR}, J=${expectedJ}, Total=${expectedTotal}`,
            testProj
        );
    } catch (e) {
        addResult('calculateFinalResults - Group weights', false, e.message);
    }
    
    // Test 4: Verify constants are correct
    try {
        const catOk = CATEGORY_WEIGHTS.impact === 0.4 && 
                      CATEGORY_WEIGHTS.readiness === 0.4 && 
                      CATEGORY_WEIGHTS.presentation === 0.2;
        const catSum = CATEGORY_WEIGHTS.impact + CATEGORY_WEIGHTS.readiness + CATEGORY_WEIGHTS.presentation;
        
        addResult(
            'Constants - CATEGORY_WEIGHTS sum to 1.0',
            catOk && Math.abs(catSum - 1.0) < 0.001,
            `Impact=${CATEGORY_WEIGHTS.impact}, Readiness=${CATEGORY_WEIGHTS.readiness}, Presentation=${CATEGORY_WEIGHTS.presentation}, Sum=${catSum}`
        );
        
        const grpOk = GROUP_WEIGHTS.participants === 0.4 && 
                      GROUP_WEIGHTS.row === 0.2 && 
                      GROUP_WEIGHTS.judges === 0.4;
        const grpSum = GROUP_WEIGHTS.participants + GROUP_WEIGHTS.row + GROUP_WEIGHTS.judges;
        
        addResult(
            'Constants - GROUP_WEIGHTS sum to 1.0',
            grpOk && Math.abs(grpSum - 1.0) < 0.001,
            `Participants=${GROUP_WEIGHTS.participants}, RoW=${GROUP_WEIGHTS.row}, Judges=${GROUP_WEIGHTS.judges}, Sum=${grpSum}`
        );
        
        addResult(
            'Constants - SCALE_FACTOR is 0.8 (4/5)',
            Math.abs(SCALE_FACTOR - 0.8) < 0.001,
            `SCALE_FACTOR=${SCALE_FACTOR}`
        );
    } catch (e) {
        addResult('Constants verification', false, e.message);
    }
    
    // ==================== INTEGRATION TESTS (if spreadsheet provided) ====================
    if (accessToken && spreadsheetId) {
        console.log('\n─── INTEGRATION TESTS: Spreadsheet Data ───');
        
        try {
            // Read all sheets
            const sheets = [
                'Participants Votes', 'RoW Votes', 'Judges Votes',
                'Participants Weighted Results', 'RoW Weighted Results', 'Judges Weighted Results',
                'Final Weighted Results'
            ];
            
            const sheetData = {};
            for (const sheet of sheets) {
                const response = await fetch(
                    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(sheet + '!A1:Z1000')}`,
                    { headers: { 'Authorization': `Bearer ${accessToken}` } }
                );
                if (response.ok) {
                    const data = await response.json();
                    sheetData[sheet] = data.values || [];
                } else {
                    sheetData[sheet] = null;
                }
            }
            
            // Test 5: All required sheets exist and have headers
            for (const sheet of sheets) {
                const exists = sheetData[sheet] !== null && sheetData[sheet].length > 0;
                addResult(
                    `Sheet exists: ${sheet}`,
                    exists,
                    exists ? `Has ${sheetData[sheet].length - 1} data rows` : 'Sheet missing or empty'
                );
            }
            
            // Test 6: Verify headers are correct
            const votesHeaders = ['Timestamp', 'Email', 'Form', 'Category', 'Project', 'Rank', 'Points'];
            const weightedHeaders = ['Project', 'Business Impact (40%)', 'Production Readiness (40%)', 'Presentation (20%)', 'Total Score'];
            const finalHeaders = ['Project', 'Participants (40%)', 'RoW (20%, scaled 0.8)', 'Judges (40%, scaled 0.8)', 'Final Score'];
            
            for (const sheet of ['Participants Votes', 'RoW Votes', 'Judges Votes']) {
                if (sheetData[sheet] && sheetData[sheet].length > 0) {
                    const headers = sheetData[sheet][0];
                    const ok = votesHeaders.every((h, i) => headers[i] === h);
                    addResult(
                        `Headers correct: ${sheet}`,
                        ok,
                        ok ? 'All 7 vote headers present' : `Expected ${votesHeaders.join(', ')}`
                    );
                }
            }
            
            for (const sheet of ['Participants Weighted Results', 'RoW Weighted Results', 'Judges Weighted Results']) {
                if (sheetData[sheet] && sheetData[sheet].length > 0) {
                    const headers = sheetData[sheet][0];
                    const ok = weightedHeaders.every((h, i) => headers[i] === h);
                    addResult(
                        `Headers correct: ${sheet}`,
                        ok,
                        ok ? 'All 5 weighted headers present' : `Expected ${weightedHeaders.join(', ')}`
                    );
                }
            }
            
            if (sheetData['Final Weighted Results'] && sheetData['Final Weighted Results'].length > 0) {
                const headers = sheetData['Final Weighted Results'][0];
                const ok = finalHeaders.every((h, i) => headers[i] === h);
                addResult(
                    'Headers correct: Final Weighted Results',
                    ok,
                    ok ? 'All 5 final headers present (with group percentages)' : `Got: ${headers.join(', ')}`
                );
            }
            
            // Test 7: Cross-check project counts
            const getProjects = (rows) => {
                if (!rows || rows.length < 2) return new Set();
                return new Set(rows.slice(1).map(r => r[0]).filter(Boolean));
            };
            
            const pProjects = getProjects(sheetData['Participants Weighted Results']);
            const rProjects = getProjects(sheetData['RoW Weighted Results']);
            const jProjects = getProjects(sheetData['Judges Weighted Results']);
            const fProjects = getProjects(sheetData['Final Weighted Results']);
            
            // All projects in weighted results should be in final results
            const allWeightedProjects = new Set([...pProjects, ...rProjects, ...jProjects]);
            const allInFinal = [...allWeightedProjects].every(p => fProjects.has(p));
            const finalHasAll = [...fProjects].every(p => allWeightedProjects.has(p));
            
            addResult(
                'Cross-check: All weighted projects appear in Final',
                allInFinal && finalHasAll,
                `Weighted projects: ${allWeightedProjects.size}, Final projects: ${fProjects.size}`
            );
            
            // Test 8: Verify Final Weighted Results calculations
            if (sheetData['Final Weighted Results'] && sheetData['Final Weighted Results'].length > 1) {
                console.log('\n─── CALCULATION VERIFICATION ───');
                
                let calcErrors = 0;
                const finalRows = sheetData['Final Weighted Results'].slice(1);
                
                for (const row of finalRows.slice(0, 5)) { // Check first 5 projects
                    const [project, pContrib, rContrib, jContrib, total] = row;
                    const expectedTotal = parseFloat(pContrib) + parseFloat(rContrib) + parseFloat(jContrib);
                    const actualTotal = parseFloat(total);
                    
                    if (Math.abs(expectedTotal - actualTotal) > 0.01) {
                        console.log(`  ⚠️ ${project}: P(${pContrib}) + R(${rContrib}) + J(${jContrib}) = ${expectedTotal.toFixed(2)}, but got ${actualTotal}`);
                        calcErrors++;
                    }
                }
                
                addResult(
                    'Final totals = sum of contributions',
                    calcErrors === 0,
                    calcErrors === 0 ? 'All checked rows have correct totals' : `${calcErrors} calculation errors found`
                );
            }
            
            // Test 9: Verify sorting (Final should be sorted by total descending)
            if (sheetData['Final Weighted Results'] && sheetData['Final Weighted Results'].length > 2) {
                const totals = sheetData['Final Weighted Results'].slice(1).map(r => parseFloat(r[4]));
                let isSorted = true;
                for (let i = 1; i < totals.length; i++) {
                    if (totals[i] > totals[i-1]) {
                        isSorted = false;
                        break;
                    }
                }
                addResult(
                    'Final Weighted Results sorted by total (descending)',
                    isSorted,
                    isSorted ? 'Correctly sorted' : 'Not properly sorted!'
                );
            }
            
            // Test 10: Verify values are numbers not strings
            if (sheetData['Final Weighted Results'] && sheetData['Final Weighted Results'].length > 1) {
                const firstDataRow = sheetData['Final Weighted Results'][1];
                const numericCols = [1, 2, 3, 4]; // P, R, J, Total columns
                let allNumeric = true;
                
                for (const col of numericCols) {
                    const val = firstDataRow[col];
                    if (typeof val === 'string' && val.includes('.') && !isNaN(parseFloat(val))) {
                        // It's a number stored as string - this is okay from Sheets API
                    } else if (typeof val !== 'number' && isNaN(parseFloat(val))) {
                        allNumeric = false;
                    }
                }
                
                addResult(
                    'Values are numeric (not text)',
                    allNumeric,
                    allNumeric ? 'All numeric columns contain valid numbers' : 'Some columns contain non-numeric values'
                );
            }
            
            // Test 11: Vote counts consistency
            const countVotes = (rows, category) => {
                if (!rows || rows.length < 2) return 0;
                return rows.slice(1).filter(r => r[3] === category).length;
            };
            
            for (const voteSheet of ['Participants Votes', 'RoW Votes', 'Judges Votes']) {
                if (sheetData[voteSheet] && sheetData[voteSheet].length > 1) {
                    const impactCount = countVotes(sheetData[voteSheet], 'Business Impact');
                    const readinessCount = countVotes(sheetData[voteSheet], 'Production Readiness');
                    const presentationCount = countVotes(sheetData[voteSheet], 'Presentation');
                    
                    // Each category should have the same number of votes (one per project per responder)
                    const balanced = impactCount === readinessCount && readinessCount === presentationCount;
                    
                    addResult(
                        `Vote balance: ${voteSheet}`,
                        balanced,
                        `Impact=${impactCount}, Readiness=${readinessCount}, Presentation=${presentationCount}`
                    );
                }
            }
            
            // Test 12: Weighted Results columns add up (Total = Impact + Readiness + Presentation)
            console.log('\n─── COLUMN SUM VERIFICATION ───');
            
            for (const sheet of ['Participants Weighted Results', 'RoW Weighted Results', 'Judges Weighted Results']) {
                if (sheetData[sheet] && sheetData[sheet].length > 1) {
                    let errors = 0;
                    const rows = sheetData[sheet].slice(1);
                    
                    for (const row of rows) {
                        const [project, impact, readiness, presentation, total] = row;
                        const i = parseFloat(impact) || 0;
                        const r = parseFloat(readiness) || 0;
                        const p = parseFloat(presentation) || 0;
                        const t = parseFloat(total) || 0;
                        const expected = i + r + p;
                        
                        if (Math.abs(expected - t) > 0.01) {
                            console.log(`  ⚠️ ${sheet} - ${project}: ${i} + ${r} + ${p} = ${expected.toFixed(2)}, got ${t}`);
                            errors++;
                        }
                    }
                    
                    addResult(
                        `Column sums: ${sheet}`,
                        errors === 0,
                        errors === 0 ? `All ${rows.length} rows: Total = Impact + Readiness + Presentation ✓` : `${errors}/${rows.length} rows have wrong totals`
                    );
                }
            }
            
            // Test 13: Verify weighted results against raw votes (recalculate and compare)
            console.log('\n─── WEIGHTED CALCULATION VERIFICATION ───');
            
            const verifyWeightedFromVotes = (votesSheet, weightedSheet, sheetName) => {
                if (!sheetData[votesSheet] || sheetData[votesSheet].length < 2) return null;
                if (!sheetData[weightedSheet] || sheetData[weightedSheet].length < 2) return null;
                
                // Recalculate weighted results from raw votes
                const votes = sheetData[votesSheet].slice(1);
                const recalcScores = {};
                
                for (const vote of votes) {
                    const [timestamp, email, form, category, project, rank, points] = vote;
                    const pts = parseFloat(points) || 0;
                    
                    if (!recalcScores[project]) {
                        recalcScores[project] = { impact: [], readiness: [], presentation: [] };
                    }
                    
                    if (category === 'Business Impact') recalcScores[project].impact.push(pts);
                    else if (category === 'Production Readiness') recalcScores[project].readiness.push(pts);
                    else if (category === 'Presentation') recalcScores[project].presentation.push(pts);
                }
                
                // Compare with actual weighted results
                const actualWeighted = sheetData[weightedSheet].slice(1);
                let mismatches = 0;
                
                for (const row of actualWeighted) {
                    const [project, impact, readiness, presentation, total] = row;
                    const scores = recalcScores[project];
                    
                    if (!scores) {
                        console.log(`  ⚠️ ${sheetName}: Project "${project}" not found in votes!`);
                        mismatches++;
                        continue;
                    }
                    
                    // Recalculate expected values
                    const avgI = scores.impact.length > 0 ? scores.impact.reduce((a,b) => a+b, 0) / scores.impact.length : 0;
                    const avgR = scores.readiness.length > 0 ? scores.readiness.reduce((a,b) => a+b, 0) / scores.readiness.length : 0;
                    const avgP = scores.presentation.length > 0 ? scores.presentation.reduce((a,b) => a+b, 0) / scores.presentation.length : 0;
                    
                    const expectedI = avgI * 0.4;
                    const expectedR = avgR * 0.4;
                    const expectedP = avgP * 0.2;
                    const expectedT = expectedI + expectedR + expectedP;
                    
                    const actualI = parseFloat(impact) || 0;
                    const actualR = parseFloat(readiness) || 0;
                    const actualP = parseFloat(presentation) || 0;
                    const actualT = parseFloat(total) || 0;
                    
                    const tol = 0.01;
                    if (Math.abs(expectedI - actualI) > tol || 
                        Math.abs(expectedR - actualR) > tol || 
                        Math.abs(expectedP - actualP) > tol ||
                        Math.abs(expectedT - actualT) > tol) {
                        console.log(`  ⚠️ ${sheetName} - ${project}:`);
                        console.log(`     Expected: I=${expectedI.toFixed(2)}, R=${expectedR.toFixed(2)}, P=${expectedP.toFixed(2)}, T=${expectedT.toFixed(2)}`);
                        console.log(`     Actual:   I=${actualI.toFixed(2)}, R=${actualR.toFixed(2)}, P=${actualP.toFixed(2)}, T=${actualT.toFixed(2)}`);
                        mismatches++;
                    }
                }
                
                return { total: actualWeighted.length, mismatches };
            };
            
            const pVerify = verifyWeightedFromVotes('Participants Votes', 'Participants Weighted Results', 'Participants');
            if (pVerify) {
                addResult(
                    'Recalc verification: Participants Weighted',
                    pVerify.mismatches === 0,
                    pVerify.mismatches === 0 ? `All ${pVerify.total} projects match recalculated values ✓` : `${pVerify.mismatches}/${pVerify.total} mismatches`
                );
            }
            
            const rVerify = verifyWeightedFromVotes('RoW Votes', 'RoW Weighted Results', 'RoW');
            if (rVerify) {
                addResult(
                    'Recalc verification: RoW Weighted',
                    rVerify.mismatches === 0,
                    rVerify.mismatches === 0 ? `All ${rVerify.total} projects match recalculated values ✓` : `${rVerify.mismatches}/${rVerify.total} mismatches`
                );
            }
            
            const jVerify = verifyWeightedFromVotes('Judges Votes', 'Judges Weighted Results', 'Judges');
            if (jVerify) {
                addResult(
                    'Recalc verification: Judges Weighted',
                    jVerify.mismatches === 0,
                    jVerify.mismatches === 0 ? `All ${jVerify.total} projects match recalculated values ✓` : `${jVerify.mismatches}/${jVerify.total} mismatches`
                );
            }
            
            // Test 14: Verify Final Weighted Results against group weighted results
            console.log('\n─── FINAL SCORE VERIFICATION ───');
            
            if (sheetData['Final Weighted Results'] && sheetData['Final Weighted Results'].length > 1) {
                const finalRows = sheetData['Final Weighted Results'].slice(1);
                let finalErrors = 0;
                
                // Build lookup maps for weighted results
                const pMap = {}, rMap = {}, jMap = {};
                if (sheetData['Participants Weighted Results']) {
                    sheetData['Participants Weighted Results'].slice(1).forEach(row => {
                        pMap[row[0]] = parseFloat(row[4]) || 0; // Total column
                    });
                }
                if (sheetData['RoW Weighted Results']) {
                    sheetData['RoW Weighted Results'].slice(1).forEach(row => {
                        rMap[row[0]] = parseFloat(row[4]) || 0;
                    });
                }
                if (sheetData['Judges Weighted Results']) {
                    sheetData['Judges Weighted Results'].slice(1).forEach(row => {
                        jMap[row[0]] = parseFloat(row[4]) || 0;
                    });
                }
                
                for (const row of finalRows) {
                    const [project, pContrib, rContrib, jContrib, finalScore] = row;
                    
                    // Get source totals
                    const pTotal = pMap[project] || 0;
                    const rTotal = rMap[project] || 0;
                    const jTotal = jMap[project] || 0;
                    
                    // Expected contributions (with scaling for R and J)
                    const expectedP = pTotal * 0.4;
                    const expectedR = (rTotal * 0.8) * 0.2; // scaled then weighted
                    const expectedJ = (jTotal * 0.8) * 0.4; // scaled then weighted
                    const expectedFinal = expectedP + expectedR + expectedJ;
                    
                    const actualP = parseFloat(pContrib) || 0;
                    const actualR = parseFloat(rContrib) || 0;
                    const actualJ = parseFloat(jContrib) || 0;
                    const actualFinal = parseFloat(finalScore) || 0;
                    
                    const tol = 0.01;
                    if (Math.abs(expectedP - actualP) > tol || 
                        Math.abs(expectedR - actualR) > tol || 
                        Math.abs(expectedJ - actualJ) > tol ||
                        Math.abs(expectedFinal - actualFinal) > tol) {
                        console.log(`  ⚠️ Final - ${project}:`);
                        console.log(`     Source totals: P=${pTotal.toFixed(2)}, R=${rTotal.toFixed(2)}, J=${jTotal.toFixed(2)}`);
                        console.log(`     Expected: P*0.4=${expectedP.toFixed(2)}, R*0.8*0.2=${expectedR.toFixed(2)}, J*0.8*0.4=${expectedJ.toFixed(2)}, Final=${expectedFinal.toFixed(2)}`);
                        console.log(`     Actual:   P=${actualP.toFixed(2)}, R=${actualR.toFixed(2)}, J=${actualJ.toFixed(2)}, Final=${actualFinal.toFixed(2)}`);
                        finalErrors++;
                    }
                }
                
                addResult(
                    'Final scores match weighted sources (with 0.8 scaling)',
                    finalErrors === 0,
                    finalErrors === 0 ? `All ${finalRows.length} final scores correctly derived from weighted results ✓` : `${finalErrors}/${finalRows.length} mismatches`
                );
            }
            
        } catch (e) {
            addResult('Integration tests', false, `Error: ${e.message}`);
        }
    } else {
        console.log('\n⚠️ Skipping integration tests (no spreadsheet provided)');
    }
    
    // ==================== SUMMARY ====================
    console.log('\n╔══════════════════════════════════════════════════════════════╗');
    console.log(`║  RESULTS: ${results.passed} PASSED, ${results.failed} FAILED                              ║`);
    console.log('╚══════════════════════════════════════════════════════════════╝');
    
    if (results.failed > 0) {
        console.log('\n❌ FAILED TESTS:');
        results.tests.filter(t => !t.passed).forEach(t => {
            console.log(`   - ${t.name}: ${t.message}`);
        });
    }
    
    return results;
}

/**
 * Quick sanity check after generating test data
 * Verifies data was written and calculations are correct
 */
async function quickSanityCheck(accessToken, spreadsheetId) {
    console.log('\n🔍 Running quick sanity check...');
    
    const checks = [];
    
    // Check Final Weighted Results has data
    try {
        const response = await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent('Final Weighted Results!A1:E20')}`,
            { headers: { 'Authorization': `Bearer ${accessToken}` } }
        );
        
        if (response.ok) {
            const data = await response.json();
            const rows = data.values || [];
            
            if (rows.length < 2) {
                checks.push({ ok: false, msg: 'Final Weighted Results is empty!' });
            } else {
                checks.push({ ok: true, msg: `Final has ${rows.length - 1} projects` });
                
                // Verify header
                const header = rows[0];
                if (header[1] !== 'Participants (40%)' || header[2] !== 'RoW (20%, scaled 0.8)' || header[3] !== 'Judges (40%, scaled 0.8)') {
                    checks.push({ ok: false, msg: 'Headers are wrong!' });
                } else {
                    checks.push({ ok: true, msg: 'Headers are correct' });
                }
                
                // Verify first row calculation
                const firstData = rows[1];
                const p = parseFloat(firstData[1]);
                const r = parseFloat(firstData[2]);
                const j = parseFloat(firstData[3]);
                const total = parseFloat(firstData[4]);
                const expected = p + r + j;
                
                if (Math.abs(expected - total) < 0.01) {
                    checks.push({ ok: true, msg: `Calculation OK: ${p} + ${r} + ${j} = ${total}` });
                } else {
                    checks.push({ ok: false, msg: `Calculation WRONG: ${p} + ${r} + ${j} should be ${expected.toFixed(2)}, got ${total}` });
                }
            }
        } else {
            checks.push({ ok: false, msg: 'Could not read Final Weighted Results' });
        }
    } catch (e) {
        checks.push({ ok: false, msg: `Error: ${e.message}` });
    }
    
    // Print results
    const passed = checks.filter(c => c.ok).length;
    const failed = checks.filter(c => !c.ok).length;
    
    console.log(`\n📊 Quick Check: ${passed}/${checks.length} passed`);
    checks.forEach(c => console.log(`   ${c.ok ? '✅' : '❌'} ${c.msg}`));
    
    return { passed, failed, checks };
}

// Make test functions globally available
window.runSanityTests = runSanityTests;
window.quickSanityCheck = quickSanityCheck;

// Export functions for use in other scripts
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        generateVotingForms,
        generateFormTemplate,
        exportVotingConfig,
        generateVotingReport,
        addVotingButtonToUI,
        showVotingFormsModal,
        runSanityTests,
        quickSanityCheck,
        calculateWeightedResults,
        calculateFinalResults
    };
}
