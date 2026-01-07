/**
 * FolderSync.gs - Google Drive Folder Synchronization
 *
 * This module scans Google Drive for all folders and maintains a list
 * in the Folders sheet. This list is used for dropdown selection in
 * Client_Registry to specify where running meeting notes should be stored.
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

const FOLDER_SYNC_CONFIG = {
  SHEET_NAME: 'Folders',
  MAX_DEPTH: 10,  // Maximum folder depth to traverse
  BATCH_SIZE: 100 // Number of folders to process before yielding
};

// ============================================================================
// MAIN SYNC FUNCTION
// ============================================================================

/**
 * Scans all Google Drive folders and updates the Folders sheet.
 * Creates a hierarchical list showing full folder paths.
 */
function syncDriveFolders() {
  Logger.log('Starting Google Drive folder sync...');

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(FOLDER_SYNC_CONFIG.SHEET_NAME);

  // Create sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(FOLDER_SYNC_CONFIG.SHEET_NAME);
    sheet.getRange(1, 1, 1, 3).setValues([['folder_path', 'folder_id', 'folder_url']]);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Clear existing data (except header)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 3).clear();
  }

  // Collect all folders
  const folders = [];
  const rootFolder = DriveApp.getRootFolder();

  // Add root folder
  folders.push({
    path: 'My Drive',
    id: rootFolder.getId(),
    url: rootFolder.getUrl()
  });

  // Recursively collect all subfolders
  collectFolders(rootFolder, 'My Drive', folders, 0);

  // Sort folders alphabetically by path
  folders.sort((a, b) => a.path.localeCompare(b.path));

  // Write to sheet
  if (folders.length > 0) {
    const data = folders.map(f => [f.path, f.id, f.url]);
    sheet.getRange(2, 1, data.length, 3).setValues(data);
  }

  // Update data validation for Client_Registry
  updateClientRegistryFolderValidation(ss, folders);

  Logger.log(`Folder sync complete. Found ${folders.length} folders.`);

  logProcessing('FOLDER_SYNC', null, `Synced ${folders.length} folders from Drive`, 'success');
}

/**
 * Recursively collects folders from Google Drive.
 *
 * @param {Folder} parentFolder - The parent folder to scan
 * @param {string} parentPath - The path string of the parent
 * @param {Object[]} folders - Array to collect folder info
 * @param {number} depth - Current depth level
 */
function collectFolders(parentFolder, parentPath, folders, depth) {
  if (depth >= FOLDER_SYNC_CONFIG.MAX_DEPTH) {
    return;
  }

  try {
    const subFolders = parentFolder.getFolders();

    while (subFolders.hasNext()) {
      const folder = subFolders.next();
      const folderName = folder.getName();
      const folderPath = `${parentPath}/${folderName}`;

      folders.push({
        path: folderPath,
        id: folder.getId(),
        url: folder.getUrl()
      });

      // Recurse into subfolders
      collectFolders(folder, folderPath, folders, depth + 1);
    }
  } catch (error) {
    Logger.log(`Error scanning folder ${parentPath}: ${error.message}`);
  }
}

/**
 * Updates the data validation dropdown for the docs_folder column in Client_Registry.
 *
 * @param {Spreadsheet} ss - The spreadsheet object
 * @param {Object[]} folders - Array of folder objects
 */
function updateClientRegistryFolderValidation(ss, folders) {
  const clientSheet = ss.getSheetByName(CONFIG.SHEETS.CLIENT_REGISTRY);

  if (!clientSheet) {
    Logger.log('Client_Registry sheet not found');
    return;
  }

  // Find the docs_folder_path column
  const headers = clientSheet.getRange(1, 1, 1, clientSheet.getLastColumn()).getValues()[0];
  const folderColIndex = headers.indexOf('docs_folder_path');

  if (folderColIndex === -1) {
    Logger.log('docs_folder_path column not found in Client_Registry');
    return;
  }

  // Create data validation rule referencing the Folders sheet
  const folderSheet = ss.getSheetByName(FOLDER_SYNC_CONFIG.SHEET_NAME);
  const numFolders = folderSheet.getLastRow() - 1;

  if (numFolders <= 0) {
    return;
  }

  // Create validation from Folders sheet column A (folder_path)
  const validationRange = folderSheet.getRange(2, 1, numFolders, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(validationRange, true)
    .setAllowInvalid(false)
    .setHelpText('Select a folder from Google Drive')
    .build();

  // Apply to all data rows in Client_Registry (column for docs_folder_path)
  const lastRow = Math.max(clientSheet.getLastRow(), 2);
  const dataRows = lastRow - 1;

  if (dataRows > 0) {
    clientSheet.getRange(2, folderColIndex + 1, dataRows, 1).setDataValidation(rule);
  }

  // Also set validation for a large range for future rows
  clientSheet.getRange(2, folderColIndex + 1, 1000, 1).setDataValidation(rule);

  Logger.log('Updated folder dropdown validation in Client_Registry');
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Gets the folder ID for a given folder path.
 *
 * @param {string} folderPath - The folder path (e.g., "My Drive/Clients/Acme")
 * @returns {string|null} The folder ID or null if not found
 */
function getFolderIdByPath(folderPath) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(FOLDER_SYNC_CONFIG.SHEET_NAME);

  if (!sheet) {
    return null;
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === folderPath) {
      return data[i][1]; // folder_id column
    }
  }

  return null;
}

/**
 * Gets the folder path for a given folder ID.
 *
 * @param {string} folderId - The folder ID
 * @returns {string|null} The folder path or null if not found
 */
function getFolderPathById(folderId) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(FOLDER_SYNC_CONFIG.SHEET_NAME);

  if (!sheet) {
    return null;
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === folderId) {
      return data[i][0]; // folder_path column
    }
  }

  return null;
}

/**
 * Creates a new folder in Google Drive.
 *
 * @param {string} folderName - Name of the new folder
 * @param {string} parentFolderId - ID of the parent folder (optional, defaults to root)
 * @returns {Object} Object with id, path, and url of the new folder
 */
function createFolder(folderName, parentFolderId) {
  let parentFolder;

  if (parentFolderId) {
    parentFolder = DriveApp.getFolderById(parentFolderId);
  } else {
    parentFolder = DriveApp.getRootFolder();
  }

  const newFolder = parentFolder.createFolder(folderName);

  // Get the full path
  const parentPath = parentFolderId ? getFolderPathById(parentFolderId) : 'My Drive';
  const fullPath = `${parentPath}/${folderName}`;

  return {
    id: newFolder.getId(),
    path: fullPath,
    url: newFolder.getUrl()
  };
}

/**
 * Moves a file to a specific folder.
 *
 * @param {string} fileId - The file ID to move
 * @param {string} targetFolderId - The destination folder ID
 */
function moveFileToFolder(fileId, targetFolderId) {
  const file = DriveApp.getFileById(fileId);
  const targetFolder = DriveApp.getFolderById(targetFolderId);

  // Move file to target folder
  file.moveTo(targetFolder);

  Logger.log(`Moved file ${fileId} to folder ${targetFolderId}`);
}

/**
 * Gets or creates a folder by path.
 * Creates intermediate folders if they don't exist.
 *
 * @param {string} folderPath - The full folder path (e.g., "My Drive/Clients/Acme")
 * @returns {string} The folder ID
 */
function getOrCreateFolderByPath(folderPath) {
  // First check if it exists in our cache
  const existingId = getFolderIdByPath(folderPath);
  if (existingId) {
    return existingId;
  }

  // Parse the path and create folders as needed
  const parts = folderPath.split('/');
  let currentFolder = DriveApp.getRootFolder();

  // Skip "My Drive" if it's the first part
  const startIndex = parts[0] === 'My Drive' ? 1 : 0;

  for (let i = startIndex; i < parts.length; i++) {
    const folderName = parts[i];
    const subFolders = currentFolder.getFoldersByName(folderName);

    if (subFolders.hasNext()) {
      currentFolder = subFolders.next();
    } else {
      // Create the folder
      currentFolder = currentFolder.createFolder(folderName);
      Logger.log(`Created folder: ${folderName}`);
    }
  }

  return currentFolder.getId();
}

// ============================================================================
// SCHEDULED SYNC
// ============================================================================

/**
 * Handler for scheduled folder sync.
 * Called by trigger (e.g., daily).
 */
function runFolderSync() {
  try {
    logProcessing('FOLDER_SYNC', null, 'Starting scheduled folder sync', 'processing');
    syncDriveFolders();
  } catch (error) {
    logProcessing('FOLDER_SYNC', null, `Error: ${error.message}`, 'error');
    Logger.log(`Folder sync error: ${error.message}`);
  }
}
