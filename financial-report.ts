//This script has been made by a worker in the cooperation and development world in order to make his life easier with financial reports. You can reach him on Telegram at @mortacci. You can use this code as you wish, and if you improve it you can find the repository at https://github.com/Taccimor/financial-report and commit changes, just keep this line in order to preserve the possibility to have contacts.
function downloadFiles() {
  const sheet = SpreadsheetApp.getActiveSheet();
  SpreadsheetApp.flush();
  const data = sheet.getDataRange().getValues();
  
  // ===== CONFIGURATION =====
  const ROOT_FOLDER_URL = "URL OF YOUR FOLDER"; // Replace with the URL of the folder where the files will be copied
  const START_ROW = 2; // Set starting row number here. Be careful! It's the number of the row, not the number of the expenditure!
  const folderNameCol = 0; // Column that contains the name of the folders (i.e. the progressive number, usually). 0 = A, 1 = B, and so on...
  const parentFolderCol = 3; // Column that contains the budget line code. 0 = A, 1 = B, and so on...
  const urlCol = 8; // Column that contains the URL of the files. 0 = A, 1 = B, and so on...
  // =========================
  
  const defaultParentName = "No budget line";
  const startIndex = Math.max(1, START_ROW - 1); // -1 to convert to 0-based
  const rootFolderId = ROOT_FOLDER_URL.match(/[-\w]{25,}/)[0];
  const rootFolder = DriveApp.getFolderById(rootFolderId);

  for (let i = startIndex; i < data.length; i++) {
    const row = data[i];
    const folderName = row[folderNameCol].toString().trim();
    let parentFolderName = row[parentFolderCol] ? row[parentFolderCol].toString().trim() : "";
    const url = row[urlCol].toString().trim();

    // Skip rows with missing essential values
    if (!folderName || !url) {
      Logger.log(`Skipping row ${i+1}: Missing progressive number or URL`);
      continue;
    }

    // Use default if parent folder name is empty
    if (!parentFolderName) {
      parentFolderName = defaultParentName;
    }

    const match = url.match(/[-\w]{25,}/);
    if (!match) {
      Logger.log(`Skipping row ${i+1} (expenditure n.${folderName}): Invalid URL format`);
      continue;
    }
    const id = match[0];
    
    try {
      // Create parent folder
      let parentFolder;
      try {
        const parentIterator = rootFolder.getFoldersByName(parentFolderName);
        if (parentIterator.hasNext()) {
          parentFolder = parentIterator.next();
        } else {
          parentFolder = rootFolder.createFolder(parentFolderName);
          Logger.log(`Created parent folder: ${parentFolderName}`);
        }
      } catch (e) {
        Logger.log(`Error creating parent folder '${parentFolderName}': ${e.toString()}`);
        continue;
      }
      
      // Create target folder
      let targetFolder;
      try {
        const targetIterator = parentFolder.getFoldersByName(folderName);
        if (targetIterator.hasNext()) {
          targetFolder = targetIterator.next();
        } else {
          targetFolder = parentFolder.createFolder(folderName);
          Logger.log(`Created target folder: ${folderName} in ${parentFolderName}`);
        }
      } catch (e) {
        Logger.log(`Error creating target folder '${folderName}': ${e.toString()}`);
        continue;
      }

      // Try to process as PDF file first
      try {
        const file = DriveApp.getFileById(id);
        
        // Check if it's a PDF
        if (file.getMimeType() === "application/pdf") {
          file.makeCopy(file.getName(), targetFolder);
          Logger.log(`Copied PDF file: ${file.getName()} to ${parentFolderName}/${folderName}`);
        } else {
          // If not PDF, try to process as folder instead
          throw new Error("Not a PDF file");
        }
      } catch (fileError) {
        // If file processing fails, try as folder
        try {
          const sourceFolder = DriveApp.getFolderById(id);
          copyFolder(sourceFolder, targetFolder);
          Logger.log(`Copied folder contents: ${sourceFolder.getName()} to ${parentFolderName}/${folderName}`);
        } catch (folderError) {
          Logger.log(`Skipping row ${i+1} (expenditure n.${folderName}): Resource is not a PDF file or a valid folder`);
        }
      }
      
    } catch (error) {
      Logger.log(`Error processing row ${i+1} (expenditure n.${folderName}): ${error.toString()}`);
    }
  }
  Logger.log("Financial report is now ready!");
}

// Recursive function to copy folder structure with PDFs only
function copyFolder(sourceFolder, targetFolder) {
  // Copy PDF files in current folder
  const files = sourceFolder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() === "application/pdf") {
      file.makeCopy(targetFolder);
    }
  }
  
  // Process subfolders recursively
  const subfolders = sourceFolder.getFolders();
  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    const newSubfolder = targetFolder.createFolder(subfolder.getName());
    copyFolder(subfolder, newSubfolder);
  }
}

function runThis() {
  downloadFiles();
}