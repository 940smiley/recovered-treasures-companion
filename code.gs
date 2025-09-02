const CONFIG = {
  MASTER_SHEET_NAME: 'MASTER',
  COL_IMAGE: 1,
  COL_BASEURL: 2,
  COL_TITLE: 3,
  COL_DESC: 4,
  COL_TAGS: 5,
  EXPORT_SHEET_NAME: 'eBayExport',
  PHOTOS_SCOPE: 'https://www.googleapis.com/auth/photoslibrary',
  SHEETS: {
    ERRORS: 'Errors',
    CATEGORIES: 'Categories',
    STAMPS: 'Stamps',
    CATAWIKI: 'catawiki',
    FACEBOOK: 'facebook',
    COLLX: 'collx',
    TCGPLAYER: 'tcgplayer',
    HOBBYDB: 'hobbydb',
    COLNECT: 'colnect'
  },
  DEFAULT_CATEGORIES: ['Stamps', 'Trading Cards']
};


function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Recovered Treasures')
    .addItem('Initialize Workbook', 'initializeWorkbook')
    .addItem('Import Photos', 'showSidebar')
    .addItem('Process All Images', 'processAllImages')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Import Photos')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function fetchRecentPhotos() {
  const token = ScriptApp.getOAuthToken();
  const url = 'https://photoslibrary.googleapis.com/v1/mediaItems?pageSize=50';

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + token
    },
    muteHttpExceptions: true
  });

  const statusCode = response.getResponseCode();
  const content = response.getContentText();

  if (statusCode === 403 && content.includes("insufficient authentication scopes")) {
    throw new Error("Missing Google Photos scope. Please reauthorize the script with the correct permissions.");
  }

  if (statusCode !== 200) {
    Logger.log(`Photos API failed with status ${statusCode}`);
    Logger.log(`Response: ${content}`);
    throw new Error(`Google Photos error: ${content}`);
  }

  const data = JSON.parse(content);
  return data.mediaItems?.map(item => ({
    id: item.id,
    baseUrl: item.baseUrl,
    filename: item.filename,
    mimeType: item.mimeType
  })) || [];
}

function importSelectedPhotosToMaster(photoArray) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
  if (!sheet) throw new Error('Master sheet not found');
  photoArray.forEach(photo => {
    const row = sheet.getLastRow() + 1;
    const imageFormula = `=IMAGE("${photo.baseUrl}=w800")`;
    sheet.getRange(row, CONFIG.COL_IMAGE).setFormula(imageFormula);
    sheet.getRange(row, CONFIG.COL_BASEURL).setValue(photo.baseUrl);
    sheet.getRange(row, CONFIG.COL_TITLE, 1, 3).clearContent();
  });
}

function saveMetadata(row, title, desc, tags) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
  sheet.getRange(row, CONFIG.COL_TITLE).setValue(title);
  sheet.getRange(row, CONFIG.COL_DESC).setValue(desc);
  sheet.getRange(row, CONFIG.COL_TAGS).setValue(tags);
}

function exportSelectedToEbay() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
  const exportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.EXPORT_SHEET_NAME) || SpreadsheetApp.getActiveSpreadsheet().insertSheet(CONFIG.EXPORT_SHEET_NAME);
  exportSheet.clear();
  const data = sheet.getDataRange().getValues();
  const selected = data.filter(row => row[0] === true); // assuming first column is checkbox
  const headers = ['Title', 'Description', 'Tags', 'Image URL'];
  exportSheet.appendRow(headers);
  selected.forEach(row => {
    exportSheet.appendRow([row[CONFIG.COL_TITLE - 1], row[CONFIG.COL_DESC - 1], row[CONFIG.COL_TAGS - 1], row[CONFIG.COL_BASEURL - 1]]);
  });
}

function initializeWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = [
    CONFIG.MASTER_SHEET_NAME,
    CONFIG.SHEETS.CATEGORIES,
    CONFIG.SHEETS.ERRORS,
    CONFIG.SHEETS.STAMPS,
    CONFIG.SHEETS.CATAWIKI,
    CONFIG.SHEETS.FACEBOOK,
    CONFIG.SHEETS.COLLX,
    CONFIG.SHEETS.TCGPLAYER,
    CONFIG.SHEETS.HOBBYDB,
    CONFIG.SHEETS.COLNECT
  ];

  requiredSheets.forEach(name => {
    if (!ss.getSheetByName(name)) {
      const newSheet = ss.insertSheet(name);
      if (name === CONFIG.SHEETS.CATEGORIES) {
        const values = CONFIG.DEFAULT_CATEGORIES.map(c => [c]);
        newSheet.getRange(1, 1, values.length, 1).setValues(values);
      }
    }
  });

  ['ebay-template', 'DONT-TOUCH'].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) sheet.protect().setWarningOnly(true);
  });
}

function processAllImages() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.MASTER_SHEET_NAME);
  const errorsSheet = ss.getSheetByName(CONFIG.SHEETS.ERRORS);
  const lastRow = sheet.getLastRow();
  for (let row = 2; row <= lastRow; row++) {
    const title = sheet.getRange(row, CONFIG.COL_TITLE).getValue();
    if (!title) {
      const baseUrl = sheet.getRange(row, CONFIG.COL_BASEURL).getValue();
      try {
        const metadata = getImageMetadata(baseUrl);
        sheet.getRange(row, CONFIG.COL_TITLE).setValue(metadata.title || '');
        sheet.getRange(row, CONFIG.COL_DESC).setValue(metadata.description || '');
        sheet.getRange(row, CONFIG.COL_TAGS).setValue(metadata.tags || '');
      } catch (err) {
        if (errorsSheet) errorsSheet.appendRow([row, err.message]);
      }
    }
  }
}

function getImageMetadata(imageUrl) {
  // Placeholder for AI metadata retrieval.
  // Replace with call to an AI service such as Google Vision API.
  return {
    title: 'Unknown Item',
    description: '',
    tags: ''
  };
}
