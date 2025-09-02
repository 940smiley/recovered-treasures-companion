const CONFIG = {
  SHEET_NAME: 'Inventory',
  COL_MEDIA_ID: 2,
  COL_TITLE: 3,
  COL_DESC: 4,
  COL_TAGS: 5,
  COL_IMAGE_URL: 6,
  EXPORT_SHEET_NAME: 'eBayExport',
  PHOTOS_SCOPE: 'https://www.googleapis.com/auth/photoslibrary'
};


function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Recovered Treasures')
    .addItem('Import Photos', 'showSidebar')
    .addItem('Reverse Search', 'reverseSearchImage')
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



function savePhotoMetadata(photo) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.appendRow([photo.filename, photo.baseUrl, photo.mimeType]);
}


function importPhotosBatch(photoArray) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  photoArray.forEach(photo => {
    sheet.appendRow([photo.filename, photo.baseUrl, photo.mimeType]);
  });
}

function reverseSearchImage(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const imageUrl = sheet.getRange(row, CONFIG.COL_IMAGE_URL).getValue();
  if (!imageUrl) throw new Error('No image URL found in row');
  const searchUrl = `https://www.google.com/searchbyimage?image_url=${encodeURIComponent(imageUrl)}`;
  const html = HtmlService.createHtmlOutput(`<script>window.open('${searchUrl}', '_blank');google.script.host.close();</script>`);
  SpreadsheetApp.getUi().showModalDialog(html, 'Reverse Image Search');
}

function saveMetadata(row, title, desc, tags) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  sheet.getRange(row, CONFIG.COL_TITLE).setValue(title);
  sheet.getRange(row, CONFIG.COL_DESC).setValue(desc);
  sheet.getRange(row, CONFIG.COL_TAGS).setValue(tags);
}

function exportSelectedToEbay() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const exportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.EXPORT_SHEET_NAME) || SpreadsheetApp.getActiveSpreadsheet().insertSheet(CONFIG.EXPORT_SHEET_NAME);
  exportSheet.clear();
  const data = sheet.getDataRange().getValues();
  const selected = data.filter(row => row[0] === true); // assuming first column is checkbox
  const headers = ['Title', 'Description', 'Tags', 'Image URL'];
  exportSheet.appendRow(headers);
  selected.forEach(row => {
    exportSheet.appendRow([row[CONFIG.COL_TITLE - 1], row[CONFIG.COL_DESC - 1], row[CONFIG.COL_TAGS - 1], row[CONFIG.COL_IMAGE_URL - 1]]);
  });
}
