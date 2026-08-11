// Google Apps Script — attach to the same spreadsheet the "이미지 URL" sheet lives in.
//
// Setup:
//   1. Open that spreadsheet > Extensions > Apps Script.
//   2. Paste this whole file in, replacing the default Code.gs contents.
//   3. Add a sheet named exactly SHEET_NAME below, with header row: 브랜드 | 소재명 | 이미지 URL
//   4. Create a Drive folder for the images, share it "Anyone with the link", and
//      paste its ID (the long string in the folder's URL) into IMAGE_FOLDER_ID.
//   5. Deploy > New deployment > type "Web app" > Execute as: Me > Who has access: Anyone.
//   6. Copy the resulting /exec URL into settings.creativeUploadEndpoint in index.html,
//      and the sheet's published-CSV URL (File > Share > Publish to web > CSV) into
//      settings.creativeImageCsvUrl.
//
// The dashboard POSTs { brand, creative, imageBase64, mimeType } as a text/plain body
// (dodges a CORS preflight Apps Script doesn't answer) whenever someone drops/pastes/
// picks an image with an upload endpoint configured.

const IMAGE_FOLDER_ID = "여기에 Drive 폴더 ID";
const SHEET_NAME = "이미지URL"; // must match the sheet settings.creativeImageCsvUrl publishes

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const { brand, creative, imageBase64, mimeType } = body;
  if (!brand || !creative || !imageBase64) {
    return ContentService.createTextOutput(JSON.stringify({ error: "missing fields" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const bytes = Utilities.base64Decode(imageBase64);
  const blob = Utilities.newBlob(bytes, mimeType || "image/jpeg", brand + "_" + creative + "_" + Date.now() + ".jpg");
  const file = DriveApp.getFolderById(IMAGE_FOLDER_ID).createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = "https://drive.google.com/uc?export=view&id=" + file.getId();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  sheet.appendRow([brand, creative, url]);
  // ponytail: appends a new row per upload instead of updating one in place — the
  // dashboard's CSV reader keeps only the LAST row per (브랜드, 소재명) pair, so a
  // re-upload still "wins" on read. Prune old rows by hand if the sheet gets big, or
  // add a real find-and-update loop here if re-uploads happen a lot.

  return ContentService.createTextOutput(JSON.stringify({ url }))
    .setMimeType(ContentService.MimeType.JSON);
}
