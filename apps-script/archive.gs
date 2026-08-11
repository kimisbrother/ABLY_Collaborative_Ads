// Google Apps Script — one shared backend for two dashboard features:
//   1. Creative image uploads (Drive storage + a "브랜드/소재명/이미지 URL" sheet)
//   2. Weekly comment tabs (a "브랜드/제목/코멘트" sheet)
// Both go through the same /exec URL — the dashboard tags each POST with a "type"
// field and this file routes on it.
//
// Setup:
//   1. Open the spreadsheet that already publishes the login/data CSVs (or a new one,
//      doesn't matter) > Extensions > Apps Script.
//   2. Paste this whole file in, replacing the default Code.gs contents.
//   3. Add two sheets, named exactly as below:
//        IMAGE_SHEET_NAME   header row: 브랜드 | 소재명 | 이미지 URL
//        COMMENT_SHEET_NAME header row: 브랜드 | 제목 | 코멘트 | 시각
//   4. Create a Drive folder for creative images, share it "Anyone with the link", and
//      paste its ID (the long string in the folder's URL) into IMAGE_FOLDER_ID below.
//   5. Deploy > New deployment > type "Web app" > Execute as: Me > Who has access: Anyone.
//   6. Copy the resulting /exec URL into settings.appsScriptEndpoint in index.html.
//   7. Publish each sheet to CSV (File > Share > Publish to web > choose the sheet >
//      CSV) and paste those two links into settings.creativeImageCsvUrl and
//      settings.commentsCsvUrl in index.html.
//
// CORS note: the dashboard sends "Content-Type: text/plain" on purpose — that's a
// "simple request" browsers don't preflight, and Apps Script Web Apps have no
// doOptions to answer a preflight with anyway. e.postData.contents is still the raw
// JSON string regardless of that header, which is why it works.

const IMAGE_FOLDER_ID = "여기에 Drive 폴더 ID";
const IMAGE_SHEET_NAME = "이미지URL";
const COMMENT_SHEET_NAME = "코멘트";

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const result = body.type === "comment" ? saveComment(body) : saveImage(body);
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function saveImage(body) {
  const { brand, creative, imageBase64, mimeType } = body;
  if (!brand || !creative || !imageBase64) return { error: "missing fields" };

  const bytes = Utilities.base64Decode(imageBase64);
  const blob = Utilities.newBlob(bytes, mimeType || "image/jpeg", brand + "_" + creative + "_" + Date.now() + ".jpg");
  const file = DriveApp.getFolderById(IMAGE_FOLDER_ID).createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = "https://drive.google.com/uc?export=view&id=" + file.getId();

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(IMAGE_SHEET_NAME).appendRow([brand, creative, url]);
  // ponytail: appends a new row per upload instead of updating one in place — the
  // dashboard's CSV reader keeps only the LAST row per (브랜드, 소재명) pair, so a
  // re-upload still "wins" on read. Prune old rows by hand if the sheet gets big.
  return { url };
}

function saveComment(body) {
  const { brand, title, body: text } = body;
  if (!brand || !title) return { error: "missing fields" };
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(COMMENT_SHEET_NAME).appendRow([brand, title, text || "", new Date()]);
  // ponytail: same append-only scheme as images — editing a tab and saving again just
  // appends another row; the reader keeps the latest body per (브랜드, 제목).
  return { ok: true };
}
