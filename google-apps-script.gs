/**
 * Suivi financier — Google Drive relay (Google Apps Script)
 *
 * Deploy once in your own Google account:
 *   1. script.google.com → New project → paste this whole file (replace the default code).
 *   2. Change SECRET below to a long random string of your choice (write it down).
 *   3. Deploy → New deployment → type "Web app"
 *        - Execute as: Me
 *        - Who has access: Anyone
 *      → Authorize (Advanced → "Go to … (unsafe)" is expected: it is your own script).
 *   4. Copy the Web app URL (ends with /exec) and paste it with the SECRET in the app
 *      (Menu → Google Drive → Configuration Drive).
 *
 * After editing this code later: Deploy → Manage deployments → Edit → Version "New version"
 * (keeps the same /exec URL).
 *
 * Files are stored in a "Suivi financier" folder at the root of your Drive:
 *   - suivi_financier.json : sync source of truth (read and written by the app)
 *   - suivi_financier.xlsx : Excel mirror (written by the app, read only on manual import)
 * Updates keep the same file IDs, so Drive's version history is preserved.
 */

const SECRET = "CHANGE_ME_TO_A_LONG_RANDOM_STRING";

const FOLDER_NAME = "Suivi financier";
const JSON_NAME = "suivi_financier.json";
const XLSX_NAME = "suivi_financier.xlsx";
const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return out_({ ok: false, error: "Invalid request" });
  }
  if (SECRET === "CHANGE_ME_TO_A_LONG_RANDOM_STRING") {
    return out_({ ok: false, error: "Script not configured: change SECRET and redeploy" });
  }
  if (!req || req.key !== SECRET) return out_({ ok: false, error: "Invalid key" });

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    switch (req.action) {
      case "ping":
        return out_({ ok: true });
      case "getJson":
        return out_(getFile_(JSON_NAME, false));
      case "putJson":
        if (typeof req.content !== "string") return out_({ ok: false, error: "Missing content" });
        JSON.parse(req.content); // refuse to store invalid JSON
        return out_(putFile_(JSON_NAME, Utilities.newBlob(req.content, "application/json", JSON_NAME)));
      case "getExcel":
        return out_(getFile_(XLSX_NAME, true));
      case "putExcel":
        if (typeof req.b64 !== "string") return out_({ ok: false, error: "Missing b64" });
        return out_(putFile_(XLSX_NAME, Utilities.newBlob(Utilities.base64Decode(req.b64), XLSX_MIME, XLSX_NAME)));
      default:
        return out_({ ok: false, error: "Unknown action" });
    }
  } catch (err) {
    return out_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return out_({ ok: true, info: "Suivi financier — Drive relay is running" });
}

function folder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty("FOLDER_ID");
  if (id) {
    try {
      const f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (e) { /* folder deleted: recreate below */ }
  }
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  let f = null;
  while (it.hasNext()) { const c = it.next(); if (!c.isTrashed()) { f = c; break; } }
  if (!f) f = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty("FOLDER_ID", f.getId());
  return f;
}

function findFile_(name) {
  const it = folder_().getFilesByName(name);
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed()) return f;
  }
  return null;
}

function getFile_(name, binary) {
  const f = findFile_(name);
  if (!f) return { ok: true, exists: false };
  const blob = f.getBlob();
  const res = { ok: true, exists: true, modified: f.getLastUpdated().toISOString() };
  if (binary) res.b64 = Utilities.base64Encode(blob.getBytes());
  else res.content = blob.getDataAsString("UTF-8");
  return res;
}

function putFile_(name, blob) {
  const f = findFile_(name);
  if (!f) {
    const nf = folder_().createFile(blob);
    return { ok: true, modified: nf.getLastUpdated().toISOString() };
  }
  // Replace the content in place (same file ID → Drive version history is kept).
  const resp = UrlFetchApp.fetch(
    "https://www.googleapis.com/upload/drive/v3/files/" + f.getId() + "?uploadType=media",
    {
      method: "patch",
      contentType: blob.getContentType(),
      payload: blob.getBytes(),
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    }
  );
  if (resp.getResponseCode() < 300) return { ok: true, modified: new Date().toISOString() };
  // Fallback if the Drive API call is refused: replace the file (history of that file is lost).
  f.setTrashed(true);
  const nf = folder_().createFile(blob);
  return { ok: true, modified: nf.getLastUpdated().toISOString(), replaced: true };
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
