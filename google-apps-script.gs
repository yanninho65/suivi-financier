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
 *
 * Light check (getHash): on each putJson the app sends its own hash of the data; the script stores it
 * (with the MD5 of the stored content) so the app can check for changes without downloading the JSON.
 * If the file was changed outside the app, the MD5 no longer matches and getHash returns hash:null
 * (the app then falls back to a full download).
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
      case "getHash":
        return out_(getHash_());
      case "putJson": {
        if (typeof req.content !== "string") return out_({ ok: false, error: "Missing content" });
        JSON.parse(req.content); // refuse to store invalid JSON
        const res = putFile_(JSON_NAME, Utilities.newBlob(req.content, "application/json", JSON_NAME));
        rememberHash_(req.content, req.hash);
        return out_(res);
      }
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

function md5Hex_(str) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ("0" + (b & 0xff).toString(16)).slice(-2); }).join("");
}

// Stores the app's hash of the JSON just written, tied to the MD5 of that exact content.
function rememberHash_(content, hash) {
  const props = PropertiesService.getScriptProperties();
  if (typeof hash === "string" && hash) {
    props.setProperties({ JSON_HASH: hash, JSON_MD5: md5Hex_(content) });
  } else {
    props.deleteProperty("JSON_HASH");
    props.deleteProperty("JSON_MD5");
  }
}

// Returns the stored hash only if the file on Drive is still exactly the content written by the app.
function getHash_() {
  const f = findFile_(JSON_NAME);
  if (!f) return { ok: true, exists: false };
  const props = PropertiesService.getScriptProperties();
  const hash = props.getProperty("JSON_HASH");
  const md5 = props.getProperty("JSON_MD5");
  if (!hash || !md5) return { ok: true, exists: true, hash: null };
  const same = md5Hex_(f.getBlob().getDataAsString("UTF-8")) === md5;
  return { ok: true, exists: true, hash: same ? hash : null, modified: f.getLastUpdated().toISOString() };
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
