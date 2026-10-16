/**
 * GH Meal System - Drive backup receiver (free, no Google Cloud needed).
 * Saves the nightly backup (.xlsx) sent by the app into YOUR private Drive folder.
 * Use this as its OWN Apps Script project with its OWN Drive folder and its OWN token.
 *
 * SETUP: see README_BACKUP.md
 */
const FOLDER_ID = 'PASTE_YOUR_MEAL_BACKUP_FOLDER_ID_HERE';   // the part after /folders/ in the folder link
const TOKEN     = 'PASTE_A_LONG_RANDOM_SECRET';              // must equal GDRIVE_WEBAPP_TOKEN on the server

const FILE_RE = /^GH_MealBackup_[0-9_]+\.xlsx$/;
const MIME    = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d.token || d.token !== TOKEN) return out({ ok: false, error: 'Wrong token' });
    const folder = DriveApp.getFolderById(FOLDER_ID);

    if (d.action === 'ping') return out({ ok: true, folder: folder.getName() });

    if (d.action === 'upload') {
      if (!FILE_RE.test(d.name || '')) return out({ ok: false, error: 'Bad file name' });
      const blob = Utilities.newBlob(Utilities.base64Decode(d.data), MIME, d.name);
      const file = folder.createFile(blob);
      const removed = prune(folder, parseInt(d.keep, 10) || 0);
      return out({ ok: true, name: file.getName(), size: file.getSize(), removed: removed });
    }
    return out({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function prune(folder, keep) {
  if (keep < 1) return 0;
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (FILE_RE.test(f.getName())) files.push(f);
  }
  files.sort(function (a, b) { return a.getName() < b.getName() ? 1 : -1; }); // newest first
  let n = 0;
  for (let i = keep; i < files.length; i++) { files[i].setTrashed(true); n++; }
  return n;
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
