/**
 * Park Cleaners PDF Saver.
 *
 * The invoice app runs as whoever is using it, so any file it creates is
 * owned by - and counts against the Drive storage of - that person. Staff
 * accounts with full storage then can't save invoices ("Limit Exceeded:
 * Drive"). This small web app runs as the admin instead, and saves, trashes
 * and restores PDFs in the Invoices folder on the invoice app's behalf.
 *
 * Requests must carry SAVER_KEY (Config.js, kept out of git). The first
 * valid request fixes which folder is the Invoices folder; after that only
 * files in it, or in folders directly inside it, can be touched.
 */

function doPost(e) {
  let result;
  try {
    const req = JSON.parse(e.postData.contents);
    if (!SAVER_KEY || req.key !== SAVER_KEY) throw new Error("Not authorized.");
    const root = invoiceRoot_(req.rootFolderId);

    if (req.action === 'save') {
      const folder = DriveApp.getFolderById(req.folderId || root.getId());
      if (!isRootOrChild_(root, folder)) throw new Error("That folder is outside the Invoices folder.");
      if (req.cleanupHours) trashOldFiles_(folder, req.cleanupHours);

      const blob = Utilities.newBlob(Utilities.base64Decode(req.pdfBase64), 'application/pdf', String(req.fileName));
      const file = folder.createFile(blob);
      if (req.makePublic) {
        try {
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (err) {
          console.warn("Couldn't make " + file.getName() + " public: " + err.message);
        }
      }
      result = { ok: true, id: file.getId(), url: file.getUrl() };

    } else if (req.action === 'trash' || req.action === 'untrash') {
      const file = DriveApp.getFileById(req.fileId);
      if (!isInInvoices_(root, file)) throw new Error("That file is outside the Invoices folder.");
      file.setTrashed(req.action === 'trash');
      result = { ok: true };

    } else {
      throw new Error("Unknown action: " + req.action);
    }
  } catch (err) {
    result = { ok: false, error: err && err.message || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function invoiceRoot_(requestedId) {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('INVOICE_FOLDER_ID');
  if (!id) {
    if (!requestedId) throw new Error("The Invoices folder isn't known yet.");
    id = DriveApp.getFolderById(requestedId).getId();
    props.setProperty('INVOICE_FOLDER_ID', id);
  }
  return DriveApp.getFolderById(id);
}

function isRootOrChild_(root, folder) {
  if (folder.getId() === root.getId()) return true;
  const parents = folder.getParents();
  while (parents.hasNext()) {
    if (parents.next().getId() === root.getId()) return true;
  }
  return false;
}

function isInInvoices_(root, file) {
  const parents = file.getParents();
  while (parents.hasNext()) {
    if (isRootOrChild_(root, parents.next())) return true;
  }
  return false;
}

function trashOldFiles_(folder, hours) {
  const cutoff = Date.now() - Number(hours) * 60 * 60 * 1000;
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getDateCreated().getTime() < cutoff) {
      try { f.setTrashed(true); } catch (e) { /* not ours to trash; leave it */ }
    }
  }
}

/** Run once from the editor so Google asks the admin to approve Drive access. */
function authorize() {
  DriveApp.getRootFolder();
  console.log("Authorized. The PDF saver can now save invoices.");
}
