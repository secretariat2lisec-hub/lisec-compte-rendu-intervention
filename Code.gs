const CONFIG = {
  spreadsheetName: "LISEC - Comptes rendus interventions",
  driveRootFolderName: "LISEC - Comptes rendus interventions",
  emailTo: "secretariat2.lisec@gmail.com",
  templateDocId: "",
  wordTemplateUrl: "https://raw.githubusercontent.com/secretariat2lisec-hub/lisec-compte-rendu-intervention/main/Masque-NOTE-TECHNIQUE-AUTOMATIQUE.docx",
  uploadRootFolderName: "_Transferts photos en cours",
  uploadManifestFileName: "upload-manifest.json"
};

const DOSSIER_FILE_NAME = "dossier.json";
const DOSSIER_HEADERS = [
  "ID intervention",
  "Statut",
  "Date visite",
  "Client",
  "Lieu",
  "Ingenieur",
  "Derniere modification",
  "ID dossier Drive",
  "Rapport"
];

function doPost(e) {
  const properties = PropertiesService.getScriptProperties();
  let payload = {};
  try {
    payload = JSON.parse(e.postData.contents || "{}");
    let result;
    if (payload.requestType === "photo-batch") {
      result = handlePhotoBatch_(payload);
    } else if (payload.requestType === "discard-photo-upload") {
      result = discardPhotoUpload_(payload.photoUploadSessionId);
    } else {
      result = handleSubmission(payload);
    }
    properties.setProperty("LAST_SUBMISSION_DIAGNOSTIC", JSON.stringify({
      ok: true,
      at: new Date().toISOString(),
      interventionId: result.interventionId || "",
      workflowAction: result.workflowAction || "",
      requestType: payload.requestType || "submission"
    }));
    return jsonResponse({ ok: true, result });
  } catch (error) {
    const detail = String(error && error.stack ? error.stack : error);
    console.error(detail);
    markTransferError_(payload.photoUploadSessionId, detail);
    properties.setProperty("LAST_SUBMISSION_DIAGNOSTIC", JSON.stringify({
      ok: false,
      at: new Date().toISOString(),
      error: detail
    }));
    return jsonResponse({ ok: false, error: String(error && error.message ? error.message : error) });
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.api === "transfer-status") {
    const status = getTransferStatus_(e.parameter.photoUploadSessionId || "");
    return javascriptResponse_(e.parameter.callback || "", status);
  }

  if (e && e.parameter && e.parameter.api === "last-submission") {
    assertSecretariatAccess_({ accessCode: e.parameter.accessCode || "" });
    const raw = PropertiesService.getScriptProperties().getProperty("LAST_SUBMISSION_DIAGNOSTIC");
    return jsonResponse(raw ? JSON.parse(raw) : { ok: false, error: "Aucun diagnostic disponible." });
  }

  if (e && e.parameter && e.parameter.api === "status") {
    return jsonResponse({ ok: true, message: "LISEC Apps Script pret a recevoir les comptes rendus." });
  }

  return HtmlService
    .createHtmlOutputFromFile("Secretariat")
    .setTitle("LISEC - Secretariat")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function handleSubmission(payload) {
  const workflowAction = payload.workflowAction === "completion" ? "completion" : "report";
  const interventionId = makeInterventionId(payload);
  const rootFolder = getOrCreateFolder_(CONFIG.driveRootFolderName);
  const interventionFolder = getOrCreateSubFolder_(rootFolder, interventionId);
  const photoFolder = getOrCreateSubFolder_(interventionFolder, "Photos");

  const photoRecords = payload.photoUploadSessionId
    ? adoptUploadedPhotos_(payload, photoFolder)
    : savePhotos_(payload, interventionId, photoFolder);
  const expectedPhotoCount = Number(payload.photoCount || 0);
  if (payload.photoUploadSessionId && photoRecords.length !== expectedPhotoCount) {
    throw new Error("Transfert photo incomplet : " + photoRecords.length + " photo(s) reçue(s) sur " + expectedPhotoCount + ".");
  }
  syncPhotoMetadata_(payload, photoRecords);
  const spreadsheet = getOrCreateSpreadsheet_();

  writeDatabaseSheets_(spreadsheet, payload, interventionId, photoRecords);
  writeInterventionSheet_(spreadsheet, payload, interventionId, photoRecords);

  let dossier = saveDossierState_(
    payload,
    interventionId,
    interventionFolder,
    photoRecords,
    workflowAction === "completion" ? "A_COMPLETER" : "GENERATION_EN_COURS",
    ""
  );
  upsertDossierIndex_(spreadsheet, dossier);

  if (workflowAction === "completion") {
    sendCompletionEmail_(payload, interventionId, interventionFolder, spreadsheet, photoRecords);
    const result = {
      interventionId,
      workflowAction,
      spreadsheetUrl: spreadsheet.getUrl(),
      folderUrl: interventionFolder.getUrl(),
      reportUrl: "",
      photoCount: photoRecords.length
    };
    markTransferFinalized_(payload.photoUploadSessionId, result);
    return result;
  }

  const reportFile = createWordReport_(payload, interventionId, interventionFolder, photoRecords);
  const docxBlob = reportFile.getBlob().setName(interventionId + ".docx");

  sendSummaryEmail_(payload, interventionId, reportFile, docxBlob, photoRecords);
  dossier = saveDossierState_(
    payload,
    interventionId,
    interventionFolder,
    photoRecords,
    "RAPPORT_GENERE",
    reportFile.getUrl()
  );
  upsertDossierIndex_(spreadsheet, dossier);

  const result = {
    interventionId,
    workflowAction,
    spreadsheetUrl: spreadsheet.getUrl(),
    folderUrl: interventionFolder.getUrl(),
    reportUrl: reportFile.getUrl(),
    photoCount: photoRecords.length
  };
  markTransferFinalized_(payload.photoUploadSessionId, result);
  return result;
}

function listInterventionsForSecretariat(request) {
  assertSecretariatAccess_(request);
  const spreadsheet = getOrCreateSpreadsheet_();
  const sheet = ensureDossierIndexSheet_(spreadsheet);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  return values.slice(1).map((row) => ({
    interventionId: String(row[0] || ""),
    status: String(row[1] || "A_COMPLETER"),
    visitDate: formatSheetDate_(row[2]),
    client: String(row[3] || ""),
    location: String(row[4] || ""),
    engineer: String(row[5] || ""),
    updatedAt: formatSheetDateTime_(row[6]),
    folderId: String(row[7] || ""),
    reportUrl: String(row[8] || "")
  })).filter((item) => item.interventionId).sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

function getInterventionForSecretariat(request) {
  assertSecretariatAccess_(request);
  const interventionId = request && request.interventionId;
  const folder = findDossierFolder_(interventionId);
  const dossier = readDossierState_(folder);
  if (!dossier) throw new Error("Dossier introuvable : " + interventionId);
  return dossier;
}

function saveInterventionForSecretariat(request) {
  assertSecretariatAccess_(request);
  if (!request || !request.interventionId || !request.payload) {
    throw new Error("Dossier incomplet.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const folder = findDossierFolder_(request.interventionId);
    const existing = readDossierState_(folder);
    if (!existing) throw new Error("Dossier introuvable : " + request.interventionId);

    syncPhotoMetadata_(request.payload, existing.photos || []);
    existing.payload = sanitizePayload_(request.payload, existing.photos || []);
    existing.payload.workflowAction = "completion";
    existing.status = "EN_REVISION";
    existing.updatedAt = new Date().toISOString();
    writeDossierState_(folder, existing);
    upsertDossierIndex_(getOrCreateSpreadsheet_(), existing);
    return existing;
  } finally {
    lock.releaseLock();
  }
}

function generateReportForSecretariat(request) {
  assertSecretariatAccess_(request);
  if (!request || !request.interventionId) throw new Error("Dossier non selectionne.");
  if (request.payload) saveInterventionForSecretariat(request);

  const interventionId = request.interventionId;
  const folder = findDossierFolder_(interventionId);
  const dossier = readDossierState_(folder);
  if (!dossier) throw new Error("Dossier introuvable : " + interventionId);

  const payload = dossier.payload || {};
  payload.workflowAction = "report";
  const photoRecords = loadPhotoRecords_(dossier.photos || []);
  const reportFile = createWordReport_(payload, interventionId, folder, photoRecords);
  const docxBlob = reportFile.getBlob().setName(interventionId + ".docx");
  sendSummaryEmail_(payload, interventionId, reportFile, docxBlob, photoRecords);

  dossier.payload = sanitizePayload_(payload, dossier.photos || []);
  dossier.status = "RAPPORT_GENERE";
  dossier.reportUrl = reportFile.getUrl();
  dossier.updatedAt = new Date().toISOString();
  dossier.generatedAt = dossier.updatedAt;
  writeDossierState_(folder, dossier);
  upsertDossierIndex_(getOrCreateSpreadsheet_(), dossier);

  return {
    interventionId,
    status: dossier.status,
    reportUrl: dossier.reportUrl,
    updatedAt: dossier.updatedAt
  };
}

function assertSecretariatAccess_(request) {
  const expected = PropertiesService.getScriptProperties().getProperty("SECRETARIAT_ACCESS_CODE");
  if (!expected) {
    throw new Error("Le code d'acces du secretariat n'est pas encore configure dans les proprietes du script.");
  }
  if (!request || String(request.accessCode || "") !== String(expected)) {
    throw new Error("Code d'acces incorrect.");
  }
}

function syncPhotoMetadata_(payload, photos) {
  (payload.levels || []).forEach((level) => {
    (level.entries || []).forEach((entry) => {
      const key = entryKey_(level.name, entry);
      (photos || []).forEach((photo) => {
        if (!photo.isSitePhoto && photo.entryKey === key) {
          photo.levelName = level.name || "";
          photo.localisation = entry.localisation || "";
        }
      });
    });
  });

  const sitePhoto = (photos || []).find((photo) => photo.isSitePhoto);
  if (sitePhoto) sitePhoto.localisation = fullSiteAddress_(payload);
}

function saveDossierState_(payload, interventionId, folder, photoRecords, status, reportUrl) {
  const photos = photoRecordsMetadata_(photoRecords);
  const dossier = {
    interventionId,
    status,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    folderId: folder.getId(),
    folderUrl: folder.getUrl(),
    reportUrl: reportUrl || "",
    payload: sanitizePayload_(payload, photos),
    photos
  };

  const previous = readDossierState_(folder);
  if (previous && previous.createdAt) dossier.createdAt = previous.createdAt;
  writeDossierState_(folder, dossier);
  return dossier;
}

function sanitizePayload_(payload, photos) {
  const source = payload || {};
  const copy = Object.assign({}, source);
  const sitePhoto = (photos || []).find((photo) => photo.isSitePhoto);
  copy.sitePhoto = sitePhoto ? photoForInterface_(sitePhoto) : null;
  copy.presentPeopleEntries = Array.isArray(source.presentPeopleEntries) ? source.presentPeopleEntries.slice(0, 4) : [];
  copy.constructionItems = Array.isArray(source.constructionItems) ? source.constructionItems.slice(0, 6) : [];
  copy.diffusionEntries = Array.isArray(source.diffusionEntries) ? source.diffusionEntries.slice(0, 3) : [];
  copy.levels = (source.levels || []).map((level) => ({
    id: level.id || "",
    name: level.name || "",
    entries: (level.entries || []).map((entry) => {
      const key = entryKey_(level.name, entry);
      return {
        id: entry.id || "",
        localisation: entry.localisation || "",
        comment: entry.comment || "",
        gravity: entry.gravity || "",
        photos: (photos || [])
          .filter((photo) => !photo.isSitePhoto && photo.entryKey === key)
          .map(photoForInterface_)
      };
    })
  }));
  return copy;
}

function photoRecordsMetadata_(photoRecords) {
  return (photoRecords || []).map((photo) => ({
    levelName: photo.levelName || "",
    localisation: photo.localisation || "",
    entryKey: photo.entryKey || "",
    isSitePhoto: Boolean(photo.isSitePhoto),
    name: photo.name || photo.fileName || "photo.jpg",
    fileName: photo.fileName || photo.name || "photo.jpg",
    url: photo.url || "",
    fileId: photo.fileId || "",
    thumbnailUrl: photo.fileId ? `https://drive.google.com/thumbnail?id=${photo.fileId}&sz=w800` : ""
  }));
}

function photoForInterface_(photo) {
  return {
    name: photo.name || photo.fileName || "photo.jpg",
    url: photo.url || "",
    fileId: photo.fileId || "",
    thumbnailUrl: photo.thumbnailUrl || (photo.fileId ? `https://drive.google.com/thumbnail?id=${photo.fileId}&sz=w800` : "")
  };
}

function loadPhotoRecords_(photos) {
  return (photos || []).map((photo) => {
    try {
      const file = DriveApp.getFileById(photo.fileId);
      return Object.assign({}, photo, { blob: file.getBlob() });
    } catch (error) {
      return null;
    }
  }).filter(Boolean);
}

function writeDossierState_(folder, dossier) {
  const content = JSON.stringify(dossier);
  const files = folder.getFilesByName(DOSSIER_FILE_NAME);
  if (files.hasNext()) {
    files.next().setContent(content);
  } else {
    folder.createFile(DOSSIER_FILE_NAME, content, MimeType.PLAIN_TEXT);
  }
}

function readDossierState_(folder) {
  const files = folder.getFilesByName(DOSSIER_FILE_NAME);
  if (!files.hasNext()) return null;
  return JSON.parse(files.next().getBlob().getDataAsString("UTF-8"));
}

function findDossierFolder_(interventionId) {
  const spreadsheet = getOrCreateSpreadsheet_();
  const sheet = ensureDossierIndexSheet_(spreadsheet);
  const values = sheet.getDataRange().getValues();
  for (let row = 1; row < values.length; row++) {
    if (String(values[row][0]) === String(interventionId) && values[row][7]) {
      return DriveApp.getFolderById(String(values[row][7]));
    }
  }

  const root = getOrCreateFolder_(CONFIG.driveRootFolderName);
  const folders = root.getFoldersByName(interventionId);
  if (folders.hasNext()) return folders.next();
  throw new Error("Dossier Drive introuvable : " + interventionId);
}

function ensureDossierIndexSheet_(spreadsheet) {
  return ensureSheet_(spreadsheet, "Dossiers", DOSSIER_HEADERS);
}

function upsertDossierIndex_(spreadsheet, dossier) {
  const sheet = ensureDossierIndexSheet_(spreadsheet);
  const values = sheet.getDataRange().getValues();
  const payload = dossier.payload || {};
  const rowValues = [[
    dossier.interventionId,
    dossier.status || "A_COMPLETER",
    payload.visitDate || "",
    clientText_(payload),
    fullSiteAddress_(payload),
    payload.engineer || "",
    dossier.updatedAt || new Date().toISOString(),
    dossier.folderId || "",
    dossier.reportUrl || ""
  ]];

  let targetRow = 0;
  for (let row = 1; row < values.length; row++) {
    if (String(values[row][0]) === String(dossier.interventionId)) {
      targetRow = row + 1;
      break;
    }
  }

  if (targetRow) {
    sheet.getRange(targetRow, 1, 1, DOSSIER_HEADERS.length).setValues(rowValues);
  } else {
    sheet.appendRow(rowValues[0]);
  }
}

function formatSheetDate_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(value);
}

function formatSheetDateTime_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  }
  return String(value);
}

function makeInterventionId(payload) {
  const supplied = cleanName_(payload.interventionId || "").substring(0, 120);
  if (supplied) return supplied;
  const datePart = payload.visitDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const timePart = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HHmmss");
  const engineer = cleanName_(payload.engineer || "ingenieur");
  return `${datePart}-${timePart}-${engineer}`;
}

function getOrCreateSpreadsheet_() {
  const files = DriveApp.getFilesByName(CONFIG.spreadsheetName);
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      return SpreadsheetApp.openById(file.getId());
    }
  }

  const spreadsheet = SpreadsheetApp.create(CONFIG.spreadsheetName);
  setupDatabaseSheets_(spreadsheet);
  return spreadsheet;
}

function setupDatabaseSheets_(spreadsheet) {
  const first = spreadsheet.getSheets()[0];
  first.setName("Interventions");
  first.appendRow([
    "ID intervention",
    "Date reception",
    "Date visite",
    "Heure visite",
    "Ingenieur",
    "Client",
    "Adresse",
    "Code postal",
    "Ville",
    "Personnes presentes",
    "Mission",
    "Diffusion",
    "Rapport",
    "Nombre photos",
    "Reference LISEC",
    "Titre"
  ]);

  const observations = spreadsheet.insertSheet("Observations");
  observations.appendRow([
    "ID intervention",
    "Niveau",
    "Localisation",
    "Gravite",
    "Commentaire"
  ]);

  const photos = spreadsheet.insertSheet("Photos");
  photos.appendRow([
    "ID intervention",
    "Niveau",
    "Localisation",
    "Nom photo",
    "Lien Drive"
  ]);

  const dossiers = spreadsheet.insertSheet("Dossiers");
  dossiers.appendRow(DOSSIER_HEADERS);
}

function ensureSheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
    sheet.appendRow(headers);
  } else if (headers && sheet.getLastColumn() < headers.length) {
    const firstMissingColumn = sheet.getLastColumn() + 1;
    const missingHeaders = headers.slice(firstMissingColumn - 1);
    sheet.getRange(1, firstMissingColumn, 1, missingHeaders.length).setValues([missingHeaders]);
  }
  return sheet;
}

function writeDatabaseSheets_(spreadsheet, payload, interventionId, photoRecords) {
  const interventions = ensureSheet_(spreadsheet, "Interventions", [
    "ID intervention", "Date reception", "Date visite", "Heure visite", "Ingenieur", "Client", "Adresse", "Code postal", "Ville", "Personnes presentes", "Mission", "Diffusion", "Rapport", "Nombre photos", "Reference LISEC", "Titre"
  ]);
  const observations = ensureSheet_(spreadsheet, "Observations", [
    "ID intervention", "Niveau", "Localisation", "Gravite", "Commentaire"
  ]);
  const photos = ensureSheet_(spreadsheet, "Photos", [
    "ID intervention", "Niveau", "Localisation", "Nom photo", "Lien Drive"
  ]);

  interventions.appendRow([
    interventionId,
    new Date(),
    payload.visitDate || "",
    payload.visitTime || "",
    payload.engineer || "",
    clientText_(payload),
    payload.siteAddress || "",
    payload.postalCode || "",
    payload.city || "",
    presentPeopleText_(payload),
    payload.mission || "",
    diffusionText_(payload),
    "",
    photoRecords.length,
    payload.lisecReference || "",
    payload.reportTitle || ""
  ]);

  (payload.levels || []).forEach((level) => {
    (level.entries || []).forEach((entry) => {
      observations.appendRow([
        interventionId,
        level.name || "",
        entry.localisation || "",
        entry.gravity || "",
        entry.comment || ""
      ]);
    });
  });

  photoRecords.forEach((photo) => {
    photos.appendRow([
      interventionId,
      photo.levelName,
      photo.localisation,
      photo.name,
      photo.url
    ]);
  });
}

function writeInterventionSheet_(spreadsheet, payload, interventionId, photoRecords) {
  const sheetName = uniqueSheetName_(spreadsheet, interventionId.substring(0, 90));
  const sheet = spreadsheet.insertSheet(sheetName);

  sheet.appendRow(["Compte rendu LISEC", interventionId]);
  sheet.appendRow(["Traitement demande", payload.workflowAction === "completion" ? "A completer par le secretariat" : "Generer le rapport"]);
  sheet.appendRow(["Date de visite", payload.visitDate || ""]);
  sheet.appendRow(["Heure de visite", payload.visitTime || ""]);
  sheet.appendRow(["Ingenieur", payload.engineer || ""]);
  sheet.appendRow(["Client", clientText_(payload)]);
  sheet.appendRow(["Reference LISEC", payload.lisecReference || ""]);
  sheet.appendRow(["Titre", payload.reportTitle || ""]);
  sheet.appendRow(["Adresse", payload.siteAddress || ""]);
  sheet.appendRow(["Code postal", payload.postalCode || ""]);
  sheet.appendRow(["Ville", payload.city || ""]);
  sheet.appendRow(["Personnes presentes", presentPeopleText_(payload)]);
  sheet.appendRow(["Mission", payload.mission || ""]);
  sheet.appendRow(["Diffusion", diffusionText_(payload)]);
  sheet.appendRow([""]);
  sheet.appendRow(["Description de l'ouvrage"]);
  sheet.appendRow([payload.workDescription || ""]);
  sheet.appendRow([""]);
  sheet.appendRow(["Construction"]);
  sheet.appendRow([constructionText_(payload)]);
  sheet.appendRow([""]);
  sheet.appendRow(["Note particuliere sur la visite"]);
  sheet.appendRow([payload.visitNote || ""]);
  sheet.appendRow([""]);

  (payload.levels || []).forEach((level) => {
    sheet.appendRow([`Desordres sur ${level.name || ""}`]);
    sheet.appendRow(["Localisation", "Gravite", "Commentaire", "Photos"]);
    (level.entries || []).forEach((entry) => {
      const links = photoRecords
        .filter((photo) => photo.entryKey === entryKey_(level.name, entry))
        .map((photo) => photo.url)
        .join("\n");
      sheet.appendRow([
        entry.localisation || "",
        entry.gravity || "",
        entry.comment || "",
        links
      ]);
    });
    sheet.appendRow([""]);
  });

  sheet.appendRow(["Conclusion"]);
  sheet.appendRow([payload.conclusion || ""]);
  sheet.appendRow([""]);
  sheet.appendRow(["Preconisation"]);
  sheet.appendRow([payload.recommendation || ""]);
  sheet.autoResizeColumns(1, 4);
}

function savePhotos_(payload, interventionId, photoFolder) {
  const records = [];
  const sitePhoto = payload.sitePhoto;
  if (sitePhoto && sitePhoto.dataUrl) {
    const base64 = String(sitePhoto.dataUrl).split(",").pop();
    const bytes = Utilities.base64Decode(base64);
    const fileName = "photo-du-lieu.jpg";
    const blob = Utilities.newBlob(bytes, "image/jpeg", fileName);
    const file = photoFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    records.push({
      levelName: "Photo du lieu",
      localisation: fullSiteAddress_(payload),
      entryKey: "__SITE_PHOTO__",
      isSitePhoto: true,
      name: sitePhoto.name || fileName,
      fileName,
      url: file.getUrl(),
      fileId: file.getId(),
      blob
    });
  }

  (payload.levels || []).forEach((level) => {
    (level.entries || []).forEach((entry, entryIndex) => {
      (entry.photos || []).forEach((photo, photoIndex) => {
        if (!photo.dataUrl) return;
        const base64 = String(photo.dataUrl).split(",").pop();
        const bytes = Utilities.base64Decode(base64);
        const fileName = `${cleanName_(level.name || "niveau")}-${entryIndex + 1}-${photoIndex + 1}.jpg`;
        const blob = Utilities.newBlob(bytes, "image/jpeg", fileName);
        const file = photoFolder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        records.push({
          levelName: level.name || "",
          localisation: entry.localisation || "",
          entryKey: entryKey_(level.name, entry),
          isSitePhoto: false,
          name: photo.name || fileName,
          fileName,
          url: file.getUrl(),
          fileId: file.getId(),
          blob
        });
      });
    });
  });
  return records;
}

function handlePhotoBatch_(payload) {
  const sessionId = normalizeUploadIdentifier_(payload.photoUploadSessionId, "session de transfert");
  const batchId = normalizeUploadIdentifier_(payload.batchId, "lot de photos");
  const photos = Array.isArray(payload.photos) ? payload.photos : [];
  const folder = getTransferFolder_(sessionId, true);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const manifest = readTransferManifest_(folder, sessionId);
    if (manifest.receivedBatchIds.indexOf(batchId) !== -1) {
      return transferStatusFromManifest_(manifest, true);
    }

    photos.forEach((photo, index) => {
      const uploadKey = normalizeUploadIdentifier_(photo.uploadKey || (batchId + "-" + index), "photo");
      if (manifest.photos.some((item) => item.uploadKey === uploadKey)) return;
      if (!photo.dataUrl) throw new Error("Données absentes pour la photo " + uploadKey + ".");

      const blob = photoBlobFromDataUrl_(photo.dataUrl);
      const fileName = transferPhotoFileName_(photo, uploadKey, index);
      const file = folder.createFile(blob.setName(fileName));
      manifest.photos.push({
        uploadKey,
        batchId,
        levelName: photo.levelName || "",
        localisation: photo.localisation || "",
        entryKey: photo.entryId || photo.entryKey || "",
        isSitePhoto: Boolean(photo.isSitePhoto),
        name: photo.name || fileName,
        fileName,
        fileId: file.getId(),
        url: file.getUrl()
      });
      manifest.updatedAt = new Date().toISOString();
      manifest.error = "";
      writeTransferManifest_(folder, manifest);
    });

    manifest.receivedBatchIds.push(batchId);
    manifest.expectedPhotoCount = Number(payload.expectedPhotoCount || manifest.expectedPhotoCount || 0);
    manifest.updatedAt = new Date().toISOString();
    manifest.error = "";
    writeTransferManifest_(folder, manifest);
    if (payload.isFirstBatch) cleanupOldTransfers_(sessionId);
    return transferStatusFromManifest_(manifest, true);
  } finally {
    lock.releaseLock();
  }
}

function adoptUploadedPhotos_(payload, photoFolder) {
  const sessionId = normalizeUploadIdentifier_(payload.photoUploadSessionId, "session de transfert");
  const folder = getTransferFolder_(sessionId, false);
  if (!folder) throw new Error("Le transfert des photos est introuvable.");
  const manifest = readTransferManifest_(folder, sessionId);
  if (manifest.error) {
    manifest.error = "";
    manifest.updatedAt = new Date().toISOString();
    writeTransferManifest_(folder, manifest);
  }

  return (manifest.photos || []).map((photo) => {
    const file = DriveApp.getFileById(photo.fileId);
    file.moveTo(photoFolder);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileName = photo.fileName || photo.name || "photo.jpg";
    return Object.assign({}, photo, {
      fileName,
      url: file.getUrl(),
      fileId: file.getId(),
      blob: file.getBlob().setName(fileName)
    });
  });
}

function getTransferStatus_(sessionId) {
  try {
    const normalized = normalizeUploadIdentifier_(sessionId, "session de transfert");
    const folder = getTransferFolder_(normalized, false);
    if (!folder) return { ok: true, found: false, photoCount: 0, receivedBatchIds: [], finalized: false, error: "" };
    return transferStatusFromManifest_(readTransferManifest_(folder, normalized), true);
  } catch (error) {
    return { ok: false, found: false, photoCount: 0, receivedBatchIds: [], finalized: false, error: String(error.message || error) };
  }
}

function markTransferFinalized_(sessionId, result) {
  if (!sessionId) return;
  const normalized = normalizeUploadIdentifier_(sessionId, "session de transfert");
  const folder = getTransferFolder_(normalized, false);
  if (!folder) return;
  const manifest = readTransferManifest_(folder, normalized);
  manifest.finalized = true;
  manifest.error = "";
  manifest.interventionId = result.interventionId || "";
  manifest.workflowAction = result.workflowAction || "";
  manifest.folderUrl = result.folderUrl || "";
  manifest.reportUrl = result.reportUrl || "";
  manifest.updatedAt = new Date().toISOString();
  writeTransferManifest_(folder, manifest);
}

function markTransferError_(sessionId, detail) {
  if (!sessionId) return;
  try {
    const normalized = normalizeUploadIdentifier_(sessionId, "session de transfert");
    const folder = getTransferFolder_(normalized, false);
    if (!folder) return;
    const manifest = readTransferManifest_(folder, normalized);
    if (manifest.finalized) return;
    manifest.error = String(detail || "Erreur de transfert").substring(0, 2000);
    manifest.updatedAt = new Date().toISOString();
    writeTransferManifest_(folder, manifest);
  } catch (ignored) {
    console.error(String(ignored));
  }
}

function discardPhotoUpload_(sessionId) {
  const normalized = normalizeUploadIdentifier_(sessionId, "session de transfert");
  const folder = getTransferFolder_(normalized, false);
  if (folder) folder.setTrashed(true);
  return { discarded: Boolean(folder) };
}

function getTransferFolder_(sessionId, createIfMissing) {
  const root = getOrCreateFolder_(CONFIG.driveRootFolderName);
  const uploadRoot = getOrCreateSubFolder_(root, CONFIG.uploadRootFolderName);
  const folders = uploadRoot.getFoldersByName(sessionId);
  if (folders.hasNext()) return folders.next();
  return createIfMissing ? uploadRoot.createFolder(sessionId) : null;
}

function readTransferManifest_(folder, sessionId) {
  const files = folder.getFilesByName(CONFIG.uploadManifestFileName);
  if (files.hasNext()) return JSON.parse(files.next().getBlob().getDataAsString("UTF-8"));
  return {
    sessionId,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    expectedPhotoCount: 0,
    receivedBatchIds: [],
    photos: [],
    finalized: false,
    error: ""
  };
}

function writeTransferManifest_(folder, manifest) {
  const content = JSON.stringify(manifest);
  const files = folder.getFilesByName(CONFIG.uploadManifestFileName);
  if (files.hasNext()) files.next().setContent(content);
  else folder.createFile(CONFIG.uploadManifestFileName, content, MimeType.PLAIN_TEXT);
}

function transferStatusFromManifest_(manifest, found) {
  return {
    ok: !manifest.error,
    found: Boolean(found),
    photoCount: (manifest.photos || []).length,
    expectedPhotoCount: Number(manifest.expectedPhotoCount || 0),
    receivedBatchIds: (manifest.receivedBatchIds || []).slice(),
    finalized: Boolean(manifest.finalized),
    interventionId: manifest.interventionId || "",
    workflowAction: manifest.workflowAction || "",
    folderUrl: manifest.folderUrl || "",
    reportUrl: manifest.reportUrl || "",
    error: manifest.error || ""
  };
}

function photoBlobFromDataUrl_(dataUrl) {
  const text = String(dataUrl || "");
  const base64 = text.split(",").pop();
  if (!base64) throw new Error("Photo vide.");
  return Utilities.newBlob(Utilities.base64Decode(base64), "image/jpeg", "photo.jpg");
}

function transferPhotoFileName_(photo, uploadKey, index) {
  if (photo.isSitePhoto) return "photo-du-lieu.jpg";
  const level = cleanName_(photo.levelName || "niveau") || "niveau";
  const entryNumber = Number(photo.entryIndex || 0) + 1;
  const photoNumber = Number(photo.photoIndex || index || 0) + 1;
  const suffix = cleanName_(uploadKey).slice(-12) || String(index + 1);
  return level + "-" + entryNumber + "-" + photoNumber + "-" + suffix + ".jpg";
}

function normalizeUploadIdentifier_(value, label) {
  const normalized = String(value || "").replace(/[^a-zA-Z0-9_-]/g, "").substring(0, 120);
  if (!normalized) throw new Error("Identifiant de " + label + " absent.");
  return normalized;
}

function cleanupOldTransfers_(currentSessionId) {
  const root = getOrCreateFolder_(CONFIG.driveRootFolderName);
  const uploadRoot = getOrCreateSubFolder_(root, CONFIG.uploadRootFolderName);
  const folders = uploadRoot.getFolders();
  const now = Date.now();
  let inspected = 0;
  while (folders.hasNext() && inspected < 30) {
    inspected += 1;
    const folder = folders.next();
    if (folder.getName() === currentSessionId) continue;
    try {
      const manifest = readTransferManifest_(folder, folder.getName());
      const age = now - new Date(manifest.updatedAt || manifest.createdAt || 0).getTime();
      const limit = manifest.finalized ? 24 * 60 * 60 * 1000 : 7 * 24 * 60 * 60 * 1000;
      if (age > limit) folder.setTrashed(true);
    } catch (ignored) {
      console.error(String(ignored));
    }
  }
}

function createReport_(payload, interventionId, interventionFolder, photoRecords) {
  let docFile;
  if (CONFIG.templateDocId) {
    docFile = DriveApp.getFileById(CONFIG.templateDocId).makeCopy(interventionId, interventionFolder);
  } else {
    const doc = DocumentApp.create(interventionId);
    docFile = DriveApp.getFileById(doc.getId());
    interventionFolder.addFile(docFile);
    DriveApp.getRootFolder().removeFile(docFile);
  }

  const doc = DocumentApp.openById(docFile.getId());
  const body = doc.getBody();
  replacePlaceholders_(body, payload, interventionId);

  if (!CONFIG.templateDocId) {
    buildDefaultReport_(body, payload, interventionId, photoRecords);
  } else {
    body.appendPageBreak();
    appendGeneratedSections_(body, payload, photoRecords);
  }

  doc.saveAndClose();
  return docFile;
}

function replacePlaceholders_(body, payload, interventionId) {
  const values = {
    "{{ID_INTERVENTION}}": interventionId,
    "{{DATE_VISITE}}": payload.visitDate || "",
    "{{HEURE_VISITE}}": payload.visitTime || "",
    "{{INGENIEUR}}": payload.engineer || "",
    "{{DESTINATAIRE}}": clientText_(payload),
    "{{CLIENT}}": clientText_(payload),
    "{{ADRESSE_SITE}}": payload.siteAddress || "",
    "{{CODE_POSTAL}}": payload.postalCode || "",
    "{{VILLE}}": String(payload.city || "").toUpperCase(),
    "{{MISSION}}": payload.mission || "",
    "{{REFERENCE_LISEC}}": payload.lisecReference || "",
    "{{TITRE}}": payload.reportTitle || "",
    "{{DIFFUSION}}": diffusionText_(payload),
    "{{PERSONNES_PRESENTES}}": presentPeopleText_(payload),
    "{{DESCRIPTION_OUVRAGE}}": payload.workDescription || "",
    "{{CONSTRUCTION}}": constructionText_(payload),
    "{{NOTE_VISITE}}": payload.visitNote || "",
    "{{CONCLUSION}}": payload.conclusion || "",
    "{{PRECONISATION}}": payload.recommendation || ""
  };

  Object.keys(values).forEach((key) => {
    body.replaceText(escapeForReplace_(key), values[key]);
  });
}

function buildDefaultReport_(body, payload, interventionId, photoRecords) {
  body.clear();
  body.appendParagraph("NOTE TECHNIQUE NT 1").setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(payload.siteAddress || "").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph([payload.postalCode, payload.city].filter(Boolean).join(" ")).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  appendSitePhoto_(body, photoRecords);
  appendMetaTable_(body, payload, interventionId);
  appendGeneratedSections_(body, payload, photoRecords);
}

function appendGeneratedSections_(body, payload, photoRecords) {
  addSection_(body, "Mission", payload.mission);
  addSection_(body, "Description de l'ouvrage", payload.workDescription);
  addConstructionSection_(body, constructionText_(payload));

  body.appendParagraph("Constatations sur site").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  appendLegendTable_(body);
  addSection_(body, "Note particuliere sur la visite", payload.visitNote);

  (payload.levels || []).forEach((level) => {
    appendLevelSection_(body, level, photoRecords);
  });

  addSection_(body, "Conclusion", payload.conclusion);
  addSection_(body, "Preconisation", payload.recommendation);
  addSection_(body, "Diffusion", diffusionText_(payload));
}

function appendSitePhoto_(body, photoRecords) {
  const sitePhoto = (photoRecords || []).find((photo) => photo.isSitePhoto);
  if (!sitePhoto) return;
  try {
    const image = body.appendImage(sitePhoto.blob);
    const targetWidth = Math.min(460, image.getWidth());
    const ratio = targetWidth / image.getWidth();
    image.setWidth(targetWidth);
    image.setHeight(Math.round(image.getHeight() * ratio));
  } catch (error) {
    body.appendParagraph(sitePhoto.url || "");
  }
}

function addSection_(body, title, text) {
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(text || "");
}

function addConstructionSection_(body, text) {
  body.appendParagraph("Construction").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const lines = String(text || "").split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  if (!lines.length) {
    body.appendParagraph("");
    return;
  }
  lines.forEach((line) => {
    const cleaned = line.replace(/^[-•]\s*/, "");
    body.appendListItem(cleaned).setGlyphType(DocumentApp.GlyphType.BULLET);
  });
}

function appendMetaTable_(body, payload, interventionId) {
  const left = [
    `RDV du : ${payload.visitDate || ""}${payload.visitTime ? " à " + payload.visitTime : ""}`,
    "Présents :",
    presentPeopleText_(payload),
    `Référence LISEC : ${interventionId}`
  ].join("\n");
  const right = [
    "Client :",
    clientText_(payload),
    fullSiteAddress_(payload)
  ].join("\n");

  const table = body.appendTable([[left, right]]);
  table.setBorderWidth(0);
  table.getCell(0, 0).setWidth(240);
  table.getCell(0, 1).setWidth(240);
  styleTableText_(table, 10, false);
}

function appendLegendTable_(body) {
  const table = body.appendTable([
    ["Code couleur", "Risque / Degré d'urgence du traitement", ""],
    ["", "Risque faible", "Long terme (moins de 5 ans)"],
    ["", "Risque moyen", "Moyen terme (moins de 2 ans)"],
    ["", "Risque important", "Court terme (moins de 6 mois)"],
    ["", "Risque critique", "En urgence (sans délai)"]
  ]);

  const colors = ["#4F8A3D", "#D6A000", "#E26B0A", "#B91C1C"];
  table.getRow(0).editAsText().setBold(true);
  table.getCell(0, 0).setWidth(90);
  table.getCell(0, 1).setWidth(190);
  table.getCell(0, 2).setWidth(210);

  for (let i = 1; i < table.getNumRows(); i++) {
    table.getCell(i, 0).setBackgroundColor(colors[i - 1]);
  }

  styleTableText_(table, 9, false);
  body.appendParagraph("");
}

function appendLevelSection_(body, level, photoRecords) {
  body.appendParagraph(`Désordres sur ${level.name || ""}`).setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const entries = level.entries || [];
  if (!entries.length) {
    body.appendParagraph("Aucune localisation renseignée.");
    return;
  }

  const rows = [];
  entries.forEach(() => {
    rows.push([""]);
    rows.push([""]);
    rows.push([""]);
  });
  const table = body.appendTable(rows);
  table.setBorderWidth(0.5);

  entries.forEach((entry, index) => {
    const colorRow = table.getRow(index * 3);
    const textRow = table.getRow(index * 3 + 1);
    const photoRow = table.getRow(index * 3 + 2);
    const colorCell = colorRow.getCell(0);
    const contentCell = textRow.getCell(0);
    const photosCell = photoRow.getCell(0);
    const photos = photoRecords.filter((photo) => photo.entryKey === entryKey_(level.name, entry));
    const gravityColor = gravityColor_(entry.gravity);

    colorCell.setWidth(516);
    colorCell.setBackgroundColor(gravityColor);
    colorCell.setPaddingTop(0);
    colorCell.setPaddingBottom(0);
    colorCell.setPaddingLeft(0);
    colorCell.setPaddingRight(0);
    colorRow.setMinimumHeight(10);
    contentCell.setWidth(516);
    photosCell.setWidth(516);

    clearCell_(contentCell);
    clearCell_(photosCell);

    const title = contentCell.appendParagraph(`Localisation ${index + 1} : ${entry.localisation || ""}`);
    title.editAsText().setBold(true);
    contentCell.appendParagraph(entry.comment || "");

    appendPhotosToCell_(photosCell, photos);
  });

  styleTableText_(table, 10, false);
  body.appendParagraph("");
}

function appendPhotosToCell_(cell, photos) {
  if (!photos.length) {
    cell.appendParagraph("");
    return;
  }

  const paragraph = cell.appendParagraph("");
  photos.forEach((photo, index) => {
    try {
      const image = paragraph.appendInlineImage(photo.blob);
      const targetWidth = Math.max(60, Math.round(image.getWidth() * 0.15));
      const ratio = targetWidth / image.getWidth();
      image.setWidth(targetWidth);
      image.setHeight(Math.round(image.getHeight() * ratio));
      if (index < photos.length - 1) paragraph.appendText("     ");
    } catch (error) {
      paragraph.appendText(photo.url);
      if (index < photos.length - 1) paragraph.appendText("     ");
    }
  });
}

function clearCell_(cell) {
  while (cell.getNumChildren() > 0) {
    cell.removeChild(cell.getChild(0));
  }
}

function styleTableText_(table, size, boldHeader) {
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const text = row.getCell(c).editAsText();
      text.setFontFamily("Arial");
      text.setFontSize(size);
      if (boldHeader && r === 0) text.setBold(true);
    }
  }
}

function gravityColor_(gravity) {
  const value = String(gravity || "").toLowerCase();
  if (value.indexOf("faible") !== -1) return "#4F8A3D";
  if (value.indexOf("moyenne") !== -1 || value.indexOf("moyen") !== -1) return "#D6A000";
  if (value.indexOf("forte") !== -1 || value.indexOf("important") !== -1) return "#E26B0A";
  if (value.indexOf("critique") !== -1 || value.indexOf("imminent") !== -1) return "#B91C1C";
  return "#D9EAD3";
}

function exportGoogleDocAsDocx_(docId, fileName) {
  const url = `https://docs.google.com/document/d/${docId}/export?format=docx`;
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` }
  });
  return response.getBlob().setName(fileName);
}

function sendSummaryEmail_(payload, interventionId, reportFile, docxBlob, photoRecords) {
  const subject = `[RAPPORT WORD] LISEC - ${interventionId}`;
  const body = [
    "Bonjour,",
    "",
    "Le compte rendu d'intervention LISEC a ete genere au format Word.",
    `Dossier : ${interventionId}`,
    "",
    `Date de visite : ${payload.visitDate || ""}`,
    `Heure de visite : ${payload.visitTime || ""}`,
    `Ingenieur : ${payload.engineer || ""}`,
    `Client : ${clientText_(payload)}`,
    `Reference LISEC : ${payload.lisecReference || ""}`,
    `Titre : ${payload.reportTitle || ""}`,
    `Lieu de la visite : ${fullSiteAddress_(payload)}`,
    `Personnes presentes : ${presentPeopleText_(payload)}`,
    `Mission : ${payload.mission || ""}`,
    `Diffusion : ${diffusionText_(payload)}`,
    `Nombre de photos enregistrées dans Drive et intégrées au Word : ${photoRecords.length}`,
    "",
    `Le fichier Word ${interventionId}.docx est joint a ce mail.`
  ].join("\n");

  const attachments = [docxBlob];

  MailApp.sendEmail({
    to: CONFIG.emailTo,
    subject,
    body,
    attachments
  });
}

function sendCompletionEmail_(payload, interventionId, interventionFolder, spreadsheet, photoRecords) {
  const subject = `[A COMPLETER] LISEC - ${interventionId}`;
  const body = [
    "Bonjour,",
    "",
    "Un compte rendu d'intervention LISEC a ete envoye pour complement.",
    "Aucun rapport Word n'a encore ete genere.",
    `Dossier : ${interventionId}`,
    "",
    `Date de visite : ${payload.visitDate || ""}`,
    `Heure de visite : ${payload.visitTime || ""}`,
    `Ingenieur : ${payload.engineer || ""}`,
    `Client : ${clientText_(payload)}`,
    `Reference LISEC : ${payload.lisecReference || ""}`,
    `Titre : ${payload.reportTitle || ""}`,
    `Lieu de la visite : ${fullSiteAddress_(payload)}`,
    `Mission : ${payload.mission || ""}`,
    "",
    `Dossier Drive : ${interventionFolder.getUrl()}`,
    `Tableau de suivi : ${spreadsheet.getUrl()}`,
    `Nombre de photos enregistrées dans Drive : ${photoRecords.length}`
  ].join("\n");

  MailApp.sendEmail({
    to: CONFIG.emailTo,
    subject,
    body
  });
}

function getOrCreateFolder_(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function getOrCreateSubFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function clientText_(payload) {
  return payload.clientBlock || [payload.recipientCivility || "", payload.recipientName || ""].filter(Boolean).join(" ");
}

function fullSiteAddress_(payload) {
  return [
    payload.siteAddress || "",
    [payload.postalCode || "", payload.city || ""].filter(Boolean).join(" ")
  ].filter(Boolean).join("\n");
}

function presentPeopleText_(payload) {
  if (Array.isArray(payload.presentPeopleEntries)) {
    return payload.presentPeopleEntries.filter(Boolean).join("\n");
  }
  return payload.presentPeople || "";
}

function constructionText_(payload) {
  if (Array.isArray(payload.constructionItems)) {
    return payload.constructionItems.filter(Boolean).join("\n");
  }
  return payload.construction || "";
}

function diffusionText_(payload) {
  if (Array.isArray(payload.diffusionEntries)) {
    return payload.diffusionEntries.filter(Boolean).join("\n");
  }
  return payload.diffusion || "";
}

function entryKey_(levelName, entry) {
  if (entry && entry.id) return String(entry.id);
  return [levelName || "", entry.localisation || "", entry.comment || "", entry.gravity || ""].join("||");
}

function cleanName_(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9-]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .substring(0, 60) || "intervention";
}

function uniqueSheetName_(spreadsheet, baseName) {
  const safeBase = baseName.replace(/[\\/?*[\]:]/g, "-").substring(0, 95) || "Intervention";
  let name = safeBase;
  let index = 2;
  while (spreadsheet.getSheetByName(name)) {
    name = `${safeBase.substring(0, 90)}-${index}`;
    index++;
  }
  return name;
}

function escapeForReplace_(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function javascriptResponse_(callback, data) {
  const name = String(callback || "");
  if (!/^[a-zA-Z_$][a-zA-Z0-9_$]*$/.test(name)) {
    return jsonResponse(data);
  }
  const serialized = JSON.stringify(data).replace(/</g, "\\u003c");
  return ContentService
    .createTextOutput(name + "(" + serialized + ");")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
