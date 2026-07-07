const CONFIG = {
  spreadsheetName: "LISEC - Comptes rendus interventions",
  driveRootFolderName: "LISEC - Comptes rendus interventions",
  emailTo: "secretariat2.lisec@gmail.com,monasspref@gmail.com",
  templateDocId: "",
  attachPhotosToEmail: true,
  maxEmailAttachmentBytes: 18 * 1024 * 1024
};

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || "{}");
    const result = handleSubmission(payload);
    return jsonResponse({ ok: true, result });
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error && error.message ? error.message : error) });
  }
}

function doGet() {
  return jsonResponse({ ok: true, message: "LISEC Apps Script pret a recevoir les comptes rendus." });
}

function handleSubmission(payload) {
  const interventionId = makeInterventionId(payload);
  const rootFolder = getOrCreateFolder_(CONFIG.driveRootFolderName);
  const interventionFolder = getOrCreateSubFolder_(rootFolder, interventionId);
  const photoFolder = getOrCreateSubFolder_(interventionFolder, "Photos");

  const photoRecords = savePhotos_(payload, interventionId, photoFolder);
  const spreadsheet = getOrCreateSpreadsheet_();

  writeDatabaseSheets_(spreadsheet, payload, interventionId, photoRecords);
  writeInterventionSheet_(spreadsheet, payload, interventionId, photoRecords);

  const reportFile = createReport_(payload, interventionId, interventionFolder, photoRecords);
  const docxBlob = exportGoogleDocAsDocx_(reportFile.getId(), interventionId + ".docx");

  sendSummaryEmail_(payload, interventionId, reportFile, docxBlob, photoRecords);

  return {
    interventionId,
    spreadsheetUrl: spreadsheet.getUrl(),
    reportUrl: reportFile.getUrl(),
    photoCount: photoRecords.length
  };
}

function makeInterventionId(payload) {
  const datePart = payload.visitDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const timePart = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HHmmss");
  const engineer = cleanName_(payload.engineer || "ingenieur");
  return `${datePart}-${timePart}-${engineer}`;
}

function getOrCreateSpreadsheet_() {
  const files = DriveApp.getFilesByName(CONFIG.spreadsheetName);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
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
    "Titre",
    "Date visite",
    "Heure visite",
    "Ingenieur",
    "Destinataire",
    "Adresse / site",
    "Personnes presentes",
    "Rapport",
    "Nombre photos"
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
}

function ensureSheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

function writeDatabaseSheets_(spreadsheet, payload, interventionId, photoRecords) {
  const interventions = ensureSheet_(spreadsheet, "Interventions", [
    "ID intervention", "Date reception", "Titre", "Date visite", "Heure visite", "Ingenieur", "Destinataire", "Adresse / site", "Personnes presentes", "Rapport", "Nombre photos"
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
    payload.reportTitle || "",
    payload.visitDate || "",
    payload.visitTime || "",
    payload.engineer || "",
    recipientLabel_(payload),
    payload.siteAddress || "",
    payload.presentPeople || "",
    "",
    photoRecords.length
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
  sheet.appendRow(["Titre", payload.reportTitle || ""]);
  sheet.appendRow(["Date de visite", payload.visitDate || ""]);
  sheet.appendRow(["Heure de visite", payload.visitTime || ""]);
  sheet.appendRow(["Ingenieur", payload.engineer || ""]);
  sheet.appendRow(["Destinataire", recipientLabel_(payload)]);
  sheet.appendRow(["Adresse / site", payload.siteAddress || ""]);
  sheet.appendRow(["Personnes presentes", payload.presentPeople || ""]);
  sheet.appendRow([]);
  sheet.appendRow(["Description de l'ouvrage"]);
  sheet.appendRow([payload.workDescription || ""]);
  sheet.appendRow([]);
  sheet.appendRow(["Construction"]);
  sheet.appendRow([payload.construction || ""]);
  sheet.appendRow([]);
  sheet.appendRow(["Note particuliere sur la visite"]);
  sheet.appendRow([payload.visitNote || ""]);
  sheet.appendRow([]);

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
    sheet.appendRow([]);
  });

  sheet.appendRow(["Conclusion"]);
  sheet.appendRow([payload.conclusion || ""]);
  sheet.appendRow([]);
  sheet.appendRow(["Preconisation"]);
  sheet.appendRow([payload.recommendation || ""]);
  sheet.autoResizeColumns(1, 4);
}

function savePhotos_(payload, interventionId, photoFolder) {
  const records = [];
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
    "{{TITRE}}": payload.reportTitle || "",
    "{{DATE_VISITE}}": payload.visitDate || "",
    "{{HEURE_VISITE}}": payload.visitTime || "",
    "{{INGENIEUR}}": payload.engineer || "",
    "{{DESTINATAIRE}}": recipientLabel_(payload),
    "{{ADRESSE_SITE}}": payload.siteAddress || "",
    "{{PERSONNES_PRESENTES}}": payload.presentPeople || "",
    "{{DESCRIPTION_OUVRAGE}}": payload.workDescription || "",
    "{{CONSTRUCTION}}": payload.construction || "",
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
  body.appendParagraph("NOTE TECHNIQUE").setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(payload.reportTitle || "Compte rendu d'intervention").setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
  body.appendParagraph(`Reference : ${interventionId}`);
  body.appendParagraph(`Date de visite : ${payload.visitDate || ""}`);
  body.appendParagraph(`Heure de visite : ${payload.visitTime || ""}`);
  body.appendParagraph(`Ingenieur : ${payload.engineer || ""}`);
  body.appendParagraph(`Destinataire : ${recipientLabel_(payload)}`);
  body.appendParagraph(`Adresse / site : ${payload.siteAddress || ""}`);
  body.appendParagraph(`Personnes presentes : ${payload.presentPeople || ""}`);
  appendGeneratedSections_(body, payload, photoRecords);
}

function appendGeneratedSections_(body, payload, photoRecords) {
  addSection_(body, "Description de l'ouvrage", payload.workDescription);
  addSection_(body, "Construction", payload.construction);
  addSection_(body, "Note particuliere sur la visite", payload.visitNote);

  (payload.levels || []).forEach((level) => {
    body.appendParagraph(`Desordres sur ${level.name || ""}`).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    if (!(level.entries || []).length) {
      body.appendParagraph("Aucune localisation renseignee.");
    }
    (level.entries || []).forEach((entry, index) => {
      body.appendParagraph(`Localisation ${index + 1} : ${entry.localisation || ""}`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
      body.appendParagraph(`Gravite : ${entry.gravity || ""}`);
      body.appendParagraph(entry.comment || "");
      const photos = photoRecords.filter((photo) => photo.entryKey === entryKey_(level.name, entry));
      photos.forEach((photo) => {
        try {
          const image = body.appendImage(photo.blob);
          const maxWidth = 420;
          if (image.getWidth() > maxWidth) {
            const ratio = maxWidth / image.getWidth();
            image.setWidth(maxWidth);
            image.setHeight(Math.round(image.getHeight() * ratio));
          }
          body.appendParagraph(photo.url);
        } catch (error) {
          body.appendParagraph(photo.url);
        }
      });
    });
  });

  addSection_(body, "Conclusion", payload.conclusion);
  addSection_(body, "Preconisation", payload.recommendation);
}

function addSection_(body, title, text) {
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(text || "");
}

function exportGoogleDocAsDocx_(docId, fileName) {
  const url = `https://docs.google.com/document/d/${docId}/export?format=docx`;
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` }
  });
  return response.getBlob().setName(fileName);
}

function sendSummaryEmail_(payload, interventionId, reportFile, docxBlob, photoRecords) {
  const subject = `LISEC - Compte rendu intervention - ${payload.visitDate || interventionId}`;
  const photoLinks = photoRecords.map((photo) => `- ${photo.levelName} / ${photo.localisation} : ${photo.url}`).join("\n");
  const body = [
    "Bonjour,",
    "",
    "Un nouveau compte rendu d'intervention LISEC a ete envoye.",
    "",
    `Titre : ${payload.reportTitle || ""}`,
    `Date de visite : ${payload.visitDate || ""}`,
    `Heure de visite : ${payload.visitTime || ""}`,
    `Ingenieur : ${payload.engineer || ""}`,
    `Destinataire : ${recipientLabel_(payload)}`,
    `Adresse / site : ${payload.siteAddress || ""}`,
    `Personnes presentes : ${payload.presentPeople || ""}`,
    "",
    `Rapport Google Docs : ${reportFile.getUrl()}`,
    "",
    "Liens Drive des photos :",
    photoLinks || "Aucune photo.",
    "",
    "Le rapport Word est joint a ce mail."
  ].join("\n");

  const attachments = [docxBlob];
  if (CONFIG.attachPhotosToEmail) {
    let currentSize = docxBlob.getBytes().length;
    photoRecords.forEach((photo) => {
      const size = photo.blob.getBytes().length;
      if (currentSize + size <= CONFIG.maxEmailAttachmentBytes) {
        attachments.push(photo.blob.setName(photo.fileName));
        currentSize += size;
      }
    });
  }

  MailApp.sendEmail({
    to: CONFIG.emailTo,
    subject,
    body,
    attachments
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

function recipientLabel_(payload) {
  return [payload.recipientCivility || "", payload.recipientName || ""].filter(Boolean).join(" ");
}

function entryKey_(levelName, entry) {
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
