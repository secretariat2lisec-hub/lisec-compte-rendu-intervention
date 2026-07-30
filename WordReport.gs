const WORD_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

function createWordReport_(payload, interventionId, interventionFolder, photoRecords) {
  const response = UrlFetchApp.fetch(CONFIG.wordTemplateUrl, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    throw new Error("Le masque Word LISEC est inaccessible (HTTP " + response.getResponseCode() + ").");
  }

  let parts = Utilities.unzip(response.getBlob().setName("masque-lisec.docx"));
  const documentPart = wordPart_(parts, "word/document.xml");
  const relationshipsPart = wordPart_(parts, "word/_rels/document.xml.rels");
  const stylesPart = wordPart_(parts, "word/styles.xml");
  let documentXml = documentPart.getDataAsString("UTF-8");
  let stylesXml = stylesPart.getDataAsString("UTF-8");
  const state = {
    parts,
    relationshipsXml: relationshipsPart.getDataAsString("UTF-8"),
    nextRelationshipNumber: nextWordRelationshipNumber_(relationshipsPart.getDataAsString("UTF-8")),
    nextImageNumber: 1,
    nextDrawingId: nextWordDrawingId_(documentXml)
  };

  const sitePhoto = (photoRecords || []).find((photo) => photo.isSitePhoto && photo.blob);
  const sitePhotoXml = sitePhoto
    ? wordSitePhotoXml_(state, sitePhoto)
    : "";
  documentXml = replaceWordParagraphToken_(documentXml, "{{PHOTO_PRESENTATION}}", sitePhotoXml);

  const constatsXml = wordConstatsXml_(state, payload, photoRecords || []);
  documentXml = replaceWordParagraphToken_(documentXml, "{{BLOC_CONSTATS}}", constatsXml);
  documentXml = replaceWordParagraphToken_(documentXml, "{{BLOC_VERIFICATIONS_STRUCTURELLES}}", "");
  documentXml = replaceWordParagraphToken_(documentXml, "{{CONSTRUCTION}}", wordConstructionXml_(payload));

  documentXml = removeWordParagraphContaining_(documentXml, "{{CARTE_LOCALISATION}}");
  documentXml = removeWordParagraphContaining_(documentXml, "Ci-dessous une vue satellite");
  documentXml = removeWordParagraphContaining_(documentXml, "Vue satellite");
  documentXml = replaceWordToken_(documentXml, "NOTE TECHNIQUE NT 1", payload.reportTitle || "");
  documentXml = appendWordParagraphValue_(documentXml, "Référence LISEC", payload.lisecReference || "");
  documentXml = insertBeforeWordParagraph_(documentXml, "Conclusion", wordPageBreakXml_());

  if (payload.mission) {
    documentXml = insertAfterWordParagraph_(documentXml, "Objet", wordBodyParagraphXml_(payload.mission));
  }
  if (diffusionForWord_(payload)) {
    documentXml = insertAfterWordParagraph_(documentXml, "Diffusion", wordBodyParagraphXml_(diffusionForWord_(payload)));
  }

  const values = {
    "{{ADRESSE_SITE}}": payload.siteAddress || "",
    "{{CODE_POSTAL}}": payload.postalCode || "",
    "{{VILLE}}": String(payload.city || "").toUpperCase(),
    "{{DATE_VISITE}}": payload.visitDate || "",
    "{{HEURE_VISITE}}": payload.visitTime || "",
    "{{PERSONNES_PRESENTES}}": presentPeopleForWord_(payload),
    "{{INGENIEUR}}": payload.engineer || "",
    "{{DESCRIPTION_OUVRAGE}}": payload.workDescription || "",
    "{{NOTE_VISITE}}": payload.visitNote || "",
    "{{CONCLUSION}}": payload.conclusion || "",
    "{{PRECONISATION}}": payload.recommendation || ""
  };
  Object.keys(values).forEach((token) => {
    documentXml = replaceWordToken_(documentXml, token, values[token]);
  });

  const remainingToken = documentXml.match(/\{\{[^{}]+\}\}/);
  if (remainingToken) {
    throw new Error("Champ Word non traite : " + remainingToken[0]);
  }

  stylesXml = alignWordBodyStylesLeft_(stylesXml);

  state.parts = replaceWordPart_(state.parts, "word/document.xml", documentXml, "application/xml");
  state.parts = replaceWordPart_(
    state.parts,
    "word/_rels/document.xml.rels",
    state.relationshipsXml,
    "application/xml"
  );
  state.parts = replaceWordPart_(state.parts, "word/styles.xml", stylesXml, "application/xml");

  const fileName = interventionId + ".docx";
  const docxBlob = Utilities.zip(state.parts, fileName)
    .setContentType(WORD_MIME_TYPE)
    .setName(fileName);
  return interventionFolder.createFile(docxBlob);
}

function wordPart_(parts, name) {
  const part = (parts || []).find((blob) => blob.getName() === name);
  if (!part) throw new Error("Partie absente du masque Word : " + name);
  return part;
}

function replaceWordPart_(parts, name, content, contentType) {
  return parts.map((blob) => {
    if (blob.getName() !== name) return blob;
    return Utilities.newBlob(content, contentType || "application/xml", name);
  });
}

function replaceWordToken_(xml, token, value) {
  return xml.replace(new RegExp(flexibleWordTextPattern_(token), "g"), wordInlineText_(value));
}

function replaceWordParagraphToken_(xml, token, replacementXml) {
  const regex = wordParagraphPattern_(token, "g");
  return xml.replace(regex, replacementXml || "");
}

function removeWordParagraphContaining_(xml, text) {
  return xml.replace(wordParagraphPattern_(text, "g"), "");
}

function insertAfterWordParagraph_(xml, text, insertionXml) {
  let inserted = false;
  return xml.replace(wordParagraphPattern_(text, "g"), (paragraphXml) => {
    if (inserted) return paragraphXml;
    inserted = true;
    return paragraphXml + insertionXml;
  });
}

function insertBeforeWordParagraph_(xml, text, insertionXml) {
  let inserted = false;
  return xml.replace(wordParagraphPattern_(text, "g"), (paragraphXml) => {
    if (inserted) return paragraphXml;
    inserted = true;
    return insertionXml + paragraphXml;
  });
}

function appendWordParagraphValue_(xml, label, value) {
  if (!value) return xml;
  let inserted = false;
  return xml.replace(wordParagraphPattern_(label, "g"), (paragraphXml) => {
    if (inserted) return paragraphXml;
    inserted = true;
    const valueRun = '<w:r><w:t xml:space="preserve"> ' + escapeWordXml_(value) + '</w:t></w:r>';
    return paragraphXml.replace(/<\/w:p>$/, valueRun + "</w:p>");
  });
}

function wordPageBreakXml_() {
  return '<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>';
}

function alignWordBodyStylesLeft_(stylesXml) {
  return String(stylesXml || "").replace(
    /<w:jc\b[^>]*w:val="(?:both|distribute|thaiDistribute)"[^>]*\/>/g,
    '<w:jc w:val="left"/>'
  );
}

function wordParagraphPattern_(text, flags) {
  const content = "(?:(?!<w:p\\b)[\\s\\S])*?";
  return new RegExp(
    "<w:p\\b[^>]*>" + content + flexibleWordTextPattern_(text) + content + "<\\/w:p>",
    flags || ""
  );
}

function flexibleWordTextPattern_(text) {
  return String(text).split("").map(escapeWordRegex_).join("(?:<[^>]+>)*");
}

function escapeWordRegex_(character) {
  return character.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function wordInlineText_(value) {
  const lines = String(value || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  return lines.map(escapeWordXml_).join('</w:t><w:br/><w:t xml:space="preserve">');
}

function escapeWordXml_(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function presentPeopleForWord_(payload) {
  const people = Array.isArray(payload.presentPeopleEntries)
    ? payload.presentPeopleEntries
    : String(payload.presentPeople || "").split(/\r?\n/);
  return people.map((item) => String(item || "").trim()).filter(Boolean).join(", ");
}

function diffusionForWord_(payload) {
  const recipients = Array.isArray(payload.diffusionEntries)
    ? payload.diffusionEntries
    : String(payload.diffusion || "").split(/\r?\n/);
  return recipients.map((item) => String(item || "").trim()).filter(Boolean).join(", ");
}

function wordConstructionXml_(payload) {
  const items = Array.isArray(payload.constructionItems)
    ? payload.constructionItems
    : String(payload.construction || "").split(/\r?\n/);
  return items.map((item) => String(item || "").trim()).filter(Boolean).map((item) =>
    '<w:p><w:pPr><w:spacing w:before="0" w:after="60"/><w:ind w:left="1701"/><w:jc w:val="left"/></w:pPr>' +
    '<w:r><w:t xml:space="preserve">- ' + escapeWordXml_(item) + '</w:t></w:r></w:p>'
  ).join("");
}

function wordBodyParagraphXml_(text) {
  return '<w:p><w:pPr><w:spacing w:before="0" w:after="120"/><w:jc w:val="left"/></w:pPr><w:r><w:t xml:space="preserve">' +
    wordInlineText_(text) +
    "</w:t></w:r></w:p>";
}

function wordSitePhotoXml_(state, photo) {
  const image = registerWordImage_(state, photo.blob, 2736000, 2052000);
  const alt = "Vue generale de l'ouvrage";
  return wordImageParagraphXml_(image, alt) +
    '<w:p><w:pPr><w:spacing w:before="40" w:after="120"/><w:jc w:val="center"/></w:pPr>' +
    '<w:r><w:rPr><w:i/><w:sz w:val="18"/></w:rPr><w:t>Vue générale de l&apos;ouvrage</w:t></w:r></w:p>';
}

function wordConstatsXml_(state, payload, photoRecords) {
  let xml = "";
  let photoNumber = 1;
  (payload.levels || []).forEach((level) => {
    const entries = (level.entries || []).filter((entry) => {
      const photos = photoRecords.filter((photo) => photo.entryKey === entryKey_(level.name, entry));
      return Boolean(entry.localisation || entry.comment || entry.gravity || photos.length);
    });
    if (!entries.length) return;

    const levelName = String(level.name || "").trim();
    const heading = levelName ? "Désordres sur " + levelName.toLowerCase() : "Désordres constatés";
    xml += '<w:p><w:pPr><w:pStyle w:val="11"/><w:keepNext/><w:spacing w:before="240" w:after="120"/></w:pPr>' +
      '<w:r><w:t xml:space="preserve">' + escapeWordXml_(heading) + "</w:t></w:r></w:p>";

    entries.forEach((entry, index) => {
      const location = String(entry.localisation || "").trim();
      const title = "Localisation " + (index + 1) + (location ? " : " + location : "");
      xml += '<w:p><w:pPr><w:keepNext/><w:spacing w:before="160" w:after="40"/></w:pPr>' +
        '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">' + escapeWordXml_(title) + "</w:t></w:r></w:p>";
      xml += wordGravityLineXml_(entry.gravity);
      xml += wordBodyParagraphXml_(entry.comment || "");

      const photos = photoRecords.filter((photo) => photo.entryKey === entryKey_(level.name, entry) && photo.blob);
      const gallery = wordPhotoTableXml_(state, photos, location, photoNumber);
      xml += gallery.xml;
      photoNumber = gallery.nextPhotoNumber;
    });
  });
  return xml;
}

function wordGravityLineXml_(gravity) {
  const color = wordGravityColor_(gravity);
  return '<w:p><w:pPr><w:keepNext/><w:pBdr><w:bottom w:val="single" w:sz="28" w:space="2" w:color="' +
    color +
    '"/></w:pBdr><w:spacing w:before="0" w:after="120"/><w:ind w:left="0" w:right="0"/></w:pPr></w:p>';
}

function wordGravityColor_(gravity) {
  const value = String(gravity || "").toLowerCase();
  if (value.indexOf("faible") !== -1) return "4F8A3D";
  if (value.indexOf("moyenne") !== -1 || value.indexOf("moyen") !== -1) return "D6A000";
  if (value.indexOf("forte") !== -1 || value.indexOf("important") !== -1) return "E26B0A";
  if (value.indexOf("critique") !== -1 || value.indexOf("imminent") !== -1) return "B91C1C";
  return "BFBFBF";
}

function wordPhotoTableXml_(state, photos, location, firstPhotoNumber) {
  if (!photos.length) return { xml: "", nextPhotoNumber: firstPhotoNumber };
  let xml = '<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="8880"/><w:tblLayout w:type="fixed"/>' +
    '<w:tblInd w:w="120" w:type="dxa"/><w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/>' +
    '<w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/>' +
    '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4440"/><w:gridCol w:w="4440"/></w:tblGrid>';
  let photoNumber = firstPhotoNumber;
  for (let index = 0; index < photos.length; index += 2) {
    xml += '<w:tr><w:trPr><w:cantSplit/></w:trPr>';
    for (let column = 0; column < 2; column++) {
      const photo = photos[index + column];
      if (!photo) {
        xml += wordEmptyPhotoCellXml_();
        continue;
      }
      const caption = "Photo " + String(photoNumber).padStart(2, "0") + (location ? " - " + location : "");
      const image = registerWordImage_(state, photo.blob, 2592000, 2592000);
      xml += wordPhotoCellXml_(image, caption);
      photoNumber += 1;
    }
    xml += "</w:tr>";
  }
  xml += "</w:tbl><w:p><w:pPr><w:spacing w:before=\"0\" w:after=\"60\"/></w:pPr></w:p>";
  return { xml, nextPhotoNumber: photoNumber };
}

function wordPhotoCellXml_(image, caption) {
  return '<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="4440"/><w:vAlign w:val="top"/>' +
    '<w:tcMar><w:top w:w="60" w:type="dxa"/><w:left w:w="90" w:type="dxa"/>' +
    '<w:bottom w:w="80" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tcMar></w:tcPr>' +
    wordImageParagraphXml_(image, caption) +
    '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:jc w:val="center"/></w:pPr>' +
    '<w:r><w:rPr><w:i/><w:color w:val="595959"/><w:sz w:val="16"/></w:rPr>' +
    '<w:t xml:space="preserve">' + escapeWordXml_(caption) + "</w:t></w:r></w:p></w:tc>";
}

function wordEmptyPhotoCellXml_() {
  return '<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="4440"/><w:vAlign w:val="top"/>' +
    '<w:tcMar><w:top w:w="60" w:type="dxa"/><w:left w:w="90" w:type="dxa"/>' +
    '<w:bottom w:w="80" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tcMar></w:tcPr><w:p/></w:tc>';
}

function wordImageParagraphXml_(image, altText) {
  const alt = escapeWordXml_(altText || "Photo");
  return '<w:p><w:pPr><w:spacing w:before="0" w:after="40"/><w:jc w:val="center"/></w:pPr><w:r><w:drawing>' +
    '<wp:inline xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<wp:extent cx="' + image.cx + '" cy="' + image.cy + '"/>' +
    '<wp:docPr id="' + image.drawingId + '" name="Picture ' + image.drawingId + '" descr="' + alt + '" title="' + alt + '"/>' +
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="' + escapeWordXml_(image.fileName) + '"/><pic:cNvPicPr/></pic:nvPicPr>' +
    '<pic:blipFill><a:blip r:embed="' + image.relationshipId + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + image.cx + '" cy="' + image.cy + '"/></a:xfrm>' +
    '<a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline>' +
    "</w:drawing></w:r></w:p>";
}

function registerWordImage_(state, sourceBlob, maxCx, maxCy) {
  let blob = sourceBlob.copyBlob();
  let contentType = String(blob.getContentType() || "").toLowerCase();
  let extension = contentType.indexOf("png") !== -1 ? "png" : "jpeg";
  if (extension === "jpeg" && contentType.indexOf("jpeg") === -1 && contentType.indexOf("jpg") === -1) {
    blob = blob.getAs("image/jpeg");
    contentType = "image/jpeg";
  }

  const fileName = "lisec-report-" + state.nextImageNumber + "." + extension;
  const packageName = "word/media/" + fileName;
  const relationshipId = "rId" + state.nextRelationshipNumber;
  state.nextImageNumber += 1;
  state.nextRelationshipNumber += 1;

  state.parts.push(blob.setName(packageName));
  const relationship = '<Relationship Id="' + relationshipId + '" ' +
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
    'Target="media/' + fileName + '"/>';
  state.relationshipsXml = state.relationshipsXml.replace("</Relationships>", relationship + "</Relationships>");

  const sourceSize = wordImagePixelSize_(blob);
  const fittedSize = fitWordImage_(sourceSize.width, sourceSize.height, maxCx, maxCy);
  const drawingId = state.nextDrawingId;
  state.nextDrawingId += 1;
  return { relationshipId, drawingId, fileName, cx: fittedSize.cx, cy: fittedSize.cy };
}

function nextWordRelationshipNumber_(relationshipsXml) {
  const matches = Array.from(String(relationshipsXml || "").matchAll(/Id="rId(\d+)"/g));
  return matches.reduce((maximum, match) => Math.max(maximum, Number(match[1]) || 0), 0) + 1;
}

function nextWordDrawingId_(documentXml) {
  const matches = Array.from(String(documentXml || "").matchAll(/<wp:docPr[^>]+id="(\d+)"/g));
  return matches.reduce((maximum, match) => Math.max(maximum, Number(match[1]) || 0), 1000) + 1;
}

function fitWordImage_(width, height, maxCx, maxCy) {
  const safeWidth = Math.max(1, Number(width) || 4);
  const safeHeight = Math.max(1, Number(height) || 3);
  const scale = Math.min(maxCx / safeWidth, maxCy / safeHeight);
  return {
    cx: Math.max(1, Math.round(safeWidth * scale)),
    cy: Math.max(1, Math.round(safeHeight * scale))
  };
}

function wordImagePixelSize_(blob) {
  const bytes = blob.getBytes().map((value) => value < 0 ? value + 256 : value);
  if (bytes.length >= 24 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) {
    return {
      width: wordUnsignedInt32_(bytes, 16),
      height: wordUnsignedInt32_(bytes, 20)
    };
  }
  if (bytes.length >= 4 && bytes[0] === 0xff && bytes[1] === 0xd8) {
    let offset = 2;
    const startOfFrameMarkers = [0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf];
    while (offset + 8 < bytes.length) {
      if (bytes[offset] !== 0xff) {
        offset += 1;
        continue;
      }
      const marker = bytes[offset + 1];
      offset += 2;
      if (marker === 0xd8 || marker === 0xd9) continue;
      if (offset + 1 >= bytes.length) break;
      const length = bytes[offset] * 256 + bytes[offset + 1];
      if (startOfFrameMarkers.indexOf(marker) !== -1 && offset + 6 < bytes.length) {
        return {
          height: bytes[offset + 3] * 256 + bytes[offset + 4],
          width: bytes[offset + 5] * 256 + bytes[offset + 6]
        };
      }
      if (length < 2) break;
      offset += length;
    }
  }
  return { width: 4, height: 3 };
}

function wordUnsignedInt32_(bytes, offset) {
  return (((bytes[offset] * 256 + bytes[offset + 1]) * 256 + bytes[offset + 2]) * 256 + bytes[offset + 3]);
}
