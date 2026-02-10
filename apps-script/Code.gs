/** Toolbox Talk – Google Apps Script (Web App API + PDF + Mail)
 *
 * Routen:
 *  GET  ?route=participants
 *  POST { route: "session.create", session: {...} }
 *  POST { route: "pdf.generateAndSend", sessionId: "..." }
 */

// === HIER ANPASSEN ===
const SPREADSHEET_ID = "PASTE_SPREADSHEET_ID_HERE"; // Google Sheet ID
const SHEET_PARTICIPANTS = "participants_master";
const SHEET_SESSIONS = "toolbox_talk_sessions";

const TO_EMAIL = "hse@firma.de";     // feste Empfängeradresse
const CC_EMAIL = "leitung@firma.de"; // CC (kann auch leer sein: "")
const REPORT_FOLDER_NAME = "ToolboxTalkReports";
// ======================

function doGet(e) {
  const route = (e && e.parameter && e.parameter.route) ? e.parameter.route : "";
  try {
    if (route === "participants") return json_(getParticipants_());
    return json_({ ok: false, error: "Unknown route" });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) ? e.postData.contents : "{}");
  } catch (err) {
    return json_({ ok: false, error: "Invalid JSON" });
  }

  try {
    if (body.route === "session.create") return json_(createSession_(body.session));
    if (body.route === "pdf.generateAndSend") return json_(generateAndSendPdf_(body.sessionId));
    return json_({ ok: false, error: "Unknown route" });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

// -------------------- Data --------------------

function getParticipants_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_PARTICIPANTS);
  if (!sh) throw new Error(`Sheet not found: ${SHEET_PARTICIPANTS}`);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { ok: true, participants: [] };

  const header = values[0].map(String);
  const idx = idx_(header);

  const participants = values.slice(1)
    .map(r => ({
      participantId: String(r[idx.participantId] || "").trim(),
      name: String(r[idx.name] || "").trim(),
      company: String(r[idx.company] || "").trim(),
      active: String(r[idx.active] ?? "").toUpperCase() !== "FALSE"
    }))
    .filter(p => p.active && p.participantId && p.name);

  return { ok: true, participants };
}

function createSession_(session) {
  validateSession_(session);

  const sessionId = makeSessionId_(session);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_SESSIONS);
  if (!sh) throw new Error(`Sheet not found: ${SHEET_SESSIONS}`);

  const createdAt = new Date().toISOString();

  // Wir speichern Teilnehmer inkl. Unterschrift (DataURL) als JSON-String
  const participantsJson = JSON.stringify(session.participants || []);

  sh.appendRow([
    sessionId,
    session.dateTime,
    session.site || "",
    session.project || "",
    session.topic || "",
    session.trainer || "",
    session.summary || "",
    participantsJson,
    "", // pdfFileId
    "", // pdfUrl
    createdAt
  ]);

  return { ok: true, sessionId };
}

function generateAndSendPdf_(sessionId) {
  if (!sessionId) throw new Error("Missing sessionId");

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_SESSIONS);
  if (!sh) throw new Error(`Sheet not found: ${SHEET_SESSIONS}`);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error("No sessions in sheet");

  const header = values[0].map(String);
  const idx = idx_(header);

  // Session finden
  let rowNumber = -1;
  let row = null;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx.sessionId]) === String(sessionId)) {
      rowNumber = i + 1; // 1-basiert in Sheets
      row = values[i];
      break;
    }
  }
  if (!row) throw new Error("Session not found");

  const session = {
    sessionId: String(row[idx.sessionId]),
    dateTime: String(row[idx.dateTime]),
    site: String(row[idx.site] || ""),
    project: String(row[idx.project] || ""),
    topic: String(row[idx.topic] || ""),
    trainer: String(row[idx.trainer] || ""),
    summary: String(row[idx.summary] || ""),
    participants: JSON.parse(row[idx.participantsJson] || "[]")
  };

  // Sicherheit: alle müssen unterschrieben haben (wie von dir gefordert)
  const missing = (session.participants || []).filter(p => !p.signatureDataUrl);
  if (missing.length) {
    throw new Error("Missing signatures: " + missing.map(m => m.name).join(", "));
  }

  // HTML Template aus Datei report.html
  const html = renderReportHtml_(session);
  const pdfBlob = HtmlService.createHtmlOutput(html)
    .getBlob()
    .setName(`${session.sessionId}_ToolboxTalk.pdf`)
    .setContentType(MimeType.PDF);

  // in Drive speichern
  const folder = getOrCreateFolder_(REPORT_FOLDER_NAME);
  const file = folder.createFile(pdfBlob);
  const pdfUrl = file.getUrl();

  // Sheet updaten: pdfFileId + pdfUrl
  sh.getRange(rowNumber, idx.pdfFileId + 1).setValue(file.getId());
  sh.getRange(rowNumber, idx.pdfUrl + 1).setValue(pdfUrl);

  // Mail senden
  MailApp.sendEmail({
    to: TO_EMAIL,
    cc: CC_EMAIL || undefined,
    subject: `Toolbox Talk Dokumentation: ${session.topic} (${session.dateTime})`,
    htmlBody: `Der PDF-Bericht wurde erstellt.<br><br><a href="${pdfUrl}">PDF öffnen</a><br><br>Session-ID: ${session.sessionId}`,
    attachments: [file.getBlob()]
  });

  return { ok: true, pdfUrl, pdfFileId: file.getId() };
}

// -------------------- HTML / PDF --------------------

function renderReportHtml_(session) {
  // Wir verwenden die HTML-Datei apps-script/templates/report.html
  // In Apps Script läuft das als HtmlTemplate-Datei: report.html
  const tpl = HtmlService.createTemplateFromFile("templates/report");
  tpl.session = session;
  return tpl.evaluate().getContent();
}

// -------------------- Helpers --------------------

function validateSession_(s) {
  if (!s) throw new Error("Missing session");
  if (!s.dateTime) throw new Error("Missing dateTime");
  if (!s.topic) throw new Error("Missing topic");
  if (!s.trainer) throw new Error("Missing trainer");
  if (!Array.isArray(s.participants) || s.participants.length === 0) throw new Error("No participants");
  // Unterschriften sammeln wir im Frontend; hier lassen wir es beim Speichern zu,
  // aber beim PDF-Generieren ist es Pflicht.
}

function makeSessionId_(s) {
  const ts = String(s.dateTime).replace(/[-:T]/g, "").slice(0, 12);
  const rnd = Math.random().toString(36).slice(2, 6).toUpperCase();
  return `TT-${ts}-${rnd}`;
}

function getOrCreateFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function idx_(headerRow) {
  const map = {};
  headerRow.forEach((h, i) => map[String(h).trim()] = i);

  // Erwartete Spalten in toolbox_talk_sessions:
  // sessionId, dateTime, site, project, topic, trainer, summary, participantsJson, pdfFileId, pdfUrl, createdAt
  // Erwartete Spalten in participants_master:
  // participantId, name, company, active
  return new Proxy(map, {
    get: (obj, prop) => {
      if (prop in obj) return obj[prop];
      // fallback: manche Sheets speichern Leerzeichen unglücklich
      const key = String(prop);
      return obj[key];
    }
  });
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


