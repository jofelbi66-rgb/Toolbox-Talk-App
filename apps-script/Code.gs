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
  if (!sh) throw new

