/****************************************************************
 *  SPONSOR REGISTRATION — Google Apps Script backend
 *  ------------------------------------------------------------
 *  Receives submissions from sponsor-form.html, stores:
 *    • a summary row + one row per representative in a Google Sheet
 *    • every uploaded file (ID docs + payment slips) in Drive,
 *      inside a per-submission sub-folder.
 *
 *  ============ ONE-TIME SETUP ============
 *  1. Go to https://script.google.com  →  New project.
 *  2. Paste this whole file into Code.gs (replace the default).
 *  3. Deploy ▸ New deployment ▸ type "Web app".
 *       - Execute as:  Me
 *       - Who has access:  Anyone
 *     Copy the Web app URL and paste it into SCRIPT_URL in sponsor-form.html.
 *
 *  STORAGE: you do NOT need to fill in any IDs. On the first submission
 *  the script auto-creates a Drive folder ("Sponsor Registrations — Uploads")
 *  and a Sheet ("Sponsor Registrations — Data") in your Drive and remembers
 *  them. Run `showStorage()` any time to print their links in the log.
 *  (Optional: paste your own IDs below to use existing folder/sheet.)
 ****************************************************************/

// ====== CONFIG — fill these in (or run setup() to generate) ======
const ROOT_FOLDER_ID = "";   // Drive folder that will hold all uploads
const SPREADSHEET_ID = "";   // Google Sheet that will hold the data
// =================================================================

/* ====== SECRET SPONSOR DATA — lives ONLY here on the server ======
   Each brand's access code + private amounts. The form never contains
   these; it must send the correct code to read them back.
   ⚠ Change the codes to your own private values before sharing.   */
const TICKET_PRICE = 2900;
const EARLY_BIRD_PRICE = 2800;   // ticket price when bought during early bird
const BRAND_SECRETS = {
  "uncle-kim": { name: "Uncle Kim",       code: "UK-YG-2026", sponsorship: 25000, freeTickets: 2 },
  "macks":     { name: "Macks Marketing", code: "MM-YG-2026", sponsorship: 20000, freeTickets: 1 },
  "kpop":      { name: "Kpop Portal SL",  code: "KP-YG-2026", sponsorship: 10000, freeTickets: 1 }
};
// =================================================================

const SUBMISSIONS_SHEET = "Submissions";
const REPS_SHEET        = "Representatives";


/* ---------- main entry point ---------- */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // verify a brand's access code — returns the private amounts only on match
    if (data.action === "verify") {
      return handleVerify(data.brandKey, data.code);
    }

    // otherwise it's a submission
    return handleSubmit(data);

  } catch (err) {
    return json({ status: "error", message: String(err) });
  }
}

/* ---------- verify access code ---------- */
function handleVerify(brandKey, code) {
  const s = BRAND_SECRETS[brandKey];
  if (s && String(code).trim().toUpperCase() === s.code.toUpperCase()) {
    return json({ ok: true, sponsorship: s.sponsorship, freeTickets: s.freeTickets });
  }
  return json({ ok: false });
}

/* ---------- store a submission ---------- */
function handleSubmit(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // serialise writes so rows/IDs don't collide
  try {
    // re-verify the code server-side; never trust client-sent amounts
    const secret = BRAND_SECRETS[data.brandKey];
    if (!secret || String(data.code).trim().toUpperCase() !== secret.code.toUpperCase()) {
      return json({ status: "error", message: "Invalid or missing access code." });
    }

    // recompute money on the server from the trusted secret
    const reps        = (data.representatives || []).length;
    const earlyBird   = data.earlyBird === true;
    const ticketPrice = earlyBird ? EARLY_BIRD_PRICE : TICKET_PRICE;
    const chargeable  = Math.max(0, reps - secret.freeTickets);
    const ticketCost  = chargeable * ticketPrice;
    const totalPayable = secret.sponsorship + ticketCost;

    const ss   = getSpreadsheet_();
    const root = getRootFolder_();

    // unique id for this submission
    const subId = "SUB-" + Utilities.formatDate(new Date(), "GMT", "yyyyMMdd-HHmmss")
                         + "-" + Math.floor(Math.random() * 900 + 100);

    // a dedicated sub-folder for this submission's files
    const folderName = subId + " · " + secret.name;
    const folder = root.createFolder(folderName);

    // --- save representative ID documents + collect rep rows ---
    const repsSheet = getOrCreateSheet(ss, REPS_SHEET, [
      "Submission ID", "Brand", "Rep #", "Full Name",
      "Meal Preference", "Freebies Needed", "Favorite YG Artist",
      "YG Ult Bias", "ID Document"
    ]);

    (data.representatives || []).forEach((rep, i) => {
      let idLink = "";
      if (rep.idDocument && rep.idDocument.data) {
        const f = saveFile(folder, rep.idDocument,
          `Rep${i + 1}_${sanitize(rep.fullName)}_ID`);
        idLink = f.getUrl();
      }
      repsSheet.appendRow([
        subId, secret.name, i + 1, rep.fullName,
        rep.meal, rep.freebies, rep.favoriteArtist || "",
        rep.ultBias || "", idLink
      ]);
    });

    // --- save payment slips ---
    const slipLinks = (data.paymentSlips || []).map((slip, i) => {
      const f = saveFile(folder, slip, `PaymentSlip_${i + 1}`);
      return f.getUrl();
    });

    // --- summary row ---
    const subSheet = getOrCreateSheet(ss, SUBMISSIONS_SHEET, [
      "Submission ID", "Submitted At", "Brand", "Contact Numbers", "Sponsorship (LKR)",
      "Representatives", "Free Tickets", "Chargeable Tickets", "Early Bird",
      "Ticket Price (LKR)", "Ticket Cost (LKR)", "Total Payable (LKR)",
      "Payment Slips", "Files Folder"
    ]);
    subSheet.appendRow([
      subId,
      data.submittedAt || new Date().toISOString(),
      secret.name,
      (data.contactNumbers || []).join(", "),
      secret.sponsorship,
      reps,
      secret.freeTickets,
      chargeable,
      earlyBird ? "Yes" : "No",
      ticketPrice,
      ticketCost,
      totalPayable,
      slipLinks.join("\n"),
      folder.getUrl()
    ]);

    return json({ status: "ok", submissionId: subId, folder: folder.getUrl() });

  } catch (err) {
    return json({ status: "error", message: String(err) });
  } finally {
    lock.releaseLock();
  }
}


/* ---------- helpers ---------- */

// decode a {name, mimeType, data(base64)} object and store it in `folder`
function saveFile(folder, fileObj, fallbackBase) {
  const bytes = Utilities.base64Decode(fileObj.data);
  const ext   = (fileObj.name && fileObj.name.indexOf(".") > -1)
              ? fileObj.name.substring(fileObj.name.lastIndexOf(".")) : "";
  const name  = (fileObj.name && fileObj.name.trim())
              ? fileObj.name : (fallbackBase + ext);
  const blob  = Utilities.newBlob(bytes, fileObj.mimeType || "application/octet-stream", name);
  return folder.createFile(blob);
}

function getOrCreateSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function sanitize(s) {
  return String(s || "").replace(/[^\w\-]+/g, "_").substring(0, 40);
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/* ---------- storage resolvers (auto-create + remember) ----------
   Order of preference:
   1. the ROOT_FOLDER_ID / SPREADSHEET_ID constants, if you set them;
   2. a previously auto-created one stored in Script Properties;
   3. create a new one now and remember it.                          */
function getRootFolder_() {
  if (ROOT_FOLDER_ID) return DriveApp.getFolderById(ROOT_FOLDER_ID);
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty("ROOT_FOLDER_ID");
  if (saved) return DriveApp.getFolderById(saved);
  const folder = DriveApp.createFolder("Sponsor Registrations — Uploads");
  props.setProperty("ROOT_FOLDER_ID", folder.getId());
  return folder;
}

function getSpreadsheet_() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty("SPREADSHEET_ID");
  if (saved) return SpreadsheetApp.openById(saved);
  const ss = SpreadsheetApp.create("Sponsor Registrations — Data");
  props.setProperty("SPREADSHEET_ID", ss.getId());
  return ss;
}

/* ---------- run any time to see where data is being stored ---------- */
function showStorage() {
  const folder = getRootFolder_();
  const ss     = getSpreadsheet_();
  Logger.log("Uploads folder: %s", folder.getUrl());
  Logger.log("Data sheet:     %s", ss.getUrl());
}
