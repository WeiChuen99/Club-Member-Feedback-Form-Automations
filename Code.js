var RESPONSES_SHEET = "Form responses 1";
var EMAILS_SHEET = "SpeakerEmails";
var SENT_COLUMN = "Sent Email";
var SENT_VALUE = "yes";

// Adds a "Feedback" menu to the spreadsheet toolbar.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Feedback")
    .addItem("Send tonight's feedback emails", "sendTonightsFeedback")
    .addToUi();
}

// Sends one consolidated email per speaker containing all unsent feedback.
// Safe to re-run: each row is marked "yes" in the Sent Email column once sent.
// Can be run from the Feedback menu or a time-driven trigger.
function sendTonightsFeedback() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var responseSheet = ss.getSheetByName(RESPONSES_SHEET);
  var emailSheet = ss.getSheetByName(EMAILS_SHEET);

  if (!responseSheet || !emailSheet) {
    Logger.log("Error: missing '" + RESPONSES_SHEET + "' or '" + EMAILS_SHEET + "' tab.");
    return;
  }

  var values = responseSheet.getDataRange().getValues();
  if (values.length < 2) return;

  var header = values[0].map(function(h) { return h.toString().trim(); });
  var col = function(name) { return header.indexOf(name); };

  var colTimestamp = col("Timestamp");
  var colSpeaker = col("Speaker Name");
  var colLiked = col("What did you like about the speech?");
  var colImprove = col("What can the speaker do to make the speech even better?");
  var colSent = col(SENT_COLUMN);

  // Create the Sent Email column if it doesn't exist yet
  if (colSent === -1) {
    colSent = header.length;
    responseSheet.getRange(1, colSent + 1).setValue(SENT_COLUMN);
  }

  var emails = buildEmailLookup_(emailSheet);

  // Group unsent feedback rows by speaker (normalized name)
  var bySpeaker = {};
  var rowsBySpeaker = {};
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row[colSent] && row[colSent].toString().trim().toLowerCase() === SENT_VALUE) continue;

    var speaker = row[colSpeaker] ? row[colSpeaker].toString().trim() : "";
    var timestamp = row[colTimestamp];
    if (!speaker || !timestamp) continue; // skip blank rows

    var key = speaker.toLowerCase();
    if (!bySpeaker[key]) {
      bySpeaker[key] = { name: speaker, feedback: [] };
      rowsBySpeaker[key] = [];
    }
    bySpeaker[key].feedback.push({
      liked: row[colLiked] ? row[colLiked].toString().trim() : "",
      improve: row[colImprove] ? row[colImprove].toString().trim() : ""
    });
    rowsBySpeaker[key].push(i + 1); // 1-based sheet row number
  }

  var speakerKeys = Object.keys(bySpeaker);
  if (speakerKeys.length === 0) {
    Logger.log("No unsent feedback found.");
    return;
  }

  speakerKeys.forEach(function(key) {
    var speaker = bySpeaker[key];
    var email = emails[key];

    if (!email) {
      Logger.log("No email found for speaker '" + speaker.name + "' - rows left unsent.");
      return;
    }

    var subject = "Your Toastmasters speech feedback (" +
      speaker.feedback.length + " response" + (speaker.feedback.length > 1 ? "s" : "") + ")";

    var body = "Hi " + speaker.name + ",\n\n" +
      "Here is the feedback you received for your speech:\n";

    speaker.feedback.forEach(function(f, n) {
      body += "\n--------------------------------------------------\n" +
        "Feedback " + (n + 1) + "\n\n";
      if (f.liked) body += "What they liked:\n" + f.liked + "\n\n";
      if (f.improve) body += "Suggestions to make it even better:\n" + f.improve + "\n";
      if (!f.liked && !f.improve) body += "(No written feedback)\n";
    });

    body += "\n--------------------------------------------------\n\nGreat job on your speech!";

    GmailApp.sendEmail(email, subject, body);

    rowsBySpeaker[key].forEach(function(rowNum) {
      responseSheet.getRange(rowNum, colSent + 1).setValue(SENT_VALUE);
    });

    Logger.log("Sent " + speaker.feedback.length + " feedback entr" +
      (speaker.feedback.length > 1 ? "ies" : "y") + " to " + speaker.name + " <" + email + ">");
  });
}

// Builds a normalized (trimmed, lowercase) speaker name -> email map.
function buildEmailLookup_(emailSheet) {
  var data = emailSheet.getDataRange().getValues();
  var lookup = {};
  for (var i = 1; i < data.length; i++) {
    var name = data[i][0] ? data[i][0].toString().trim().toLowerCase() : "";
    var email = data[i][1] ? data[i][1].toString().trim() : "";
    if (name && email) lookup[name] = email;
  }
  return lookup;
}
