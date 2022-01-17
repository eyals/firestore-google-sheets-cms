/** ================== SERVICE ACCOUNT ================== */

/**
 * Returns the ServiceAccount sheet.
 * If it doesn't exist - creates one.
 */
function serviceAccountSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ServiceAccount");
  if (!sheet) {
    var newSheet = createServiceAccountSheet();
    if (newSheet == null) return null;
    return newSheet;
  } else {
    sheet.hideSheet();
    return sheet;
  }
}

/**
 * Guides the user through a series of questions.
 * The answers are added to a new, hidden sheet, called ServiceAccount,
 * which will be used for Firebase authentication.
 * This event is triggered in these cases:
 * - Trying to sync when there is no Service Account sheet
 * - Clicking the action in the menu (good for replacing the key)
 */
function createServiceAccountSheet() {

  var ui = sheetsUi();

  /** Explains what the process is for */
  var introPrompt = ui.alert('Set up a service account', 'Let\'s set up the connection to your Firebase project.\nNote: This will store your credentials in a hidden tab, so don\'t share this spreadsheet with people you don\'t trust.\n\nClick OK to continue, or Cancel to do it later.', ui.ButtonSet.OK_CANCEL);
  if (introPrompt == ui.Button.CANCEL) return null;

  /** Asks for project ID */
  var projPrompt = ui.prompt('What is your Firebase project ID?', '', ui.ButtonSet.OK_CANCEL);
  if (projPrompt.getSelectedButton() == ui.Button.CANCEL) return;
  var projectId = projPrompt.getResponseText().split("\"").join("");
  if (projectId == "") return;

  /** Instructs how to get the key file */
  var getKeyAlert = ui.alert('Get a private key', '- Visit https://console.firebase.google.com/project/' + projectId + '/settings/serviceaccounts\n- Click "Generate new private key"\n- Open the downloaded file in a text editor.\n- You\'ll need to copy two fields from there in the next screens.\n\nClick OK to contine.', ui.ButtonSet.OK_CANCEL);
  if (getKeyAlert == ui.Button.CANCEL) return null;

  /** Asks for client email */
  var emailPrompt = ui.prompt('Your Service Account email', 'Appears as \'client_email\' in your private key.\nThis is NOT your personal email.', ui.ButtonSet.OK_CANCEL);
  if (emailPrompt.getSelectedButton() == ui.Button.CANCEL) return null;
  var email = emailPrompt.getResponseText().split("\"").join("");
  if (email == "") return null;

  /** Asks for secret key */
  var keyPrompt = ui.prompt('Your Service Account key', 'Appears as \'private_key\' in your private key.\nIt\'s a long blob starting with -----BEGIN PRIVATE KEY----- and ending with -----END PRIVATE KEY-----\\n.', ui.ButtonSet.OK_CANCEL);
  if (keyPrompt.getSelectedButton() == ui.Button.CANCEL) return null;
  var key = keyPrompt.getResponseText().split("\"").join("");
  if (key == "") return null;

  /** Creates the sheet, hides it, and sets the values */
  var newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("ServiceAccount");
  newSheet.hideSheet();
  newSheet.getRange("A1:B5").setValues([
    ["projectId", projectId],
    ["email", email],
    ["key", key],
    ["", ""],
    ["WARNING:", "Do not modify this sheet if you're not sure what this is. This sheet is best as hidden.\nUse the Firestore menu to generate a service account if you fail to connect to Firebase."]
  ]);

  toast("Service account set up and ready to use");
  return newSheet;

}

var serviceAccount = () => {
  var saSheet = serviceAccountSheet();
  if (!saSheet) return null;
  return {
    email: saSheet.getRange("ServiceAccount!B2").getValue(),
    projectId: saSheet.getRange("ServiceAccount!B1").getValue(),
    key: saSheet.getRange("ServiceAccount!B3").getValue().replace(/\\n/g, '\n'),
  }
}