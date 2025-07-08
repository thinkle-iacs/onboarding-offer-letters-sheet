const ONBOARD_SPREADSHEET =
  "https://docs.google.com/spreadsheets/d/1VQ9-z3uTLnTIEHtMbyFiEbWL2G-TgVR0RQhXSFzbBGo/edit#gid=1225031524";
const ONBOARD_SHEET = "FY26 Offer List";
const NUDGE_CONFIG_SHEET = "Nudges";
const ONBOARD_HEADER_RANGE = "1:1";
const FIRST = "First Name";
const LAST = "Last Name";
const PERSONAL_EMAIL = "Email";
const IACS_ACCOUNT = "IACS Username";
const ACTION = "Action";
const OU = "OU";
const LOG = "Log";
const POSITION = "Position";
const SCHOOL = "School";
const DEPT = "Department";
const LASTFIRST = "Last, First";
const headers = [
  LASTFIRST,
  FIRST,
  LAST,
  PERSONAL_EMAIL,
  IACS_ACCOUNT,
  ACTION,
  LOG,
  OU,
  POSITION,
  SCHOOL,
  DEPT,
];

function sanitizeName(name) {
  // Convert to lowercase
  name = name.toLowerCase();
  // Replace spaces with nothing
  name = name.replace(/\s+/g, "");
  // Remove invalid characters
  name = name.replace(/[^a-z0-9._%+-]+/g, "");
  // Ensure email doesn't start or end with a dot
  name = name.replace(/[.]/, "");
  return name;
}

function getOnboardCols() {
  let onboardSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONBOARD_SHEET);
  let headerRow = onboardSheet.getRange(ONBOARD_HEADER_RANGE).getValues()[0];
  const headerMap = {};
  for (let cn = 0; cn < headerRow.length; cn++) {
    let header = headerRow[cn];
    headerMap[header] = colOffsetToLetter(cn);
  }
  return headerMap;
}

function testCols() {
  console.log("Got map: ", getOnboardCols());
}

let headerMap = getOnboardCols();
let actionCells = headerMap[ACTION] + "2:" + headerMap[ACTION]; // i.e. D2:D
console.log("actionCells are", actionCells);
addDropdown({
  cell: actionCells,
  sheet: ONBOARD_SHEET,
  values: [
    "-",
    "Create Email",
    "Add to Onboarding",
    "Send Nudges",
    "Completed Task",
  ],
  callbacks: {
    "Create Email": createEmail,
    "Add to Onboarding": addToOnboarding,
    "Send Nudges": nudgeUsers,
  },
});

function markComplete({ a1 }) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(a1);
  range.setValue("Completed");
}
function markUnset({ a1 }) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(a1);
  range.setValue("-");
}

function readNudgeConfig() {
  let nudgeSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NUDGE_CONFIG_SHEET);
  let nudgeData = nudgeSheet.getDataRange().getValues();
  let nudges = [];
  let nudgeHeaders = nudgeData[0];
  for (let i = 1; i < nudgeData.length; i++) {
    let nudge = {};
    for (let j = 0; j < nudgeData[i].length; j++) {
      nudge[nudgeHeaders[j]] = nudgeData[i][j];
    }
    if (nudge.NudgeEmail && nudge.NudgeHTML) {
      nudges.push(nudge);
    } else {
      console.error("Ignoring invalid nudge at row", i + 1, nudge);
    }
  }
  return nudges;
}

function nudgeUsers({ values, a1, range, partial }) {
  let namedValues = getNamedValues({ range, a1 });
  let nudges = readNudgeConfig();
  let nudgesToSend = nudges.filter((nudge) => {
    for (let col in nudge) {
      if (nudge[col] && nudge[col] === namedValues[col]) {
        return true;
      }
    }
  });
  console.log("Nudges to send", nudgesToSend);
  let ONBOARD_LINK = `<a href="${ONBOARD_SPREADSHEET}">Onboarding Sheet</a>`;
  let GOOGLE_LINK = `<a href="${namedValues["Google Onboarding Link"]}">Google Onboarding Tool</a>`;
  SpreadsheetApp.getActiveSpreadsheet().toast("Sending nudges...");
  nudgesToSend.forEach((nudge) => {
    let htmlBody = completeTemplate(nudge.NudgeHTML, {
      ...namedValues,
      ONBOARD_LINK,
      ONBOARD_SPREADSHEET,
      GOOGLE_LINK,
    });
    console.log("Sending nudge: ", nudge);
    GmailApp.sendEmail(
      nudge.NudgeEmail,
      completeTemplate(nudge.NudgeSubject, namedValues),
      "",
      { htmlBody }
    );
  });
  let nudgeMessage =
    "Nudged " + nudgesToSend.map((n) => n.NudgeEmail).join(", ");
  SpreadsheetApp.getActiveSpreadsheet().toast(nudgeMessage);
  addToLog({ a1, message: nudgeMessage });
  if (!partial) markComplete({ a1 });
}

function completeTemplate(template, values) {
  for (let key in values) {
    template = template.replace(new RegExp(`{{${key}}}`, "g"), values[key]);
  }
  return template;
}

function addToLog({ a1, message }) {
  let [col, row] = getColAndRowFromA1(a1);
  let row1based = row + 1;
  let onboardSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONBOARD_SHEET);
  let logCell = onboardSheet.getRange(`${headerMap[LOG]}${row1based}`);
  let logValue = logCell.getValue();
  let shortDate = Utilities.formatDate(
    new Date(),
    "GMT",
    "MM/dd/yyyy HH:mm:ss"
  );
  logCell.setValue(`${logValue}\n${shortDate}: ${message}`);
}

function createEmail({ a1 }) {
  SpreadsheetApp.getActiveSpreadsheet().toast("Creating email...");
  let [col, row] = getColAndRowFromA1(a1);
  let row1based = row + 1;

  let onboardSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONBOARD_SHEET);
  let firstName = onboardSheet
    .getRange(`${headerMap[FIRST]}${row1based}`)
    .getValue();
  let lastName = onboardSheet
    .getRange(`${headerMap[LAST]}${row1based}`)
    .getValue();
  let ou = onboardSheet.getRange(`${headerMap[OU]}${row1based}`).getValue();
  if (!ou) {
    SpreadsheetApp.getUi().alert(
      `No OU found for this user. Please enter an OU in column ${headerMap[OU]} then try again.`
    );
    markUnset({ a1 });
    return;
  }
  let position = onboardSheet
    .getRange(`${headerMap[POSITION]}${row1based}`)
    .getValue();
  let personalEmail = onboardSheet
    .getRange(`${headerMap[PERSONAL_EMAIL]}${row1based}`)
    .getValue();
  let school = onboardSheet
    .getRange(`${headerMap[SCHOOL]}${row1based}`)
    .getValue();
  let dept = onboardSheet.getRange(`${headerMap[DEPT]}${row1based}`).getValue();
  // Generate email...
  let iacsEmailRange = onboardSheet.getRange(
    `${headerMap[IACS_ACCOUNT]}${row1based}`
  );
  let iacsEmail = iacsEmailRange.getValue();
  let generatedEmail = "";
  if (!iacsEmail) {
    generatedEmail =
      sanitizeName(firstName[0] + lastName) + "@innovationcharter.org";
    iacsEmailRange.setValue(generatedEmail);
  }
  // Display a confirmation that we want to create a user...
  const user = {
    primaryEmail: iacsEmail || generatedEmail,
    password: generatePassword(),
    name: {
      givenName: firstName,
      familyName: lastName,
    },
    orgUnitPath: ou,
    recoveryEmail: personalEmail,
    organizations: [
      {
        name: school,
        title: position,
        department: dept,
      },
    ],
  };
  console.log("Creating account", user);

  var userAsCreated;
  try {
    userAsCreated = AdminDirectory.Users.insert(user);
    console.log("Created", userAsCreated);
    addToLog({
      a1,
      message: `Created email: ${userAsCreated.primaryEmail} Password: ${user.password} Recovery email: ${userAsCreated.recoveryEmail}`,
    });
  } catch (e) {
    console.log("Error creating user", e);
    if (e.details.code === 409) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `${user.primaryEmail} already exists. Change the name in the ${IACS_ACCOUNT} field
        if this account should be given a different address. Otherwise, ignore :)`
      );
      addToLog({ a1, message: `User ${iacsEmail} already exists.` });
      markUnset({ a1 });
      return {
        success: false,
        error: "User already exists",
      };
    } else {
      // in all other cases, let's log the error
      if (e.details.code === 400) {
        SpreadsheetApp.getUi().alert(
          `Something is wrong with the user request -- take a look at the "Log" column for details.`
        );
      } else if (e.details.code === 412) {
        SpreadsheetApp.getUi().alert(
          `User limit exceeded. Talk to Tom stat and take a look at the log cell!`
        );
      }
      addToLog({ a1, message: "Unknown Error:\n" + e });
      markUnset({ a1 });
      return {
        success: false,
        error: e,
      };
    }
  }
  if (userAsCreated) {
    let mailSubject = "Account for Innovation Academy";
    let mailBody = `Welcome to IACS. An account has been created for you!
      
Username: ${userAsCreated.primaryEmail} 
Password: ${user.password}

You can log in by going to mail.innovationcharter.org
(or gmail will work if our custom URL is down for some reason).

If you have questions, please reach out to support@innovationcharter.org

- The IACS Tech Team
`;
    GmailApp.sendEmail(personalEmail, mailSubject, mailBody);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Created email and sent credentials to " + personalEmail
    );
    markComplete({ a1 });
    return { success: true, user: userAsCreated };
  }
}

function addToOnboarding({ values, a1, range }) {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Copying row into Onboarding sheet"
  );
  const target = ONBOARD_SPREADSHEET;
  var namedValues = getNamedValues({ range, values, a1 });
  console.log("Got vals", namedValues);
  var targetSheet;
  try {
    targetSheet = SpreadsheetApp.openByUrl(target).getSheetByName("New Hires");
  } catch (err) {
    SpreadsheetApp.alert(
      `Unable to find sheet named ${targetSheet} in onboarding sheet at ${target}.`
    );
    markUnset({ a1 });
    return;
  }
  let headerRow = 2;
  let headerNames = targetSheet
    .getRange(`${headerRow}:${headerRow}`)
    .getValues()[0];
  let newRow = [];
  let missingCols = [];
  let copiedHeaders = [];
  for (let name in namedValues) {
    let columnIndex = headerNames.indexOf(name);
    if (columnIndex > -1) {
      newRow[columnIndex] = namedValues[name];
      copiedHeaders.push(name);
    } else {
      missingCols.push(name);
    }
  }
  console.log("New row is ", newRow);
  if (newRow.length) {
    targetSheet.appendRow(newRow);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Done copying row into Onboarding sheet!"
    );
    addToLog({
      a1,
      message: `Copied rows to Onboarding sheet: ${copiedHeaders.join(", ")}`,
    });
    markComplete({ a1 });
  } else {
    SpreadsheetApp.getUi().alert(
      `Unable to copy data. No matching column headers found; we expected them in row ${headerRow} of the sheet named ${sheet} in the onboarding spreadsheet.`
    );
    markUnset({ a1 });
  }
}

/*
 * Return a map of column names to values for the given row
 */
function getNamedValues({ range, a1 }) {
  let rowNum = range.getRow();
  let cols = getOnboardCols();
  let sheet = range.getSheet();
  let namedValues = {};
  console.log("Acting on", rowNum, a1, range.getA1Notation(), sheet.getName());
  for (let col in cols) {
    let a1 = cols[col] + rowNum;
    namedValues[col] = sheet.getRange(a1).getValue();
  }
  // inserting magic value...
  if (namedValues[PERSONAL_EMAIL].includes("@innovationcharter.org")) {
    namedValues[IACS_ACCOUNT] = namedValues[PERSONAL_EMAIL];
  }
  if (namedValues[IACS_ACCOUNT]) {
    namedValues[
      "Google Onboarding Link"
    ] = `https://tinyurl.com/iacs-onboard?u=${
      namedValues[IACS_ACCOUNT].split("@")[0]
    }`;
  } else {
    namedValues["Google Onboarding Link"] = "https://tinyurl.com/iacs-onboard";
  }
  return namedValues;
}
