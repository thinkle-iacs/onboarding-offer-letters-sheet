const ONBOARD_SHEET = "FY25 Offer List";
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
  values: ["-", "Create Email", "Add to Onboarding"],
  callbacks: {
    "Create Email": ({ values, a1 }) => {
      SpreadsheetApp.getActiveSpreadsheet().toast("Creating email...");
      console.log("Create email called from ", values, a1);
      let [col, row] = getColAndRowFromA1(a1);
      console.log("So we have row", row, "col", col);
      let row1based = row + 1;
      console.log("Using 1-based references to e.g. ", row1based);

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
      let dept = onboardSheet
        .getRange(`${headerMap[DEPT]}${row1based}`)
        .getValue();
      const logCell = onboardSheet.getRange(`${headerMap[LOG]}${row1based}`);
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
        logCell.setValue(`${logCell.getValue()}
Created Email: ${userAsCreated.primaryEmail}
Password: ${user.password}
Recovery email: ${userAsCreated.recoveryEmail}
         `);
      } catch (e) {
        console.log("Error creating user", e);
        if (e.details.code === 409) {
          SpreadsheetApp.getActiveSpreadsheet().toast(
            `${user.primaryEmail} already exists. Change the name in the ${IACS_ACCOUNT} field
          if this account should be given a different address. Otherwise, ignore :)`
          );
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
          logCell.setValue("Unknown Error:\n" + e);
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
      }
    },
    "Add to Onboarding": ({ values, a1, range }) => {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "Copying row into Onboarding sheet"
      );
      let rowNum = range.getRow();
      const target =
        "https://docs.google.com/spreadsheets/d/1VQ9-z3uTLnTIEHtMbyFiEbWL2G-TgVR0RQhXSFzbBGo/edit#gid=1225031524";
      let cols = getOnboardCols();
      let sheet = range.getSheet();
      let namedValues = {};
      console.log(
        "Acting on",
        rowNum,
        values,
        a1,
        range.getA1Notation(),
        sheet.getName()
      );
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
      }
      console.log("Got vals", namedValues);
      var targetSheet;
      try {
        targetSheet =
          SpreadsheetApp.openByUrl(target).getSheetByName("New Hires");
      } catch (err) {
        SpreadsheetApp.alert(
          `Unable to find sheet named ${targetSheet} in onboarding sheet at ${target}.`
        );
        return;
      }
      let headerRow = 2;
      let headerNames = targetSheet
        .getRange(`${headerRow}:${headerRow}`)
        .getValues()[0];
      let newRow = [];
      let missingCols = [];
      for (let name in namedValues) {
        let columnIndex = headerNames.indexOf(name);
        if (columnIndex > -1) {
          newRow[columnIndex] = namedValues[name];
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
      } else {
        SpreadsheetApp.getUi().alert(
          `Unable to copy data. No matching column headers found; we expected them in row ${headerRow} of the sheet named ${sheet} in the onboarding spreadsheet.`
        );
      }
    },
  },
});
