/**
 * @param {number} col - A 0-based index representing column location.
 * Return the letter representation of the column.
 */
function colOffsetToLetter(col) {
  let letter = "";
  while (col >= 0) {
    letter = String.fromCharCode((col % 26) + 65) + letter;
    col = Math.floor(col / 26) - 1;
  }
  return letter;
}

// Test the function
function testColOffsetToLetter() {
  const tests = [0, 1, 25, 26, 27, 52, 702, 703];
  tests.forEach((col) => {
    console.log(`${col} => ${colOffsetToLetter(col)}`);
  });
}

/**
 * @param {string} letters - A column representation (series of capital letters from A to Z);
 * Return a 0-based column offset (A=0, B=1, etc) based on a string column description
 */
function letterToColOffset(col) {
  col = col.toUpperCase();
  let offset = 0;
  let multiplier = 1;
  let letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  for (let letterIndex = col.length - 1; letterIndex >= 0; letterIndex--) {
    let theLetter = col[letterIndex];
    let theNumber = letters.indexOf(theLetter);
    if (theNumber === -1) {
      throw new Error("Invalid column string: " + col);
    }
    offset += (theNumber + 1) * multiplier;
    multiplier *= 26;
  }
  return offset - 1;
}

function testLetterToColOffset() {
  for (let col of [
    "A",
    "B",
    "Z",
    "AA",
    "AZ",
    "BA",
    "ZZ",
    "AAA",
    "AAB",
    "ABA",
    "QR",
    "&@",
  ]) {
    try {
      let result = letterToColOffset(col);
      console.log(col, "=>", result);
    } catch (e) {
      console.log(col, "=>ERROR:", e);
    }
  }
}

/**
 * @param {string} a1 - The A1 description of a cell.
 * May include a column and a row or just one or the other.
 * We return [COLUMN,ROW] where each is a numeric 0-based offset *or* null
 * where NULL represents undefined.
 */
function getColAndRowFromA1(a1) {
  // Match the column part (letters) and row part (digits)
  let columnMatch = a1.match(/^[A-Za-z]+/);
  let rowMatch = a1.match(/\d+$/);

  // If a column is present, convert it to a numeric offset; otherwise, return null
  let numericCol = columnMatch ? letterToColOffset(columnMatch[0]) : null;
  // If a row is present, convert it to a numeric offset; otherwise, return null
  let numericRow = rowMatch ? parseInt(rowMatch[0], 10) - 1 : null;

  return [numericCol, numericRow];
}

function testGetRowAndColFromA1() {
  for (let a1 of ["A2", "3", "C", "D4", "ZA9", "BB"]) {
    console.log(a1, "=>", getColAndRowFromA1(a1));
  }
}

/**
 * @param {string} target - The A1 description of the target range.
 * @param {string} cell - The A1 description of a single cell.
 */
function a1rangeMatch(target, cell) {
  target = target.toUpperCase();
  cell = cell.toUpperCase();
  if (target === cell) {
    return true;
  }
  if (target.includes(":")) {
    let [start, end] = target.split(":");
    let [startCol, startRow] = getColAndRowFromA1(start);
    let [endCol, endRow] = getColAndRowFromA1(end);
    let [cellCol, cellRow] = getColAndRowFromA1(cell);
    if (startCol === null || startCol <= cellCol) {
      // Match start col
      if (endCol === null || endCol >= cellCol) {
        // match end col
        if (startRow === null || startRow <= cellRow) {
          // match start row
          if (endRow === null || endRow >= cellRow) {
            // match end Row
            return true;
          }
        }
      }
    }
  }
  return false;
}

function testA1rangeMatch() {
  for (let [cell, target] of [
    ["A3", "A3"],
    ["A3", "B3"],
    ["A3", "a3"],
    ["A3", "A:A"],
    ["A3", "3:3"],
    ["C4", "B:D"],
    ["C4", "3:5"],
    ["D7", "B5:D7"],
    ["D7", "B5:C7"],
  ]) {
    console.log("match", cell, "in", target, "=>", a1rangeMatch(target, cell));
  }
}

/**
 * Class representing a Spreadsheet UI framework.
 */
class SpreadsheetUI {
  /**
   * Creates an instance of SpreadsheetUI.
   */
  constructor() {
    /**
     * @type {Array<{cell: string, sheet: string, callback: function(Object):void, setup: function():void}>}
     */
    this.triggers = [];
    this.debug = true;
  }

  /**
   * Adds a trigger for a specific cell, range, or named range.
   * @param {Object} options - The options object.
   * @param {string} [options.cell] - The cell range in A1 notation.
   * @param {string} [options.sheet] - Name of the sheet (or undefined for any sheet).
   * @param {string} [options.rangeName] - The name of the named range.
   * @param {function(Object):void} options.callback - The function to execute when the cell is edited.
   * @param {Object} options.callback.params - Parameters passed to the callback function.
   * @param {string} options.callback.params.a1 - The A1 notation of the cell.
   * @param {any} options.callback.params.value - The value of the cell.
   * @param {SpreadsheetApp.Spreadsheet.Range} options.callback.params.range - The range object of the cell.
   * @param {Object} options.callback.params.params - An object containing additional parameters.
   * @param {Array} options.callback.params.params.cells - Array of triggered cells.
   * @param {SpreadsheetApp.Events.SheetsOnEdit} options.callback.params.event - The original edit event object.
   * @param {function():void} [options.setup] - Function to set up initial UI state.
   */
  addTriggerCell({ cell, sheet, rangeName, callback, setup = () => {} }) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (rangeName) {
      let namedRange = ss.getRangeByName(rangeName);
      if (namedRange) {
        cell = namedRange.getA1Notation();
        sheet = namedRange.getSheet().getName();
        if (this.debug) {
          console.log(
            `Using existing named range: ${rangeName} (${cell} in sheet ${sheet})`
          );
        }
      } else {
        if (!cell) {
          throw new Error(
            `Named range ${rangeName} does not exist, and cell not provided to create it.`
          );
        }
        try {
          const sheetObj = sheet
            ? ss.getSheetByName(sheet)
            : ss.getActiveSheet();
          namedRange = sheetObj.getRange(cell);
          ss.setNamedRange(rangeName, namedRange);
          sheet = sheetObj.getName();
          if (this.debug) {
            console.log(
              `Created new named range: ${rangeName} (${cell} in sheet ${sheet})`
            );
          }
        } catch (err) {
          console.error(
            "Unable to create trigger cell",
            cell,
            sheet,
            rangeName,
            callback
          );
          console.error(err);
          return false;
        }
      }
    } else {
      if (!sheet) {
        sheet = ss.getActiveSheet().getName();
      }
    }
    // Now our triggers continue to actually refer to cell and sheet, but we have found them
    // by pointing to named ranges...
    this.triggers.push({ cell, sheet, callback, setup });
  }

  /**
   * Adds a dropdown trigger for a specific cell, range, or named range.
   * @param {Object} options - The options object.
   * @param {string} [options.cell] - The cell range in A1 notation.
   * @param {string} [options.sheet] - Name of the sheet (or undefined for any sheet).
   * @param {string} [options.rangeName] - The name of the named range.
   * @param {Array<string>} options.values - Array of possible dropdown values.
   * @param {Object<string, function(Object):void>} options.callbacks - An object mapping dropdown values to callback functions.
   */
  addDropdown({ cell, sheet, rangeName, values, callbacks, setup = () => {} }) {
    const setupValidation = () => {
      const range = rangeName
        ? SpreadsheetApp.getActiveSpreadsheet().getRangeByName(rangeName)
        : sheet
        ? SpreadsheetApp.getActiveSpreadsheet()
            .getSheetByName(sheet)
            .getRange(cell)
        : SpreadsheetApp.getActiveSpreadsheet().getRange(cell);
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(values)
        .build();
      range.setDataValidation(rule);
    };

    const callback = (params) => {
      const selectedValue = params.value;
      if (selectedValue in callbacks) {
        console.log("Dropdown callback: ", selectedValue);
        callbacks[selectedValue](params);
      }
    };

    this.addTriggerCell({
      cell,
      sheet,
      rangeName,
      callback,
      setup: () => {
        setup();
        setupValidation();
      },
    });
  }

  /**
   * Adds a checkbox trigger for a specific cell, range, or named range.
   * @param {Object} options - The options object.
   * @param {string} [options.cell] - The cell range in A1 notation.
   * @param {string} [options.sheet] - Name of the sheet (or undefined for any sheet).
   * @param {string} [options.rangeName] - The name of the named range.
   * @param {function(Object):void} [options.onTrue] - The function to execute when the checkbox is checked (TRUE).
   * @param {function(Object):void} [options.onFalse] - The function to execute when the checkbox is unchecked (FALSE).
   */
  addCheckbox({ cell, sheet, rangeName, onTrue, onFalse, setup = () => {} }) {
    const setupValidation = () => {
      const range = rangeName
        ? SpreadsheetApp.getActiveSpreadsheet().getRangeByName(rangeName)
        : sheet
        ? SpreadsheetApp.getActiveSpreadsheet()
            .getSheetByName(sheet)
            .getRange(cell)
        : SpreadsheetApp.getActiveSpreadsheet().getRange(cell);
      range.insertCheckboxes();
    };

    const callback = (params) => {
      if (params.value && onTrue) {
        onTrue(params);
      } else if (!params.value && onFalse) {
        onFalse(params);
      }
    };

    this.addTriggerCell({
      cell,
      sheet,
      rangeName,
      callback,
      setup: () => {
        setup();
        setupValidation();
      },
    });
  }

  /**
   * Sets up the initial UI state by calling the setup function for each trigger.
   */
  setupUI() {
    this.triggers.forEach((trigger) => {
      if (trigger.setup) {
        trigger.setup();
      }
    });
  }

  /**
   * Handles the onEdit event.
   * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The event object.
   */
  onEdit(e) {
    let range = e.range; //SpreadsheetApp.getActiveRange();
    const editedCell = e.range.getA1Notation();
    const editedSheet = e.range.getSheet();
    const rangeHeight = e.range.getNumRows();
    const rangeWidth = e.range.getNumColumns();
    if (this.debug) {
      console.log("Event type", e.changeType);
      console.log("Event oldValue", e.oldValue);
      console.log("Event value", e.value);
    }
    /* Add type hinting for edited Cells based on below... */
    /**
     * @type {Array<{range: GoogleAppsScript.SpreadsheetApp.Range, a1 : string, value : string}>}
     */
    const editedCells = [];
    for (let col = 0; col < rangeWidth; col++) {
      for (let row = 0; row < rangeHeight; row++) {
        let cell = range.offset(row, col, 1, 1);
        editedCells.push({
          range: cell,
          a1: cell.getA1Notation(),
          value: cell.getValue(),
        });
      }
    }
    if (this.debug) {
      console.log(
        "onEdit got",
        editedCell,
        editedCells.map((c) => c.a1),
        e.oldValue
      );
    }
    let triggered = false;
    this.triggers.forEach((trigger) => {
      if (!trigger.sheet || trigger.sheet == editedSheet.getName()) {
        let triggeredCells = editedCells.filter((cell) =>
          a1rangeMatch(trigger.cell, cell.a1)
        );
        if (triggeredCells.length) {
          triggered = true;
          // One "param" object that will be passed to each
          // trigger and can do things like handle universal
          // confirmation of an action for all items in a trigger...
          let params = {
            cells: triggeredCells,
          };
          /* We are going to call the trigger *once* per item. */
          for (let cell of triggeredCells) {
            if (this.debug)
              console.log(
                "Callback for ",
                cell.a1,
                "from",
                trigger.cell,
                cell.value
              );
            trigger.callback({
              value: cell.value,
              range: cell.range,
              a1: cell.a1,
              params,
              event: e,
            });
          }
        }
      }
    });
    if (this.debug && !triggered) {
      console.log(
        "No rule matched from",
        this.triggers.map((t) => t.cell)
      );
    }
  }

  /**
   * Checks if an edited cell is within the trigger range.
   * @param {string} editedCell - The A1 notation of the edited cell.
   * @param {string} triggerRange - The A1 notation of the trigger range.
   * @return {boolean} True if the edited cell is within the trigger range, false otherwise.
   */
  isCellInRange(editedCell, triggerRange) {
    // Add logic to handle ranges, e.g., B1:B should match with B7
    // Simple range check logic, extend as needed
    let range = SpreadsheetApp.getActiveSpreadsheet().getRange(triggerRange);
    return (
      range.getA1Notation() === editedCell ||
      range.getA1Notation().includes(editedCell)
    );
  }
}

/** @type {SpreadsheetUI} */
var ui = new SpreadsheetUI();

function getUi() {
  if (!ui) {
    ui = new SpreadsheetUI();
  }
  return ui;
}

/**
 * Event handler for spreadsheet edits.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} editEvent - The event object.
 */
function onEvent(editEvent) {
  ui.onEdit(editEvent);
}

/**
 * Initializes triggers for the script.
 */
function initializeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let hasOnEditTrigger = false;

  // Check if an onEdit trigger already exists
  for (const trigger of triggers) {
    if (
      trigger.getHandlerFunction() === "onEvent" &&
      trigger.getEventType() === ScriptApp.EventType.ON_EDIT
    ) {
      hasOnEditTrigger = true;
      break;
    }
  }

  // If no onEdit trigger exists, create one
  if (!hasOnEditTrigger) {
    ScriptApp.newTrigger("onEvent")
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onEdit()
      .create();
  } else {
    console.log("onEdit trigger already exists.");
  }
}

/**
 * Authorizes the script and alerts the user.
 */
function authorize() {
  Browser.msgBox("Script is now authorized");
}

/**
 * Sets up the UI and initializes triggers.
 */
function setup() {
  ui.setupUI();
  initializeTriggers();
}

/**
 * @typedef {Object} DropdownOptions
 * @property {string} [cell] - The cell range in A1 notation.
 * @property {string} [sheet] - Name of the sheet (or undefined for any sheet).
 * @property {string} [rangeName] - The name of the named range.
 * @property {Array<string>} values - Array of possible dropdown values.
 * @property {Object<string, function(Object):void>} callbacks - An object mapping dropdown values to callback functions.
 */

/**
 * Adds a dropdown trigger for a specific cell, range, or named range.
 * @param {DropdownOptions} options - The dropdown object.
 */
function addDropdown(params) {
  let ui = getUi();
  ui.addDropdown(params);
}

/**
 * Adds a checkbox trigger for a specific cell, range, or named range.
 * @param {Object} options - The options object, containing cell, sheet, rangeName, onTrue and onFalse
 * @param {string} [options.cell] - The cell range in A1 notation.
 * @param {string} [options.sheet] - Name of the sheet (or undefined for any sheet).
 * @param {string} [options.rangeName] - The name of the named range.
 * @param {function(Object):void} [options.onTrue] - The function to execute when the checkbox is checked (TRUE).
 * @param {function(Object):void} [options.onFalse] - The function to execute when the checkbox is unchecked (FALSE).
 */
function addCheckbox({ cell, sheet, rangeName, onTrue, onFalse }) {
  let ui = getUi();
  ui.addCheckbox({ cell, sheet, rangeName, onTrue, onFalse });
}

/**
 * Adds a trigger for a specific cell, range, or named range.
 * @param {Object} options - The options object.
 * @param {string} [options.cell] - The cell range in A1 notation.
 * @param {string} [options.sheet] - Name of the sheet (or undefined for any sheet).
 * @param {string} [options.rangeName] - The name of the named range.
 * @param {function(Object):void} options.callback - The function to execute when the cell is edited.
 * @param {Object} options.callback.params - Parameters passed to the callback function.
 * @param {string} options.callback.params.a1 - The A1 notation of the cell.
 * @param {any} options.callback.params.value - The value of the cell.
 * @param {GoogleAppsScript.Spreadsheet.Range} options.callback.params.range - The range object of the cell.
 * @param {Object} options.callback.params.params - An object containing additional parameters.
 * @param {Array} options.callback.params.params.cells - Array of triggered cells.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} options.callback.params.event - The original edit event object.
 * @param {function():void} [options.setup] - Function to set up initial UI state.
 */
function addTriggerCell(params) {
  let ui = getUi();
  ui.addTriggerCell(params);
}
