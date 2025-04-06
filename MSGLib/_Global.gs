////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                   Global Initialize ()
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// Global Variables

// Note: make all g.key1.key2 lower case not cammel case for easy ss.toLowerCase() matching

const g = {
  id: {
    db: "19jvxN_ZRWlPxbbCrcVmpa_Gk6PExT5cL134aokFkz3U",
    mastercs: "1nRNJyQil3HmUBJycMHUsQveV3SbSRtL9a046RCp0zpw",
    masterkl: "1vBvok8EAwDnJFg39SW0UNEj2fw3ixAlEESVRO72pxxk",
    ps: "18UduvedsXFztR1Q1WpWvBIuWdf8WRs5adrKAf5iZaEQ",
    //mycs - added on onOpen trigger, fCSRunMenuOrButton(), fKLRunMenuOrButton(), fCSCreateMenu(), and fKLCreateMenu() if the active sheet is mycs or mykl
    //mykl - added on onOpen trigger, fCSRunMenuOrButton(), fKLRunMenuOrButton(), fCSCreateMenu(), and fKLCreateMenu() if the active sheet is mycs or mykl

    cm: "1eo-ufGNMWOQnNuZ6idxUaKxqzWmxKAtaZ52Y8eb3JS4",
    em: "1Ic5EhcdNwJrPBcNVysRta23-adzWHw8yQvxfcy20WJE",
    rb: "1Fa5nE_CpgujjKEsc9jEN3rpo5AP1lxZjEkKmI30GZ9o",
    sg: "1InZEV8MtY9j7cuzcyrCIBuSd9_t1L9MREBp7tJmR34k",
  },
  db: {},
  mastercs: {},
  masterkl: {},
  ps: {},
  mycs: {},
  mykl: {},
  grnBS: '<span style="color: green; font-weight: bold;">',
  redItS: '<span style="color: red; font-style: italic;">',
  redBS: '<span style="color: red; font-weight: bold;">',
  bluBS: '<span style="color: blue; font-weight: bold;">',
  treS: '<span style="background-color: yellow; color: green; font-weight: bold;">',
  endS: "</span>",
};

// Including the above, other common nested objects of g:
// Note: make all g.key1.key2 lower case not cammel case for easy ss.toLowerCase() matching
// g.id[sheetname]             (e.g. g.id.mycs)
// g.obj[ss][sheetname][keys]  (e.g. g.obj.db.gmscreen)
// g.mycs[keys]
// g.mastercs[keys]
// g.db[keys]
// g.mykl[keys]
// g.masterkl[keys]
// g.ps[keys]

// onOpen //////////////////////////////////////////////////////////////////////////////////////////////////
// onOpen trigger to determine the active sheet and create the corresponding menu
function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();

  // Note: make all g.key1.key2 lower case not cammel case for easy ss.toLowerCase() matching
  // if spreadsheetId is not already in g.id then we know it is either mycs or mykl so add these id's to g.id
  if (!Object.values(g.id).includes(spreadsheetId)) {
    g.id.mycs = ss.getSheetByName("Data").getRange("F7").getValue();
    g.id.mykl = ss.getSheetByName("Data").getRange("F8").getValue();
  }
  switch (spreadsheetId) {
    case g.id.mastercs:
    case g.id.mycs:
      fCSCreateMenu();
      break;

    case g.id.db:
      fDBCreateMenu();
      break;

    case g.id.masterkl:
    case g.id.mykl:
      fKLCreateMenu();
      fKLCheckShowReminder();
      break;

    case g.id.ps:
      fPSCreateMenu();
      break;
    default:
      throw new Error(
        `In onOpen, I don't recognize the passed spreadsheetId of "${spreadsheetId}"`
      );
  }
} // End onOpen

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  Global Functions  (end Global Initialize)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// gURLFromID //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Returns the full URL of a given ID
// type can be 'sheet' or 'doc' and defaults to sheet
function gURLFromID(id, type = "sheet") {
  // If the type is not 'sheet' or 'doc', throw an error
  if (type.toLowerCase() !== "sheet" && type.toLowerCase() !== "doc")
    throw new Error(
      `In gURLFromID, the type of Google File is not understood: "${type}"`
    );

  // URL prefix based on the type
  const urlPrefix =
    type.toLowerCase() === "sheet"
      ? "https://docs.google.com/spreadsheets/d/"
      : "https://docs.google.com/document/d/";

  // URL suffix
  const urlSuffix = "/edit?usp=sharing";

  // Construct the full URL
  return urlPrefix + id + urlSuffix;
} // End gURLFromID

// gCharLvl //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Returns the character's current level as listed on CS <RaceClass> or assigns 1 if invalid
function gCharLvl() {
  let charLevel = gGetVal("mycs", "RaceClass", "level", "Val");

  // Check if charLevel is not a valid integer > 0, and assign it to 1 if it isn't
  if (!Number.isInteger(charLevel) || charLevel <= 0) {
    charLevel = 1;
  }

  return charLevel;
} // End gCharLvl

// gGetVal //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: if Table not already loaded, then LoadTable then...
// Purpose: get a value from sheetName based on header and key
// WARNING: Never use integer's as keys or headers as the code can't tell if this is a specific r/c  or a text key/column
function gGetVal(ss, sheetName, keyOrR, headerOrC) {
  ss = ss.toLowerCase();
  sheetName = sheetName.toLowerCase();

  // Load Table if not already loaded
  gLoadTable(ss, sheetName);

  const arr = g[ss][sheetName].arr;

  // Determine r based on whether key is a positive integer or not
  let r;
  if (Number.isInteger(keyOrR) && keyOrR >= 0) {
    r = keyOrR;
  } else {
    r = g[ss][sheetName].keyOf[keyOrR];
  }

  // Determine c based on whether key is a positive integer or not
  let c;
  if (Number.isInteger(headerOrC) && headerOrC >= 0) {
    c = headerOrC;
  } else {
    c = g[ss][sheetName].headerOf[headerOrC];
  }

  return arr[r][c];
} // End gGetVal

// gSetVal //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: if Table not already loaded, then LoadTable then...
// Purpose: set a value in sheetName based on header and keyOrR (key or row index)
function gSetVal(ss, sheetName, keyOrR, headerOrC, value) {
  ss = ss.toLowerCase();
  sheetName = sheetName.toLowerCase();

  // Load Table if not already loaded
  gLoadTable(ss, sheetName);

  const arr = g[ss][sheetName].arr;

  // Determine r based on whether key is a positive integer or not
  let r;
  if (Number.isInteger(keyOrR) && keyOrR >= 0) {
    r = keyOrR;
  } else {
    r = g[ss][sheetName].keyOf[keyOrR];
  }

  // Determine c based on whether key is a positive integer or not
  let c;
  if (Number.isInteger(headerOrC) && headerOrC >= 0) {
    c = headerOrC;
  } else {
    c = g[ss][sheetName].headerOf[headerOrC];
  }

  arr[r][c] = value;
} // End gSetVal

// gSaveValIndirect //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Finds the listed sheetname, Row,  and Col in <Data> and saves val to that sheet and location
function gSaveValIndirect(ss, dataKey, val) {
  const sheetName = gGetVal(ss, "data", dataKey, "SheetName");
  const r = gGetVal(ss, "data", dataKey, "Row");
  const c = gGetVal(ss, "data", dataKey, "Col");
  gSaveVal(ss, sheetName, r, c, val);
} // End gSaveValIndirect

// gSaveVal //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: if Table not already loaded, then LoadTable then...
// Purpose: Save a value to both the in-memory array and the actual Google Sheet
function gSaveVal(ss, sheetName, keyOrR, headerOrC, value) {
  ss = ss.toLowerCase();
  sheetName = sheetName.toLowerCase();

  // Load Table if not already loaded
  gLoadTable(ss, sheetName);

  const arr = g[ss][sheetName].arr;

  // Determine r based on whether key is a positive integer or not
  let r;
  if (Number.isInteger(keyOrR) && keyOrR >= 0) {
    r = keyOrR;
  } else {
    r = g[ss][sheetName].keyOf[keyOrR];
  }

  // Determine c based on whether key is a positive integer or not
  let c;
  if (Number.isInteger(headerOrC) && headerOrC >= 0) {
    c = headerOrC;
  } else {
    c = g[ss][sheetName].headerOf[headerOrC];
  }

  // Update the in-memory array
  arr[r][c] = value;

  // Update the actual Google Sheet
  var sheet = g[ss][sheetName].ref;
  sheet.getRange(r + 1, c + 1).setValue(value);
} // End gSaveVal

// gSaveRow //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: if Table not already loaded, then LoadTable then...
// Purpose: Save a row to both the in-memory array and the actual Google Sheet
function gSaveRow(ss, sheetName, keyOrR, rowArr) {
  ss = ss.toLowerCase();
  sheetName = sheetName.toLowerCase();

  // Load Table if not already loaded
  gLoadTable(ss, sheetName);

  const arr = g[ss][sheetName].arr;

  // Determine r based on whether keyOrR is a positive integer or not
  let r;
  if (Number.isInteger(keyOrR) && keyOrR >= 0) {
    r = keyOrR;
  } else {
    r = g[ss][sheetName].keyOf[keyOrR];
  }

  // Update the in-memory array
  arr[r] = rowArr;
  // Update the actual Google Sheet

  var sheet = g[ss][sheetName].ref;
  sheet.getRange(r + 1, 1, 1, rowArr.length).setValues([rowArr]);
} // End gSaveRow

// gSaveSheet //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: if Table not already loaded, then LoadTable then...
// Purpose: Save the ss.sheetname.arr to the actual sheet
function gSaveSheet(ss, sheetName) {
  ss = ss.toLowerCase();
  sheetName = sheetName.toLowerCase();

  // Load Table if not already loaded
  gLoadTable(ss, sheetName);

  const sheetRef = g[ss][sheetName].ref;
  const sheetArr = g[ss][sheetName].arr;

  // Update the actual Google Sheet
  sheetRef
    .getRange(1, 1, sheetArr.length, sheetArr[0].length)
    .setValues(sheetArr);
} // End gSaveSheet

// gCopySheet //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Copies a sheet over another sheet even if between ss
function gCopySheet(ssFrom, sheetFrom, ssTo, sheetTo) {
  gCloneSheet(ssFrom, sheetFrom, ssTo, "GASTempSheet");
  var sourceClone = gSheetRef(ssTo, "GASTempSheet");

  // Create targetSheet if needed
  var targetSheet =
    gSheetRef(ssTo, sheetTo) || gSSRef(ssTo).insertSheet(sheetTo);

  // Clear and resize targetSheet
  targetSheet.clear();

  const sourceMaxRows = sourceClone.getMaxRows();
  const sourceMaxColumns = sourceClone.getMaxColumns();
  const targetMaxRows = targetSheet.getMaxRows();
  const targetMaxColumns = targetSheet.getMaxColumns();

  // Adjust the number of rows in the targetSheet to match the sourceSheet
  if (targetMaxRows > sourceMaxRows) {
    targetSheet.deleteRows(sourceMaxRows + 1, targetMaxRows - sourceMaxRows);
  } else if (targetMaxRows < sourceMaxRows) {
    targetSheet.insertRowsAfter(targetMaxRows, sourceMaxRows - targetMaxRows);
  }

  // Adjust the number of columns in the targetSheet to match the sourceSheet
  if (targetMaxColumns > sourceMaxColumns) {
    targetSheet.deleteColumns(
      sourceMaxColumns + 1,
      targetMaxColumns - sourceMaxColumns
    );
  } else if (targetMaxColumns < sourceMaxColumns) {
    targetSheet.insertColumnsAfter(
      targetMaxColumns,
      sourceMaxColumns - targetMaxColumns
    );
  }

  // Copy data and formatting from gasTempSheet to targetSheet
  sourceClone
    .getRange(1, 1, sourceMaxRows, sourceMaxColumns)
    .copyTo(targetSheet.getRange(1, 1), { contentsOnly: false });

  // Delete gasTempSheet and force load the new data
  gSSRef(ssTo).deleteSheet(sourceClone);
  gLoadTable(ssTo, sheetTo, true);
} // End gCopySheet

// gCloneSheet //////////////////////////////////////////////////////////////////////////////////////////////////
// Helper of gCopySheet
// Purpose: Clones a sheet to another
function gCloneSheet(ssFrom, sheetFrom, ssTo, sheetTo) {
  const toSSRef = gSSRef(ssTo);
  const fromSheetRef = gSheetRef(ssFrom, sheetFrom);

  // Delete the destination sheet if it exists
  let toSheetRef = gSheetRef(ssTo, sheetTo);
  if (toSheetRef) toSSRef.deleteSheet(toSheetRef);

  // Copy the source sheet to the destination and clear data validations
  toSheetRef = fromSheetRef.copyTo(toSSRef).setName(sheetTo);
  toSheetRef.getDataRange().clearDataValidations();

  // Set calculated values from source sheet to destination sheet
  const values = fromSheetRef.getDataRange().getDisplayValues();
  toSheetRef.getRange(1, 1, values.length, values[0].length).setValues(values);

  // Force a reload of the destination sheet
  gLoadTable(ssTo, sheetTo, true);
} // End gCloneSheet

// gSSRef //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: returns sheetName's reference
function gSSRef(ss) {
  ss = ss.toLowerCase();

  let ssRef;
  switch (ss) {
    case "mycs":
      ssRef = SpreadsheetApp.openById(g.id.mycs);
      break;
    case "mastercs":
      ssRef = SpreadsheetApp.openById(g.id.mastercs);
      break;
    case "mykl":
      ssRef = SpreadsheetApp.openById(g.id.mykl);
      break;
    case "masterkl":
      ssRef = SpreadsheetApp.openById(g.id.masterkl);
      break;
    case "ps":
      ssRef = SpreadsheetApp.openById(g.id.ps);
      break;
    case "db":
      ssRef = SpreadsheetApp.openById(g.id.db);
      break;
    default:
      throw new Error(`gSSRef was passed an unknown ss: "${ss}"`);
  }
  g[ss].ref = ssRef;

  return ssRef;
} // End gSSRef

// gSheetRef //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: returns sheetName's reference
function gSheetRef(ss, sheetName) {
  ss = ss.toLowerCase();
  sheetName = sheetName.toLowerCase();
  gLoadTable(ss, sheetName);
  return g[ss][sheetName].ref;
} // End gSheetRef

// gArr //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: returns sheetName's array
function gArr(ss, sheetName) {
  ss = ss.toLowerCase();
  sheetName = sheetName.toLowerCase();
  gLoadTable(ss, sheetName);
  return g[ss][sheetName].arr;
} // End gArr

// gHeaderRow //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: returns the row (a 1D array) containing the header labels in lower case
function gHeaderRow(ss, sheetName) {
  ss = ss.toLowerCase();
  sheetName = sheetName.toLowerCase();
  gLoadTable(ss, sheetName);
  const r = g[ss][sheetName].headerR;

  return g[ss][sheetName].arr[r].map((header) => header.toLowerCase());
} // End gHeaderRow

// gHeaderC //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: returns the header's array column
function gHeaderC(ss, sheetName, header) {
  ss = ss.toLowerCase();
  sheetName = sheetName.toLowerCase();
  gLoadTable(ss, sheetName);
  const headerIndex = g[ss][sheetName].headerOf[header];
  if (headerIndex === undefined)
    throw new Error(`In gHeaderC, there is no header: "${header}"`);
  return headerIndex;
} // End gHeaderC

// gTestHeaderC //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: returns the header's array column if header exists, otherwise returns false
function gTestHeaderC(ss, sheetName, header) {
  ss = ss.toLowerCase();
  sheetName = sheetName.toLowerCase();
  gLoadTable(ss, sheetName);
  const headerIndex = g[ss][sheetName].headerOf[header];
  return headerIndex === undefined ? false : headerIndex;
} // End gTestHeaderC

// gDataFirst_R //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: returns the array row index of the first data row
function gDataFirst_R(ss, sheetName) {
  ss = ss.toLowerCase();
  sheetName = sheetName.toLowerCase();
  gLoadTable(ss, sheetName);
  return g[ss][sheetName].dataFirst_R;
} // End gDataFirst_R

// gDataLast_R //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: returns the array row index of the last data row (which is arr.length -1 as set in fLoadTable)
function gDataLast_R(ss, sheetName) {
  ss = ss.toLowerCase();
  sheetName = sheetName.toLowerCase();
  gLoadTable(ss, sheetName);
  return g[ss][sheetName].dataLast_R;
} // End gDataLast_R

// gKeyR //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Returns the key's R (row number)
// Purpose: Returns false if the key is not found
function gKeyR(ss, sheetName, key) {
  ss = ss.toLowerCase();
  sheetName = sheetName.toLowerCase();
  gLoadTable(ss, sheetName); // Assumed to load or refresh the data structure

  // Check if 'keyOf' exists and key is present within it
  if (g[ss][sheetName].keyOf && key in g[ss][sheetName].keyOf) {
    return g[ss][sheetName].keyOf[key];
  } else {
    return false; // Return false if the key is not found or 'keyOf' is not defined
  }
} // End gKeyR

// gLoadTable //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: add to g: g.ss, g.ss.ref, g.ss.sheetname, g.ss.sheetname.ref, g.ss.sheetname.arr, g[ss][sheetName].dataFirst_R, g[ss][sheetName].dataLast_R, g[ss][sheetName].headerR, g[ss][sheetName].headerOf, g[ss][sheetName].keyOf, dataFirst_R
// Purpose: g.sheetname.headerRow, g.sheetname.headerR
// WARNING: The Header row MUST be row 0, and the dataFirst_R is the value in C2 (arr[1][2])
// Purpose: forceLoad will force a reload of a table even if it already exists in memory
// Purpose: Headers and keys may have mulitple values sperated by ',' such as Blue,Ver
function gLoadTable(ss, sheetName, forceLoad = false) {
  ss = ss.toLowerCase();
  sheetName = sheetName.toLowerCase();

  // Check if table is already loaded and forceLoad is not required
  if (!forceLoad && g[ss] && g[ss][sheetName] && g[ss][sheetName].arr) return;

  // Initialize g[ss] if it doesn't exist
  if (!g[ss]) g[ss] = {};

  // Clear sheetName in case this is a forced Get
  g[ss][sheetName] = {};

  // Get the sheet by ss and sheetName or make it if it doesn't exist
  // WARNING: CAN'T use gSheetRef as it calls gLoadTable in an endless loop!!!
  var sheet = (g[ss][sheetName].ref = gMakeSheet(ss, sheetName));

  // Get the sheet data as a 2D array
  const maxRows = sheet.getMaxRows();
  const maxColumns = sheet.getMaxColumns();
  const arr = (g[ss][sheetName].arr = sheet
    .getRange(1, 1, maxRows, maxColumns)
    .getValues());
  // const arr = g[ss][sheetName].arr = sheet.getDataRange().getValues();

  // Save header R index as g[ss][sheetName].headerR AND Save row 0 as g[ss][sheetName].headerRow
  g[ss][sheetName].headerR = 0;
  g[ss][sheetName].headerRow = arr[0];

  // Get Data's First Row _R
  g[ss][sheetName].dataFirst_R = arr[1][2];

  //Get Data's Last Row _R
  g[ss][sheetName].dataLast_R = g[ss][sheetName].arr.length - 1;

  // Populate the g[ss][sheetName].headerOf mapping
  g[ss][sheetName].headerOf = {};
  arr[0].forEach((header, col) => {
    const headers = String(header).split(","); // Ensure header is treated as a string
    headers.forEach((h) => {
      g[ss][sheetName].headerOf[h.trim()] = col; // Trim to remove any leading/trailing spaces
    });
  });

  // Populate the g[ss][sheetName].keyOf mapping
  g[ss][sheetName].keyOf = {};
  arr.slice(1).forEach((row, index) => {
    const keys = String(row[0]).split(","); // Ensure row[0] is treated as a string
    keys.forEach((key) => {
      g[ss][sheetName].keyOf[key.trim()] = index + 1; // Trim to remove any leading/trailing spaces
    });
  });
} // End gLoadTable

// gGetIDFromString //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Gets the last six characters from a string after trim. Returns false if there are not at least 6 characters after trim.
// Purpose: Does NOT verify the returning string is an actual ID
function gGetIDFromString(stringWithID) {
  const trimmedString = stringWithID.trim();
  if (trimmedString.length < 6) {
    return false;
  }
  return trimmedString.slice(-6);
} // End gGetIDFromString

// gScrollToCell //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Scrolls to a particular cell in the active sheet, based on targetRow and targetCol.
function gScrollToCell(targetRow, targetCol) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.setActiveRange(sheet.getRange(targetRow + 1, targetCol + 1)); // Select the specified cell at targetRow and targetCol
} // END gScrollToCell

// gTestID //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Verifies that testNameIDorID IS a key or ends in a key found in ss.sheetname. Returns the ID if successful else returns false;
function gTestID(ss, sheetName, testNameIDorID) {
  // If testStr is not a string or is less than 6 characters long, return false
  if (typeof testNameIDorID !== "string" || testNameIDorID.length < 6)
    return false;

  // Remove any spaces, tabs, and end-of-line characters from the end of testNameIDorID
  const cleanedTestNameIDorID = testNameIDorID.trim().replace(/\s+$/, "");

  const testID = cleanedTestNameIDorID.slice(-6);

  // Convert ss and sheetName to lowercase
  ss = ss.toLowerCase();
  sheetName = sheetName.toLowerCase();

  gLoadTable(ss, sheetName);

  // Check if the key exists in the table
  if (!g[ss][sheetName].keyOf || !g[ss][sheetName].keyOf.hasOwnProperty(testID))
    return false;

  // Return the testID if the key exists
  return testID;
} // End gTestID

// gMakeSheet //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Returns the sheetName's ref from ss, also sets g.ss.ref to ss reference
function gMakeSheet(ss, sheetName) {
  if (sheetName === null || sheetName === "")
    throw new Error("gGetSheetRef was passed an empty sheetName.");

  ss = ss.toLowerCase();
  sheetName = sheetName.toLowerCase();

  // Create g.ss if doesn't exist
  if (!g[ss]) g[ss] = {};
  if (!g[ss][sheetName]) g[ss][sheetName] = {};

  const ssRef = (g[ss].ref = gSSRef(ss));

  // Check if the sheet exists; if not, clone 'BlankSheet' as a new sheet
  if (!ssRef.getSheetByName(sheetName))
    throw new Error(
      `In gMakeSheet, the sheet named "${sheetName}" doesn't exist.`
    );

  const sheetRef = (g[ss][sheetName].ref = ssRef.getSheetByName(sheetName));

  return sheetRef;
} // End gMakeSheet

// gDeleteEmptyCol0DataRows //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Deletes rows with no value (ID) in column 0
function gDeleteEmptyCol0DataRows(ss, sheetName) {
  ss = ss.toLowerCase();
  sheetName = sheetName.toLowerCase();
  gLoadTable(ss, sheetName); // Load or refresh the data structure

  let sheetRef = gSheetRef(ss, sheetName);
  let firstData_R = gDataFirst_R(ss, sheetName);
  let arr = g[ss][sheetName].arr; // Assuming arr is an array of arrays representing the sheet's data

  // Iterate backward to avoid indexing issues after row deletions
  const emptyRowArr = [];
  for (let r = arr.length - 1; r >= firstData_R; r--) {
    if (!arr[r][0]) {
      // Check if the cell in column 0 is empty
      emptyRowArr.push(r + 1); // Spreadsheet indices start at 1
    }
  }

  // Process the array of empty rows to batch delete contiguous rows
  if (emptyRowArr.length > 0) {
    let top_R = emptyRowArr[0];
    let bottom_R = emptyRowArr[0];

    for (let i = 1; i < emptyRowArr.length; i++) {
      if (emptyRowArr[i] === bottom_R - 1) {
        bottom_R = emptyRowArr[i]; // Expand the current range of contiguous rows
      } else {
        // Once a non-contiguous row is found, delete the previous contiguous block
        sheetRef.deleteRows(bottom_R, top_R - bottom_R + 1);

        // Reset the range to the current non-contiguous row
        top_R = emptyRowArr[i];
        bottom_R = emptyRowArr[i];
      }
    }

    // Delete the final block of contiguous rows
    sheetRef.deleteRows(bottom_R, top_R - bottom_R + 1);
  }
} // End gDeleteEmptyCol0DataRows

// gFillArraySection //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Fills all of the array's section (a rectangular grid) to the Val specified
function gFillArraySection(arr, top_R, bottom_R, first_C, last_C, valToFill) {
  for (let r = top_R; r <= bottom_R; r++) {
    arr[r].fill(valToFill, first_C, last_C + 1);
  }
} // End gFillArraySection

// gSimpleName //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: removes any trailing '   ' followed by any other text from a string
function gSimpleName(name) {
  return String(name).replace(/\s{3}.*$/, "");
} // End gSimpleName

// gSortArraySection //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Sorts a specified section of a 2D array based on a specific column
// Parameters:
//   arr - The 2D array to be sorted
//   firstRow - The first row of the section to be sorted
//   lastRow - The last row of the section to be sorted
//   firstCol - The first column of the section to be sorted
//   lastCol - The last column of the section to be sorted
//   sortCol - The column to be sorted upon
function gSortArraySection(arr, firstRow, lastRow, firstCol, lastCol, sortCol) {
  // Validate sortCol
  if (sortCol < firstCol || sortCol > lastCol) {
    throw new Error(
      "Invalid sortCol: must be within the range of firstCol and lastCol"
    );
  }

  // Extract the section to be sorted
  let sortSection = arr
    .slice(firstRow, lastRow + 1)
    .map((row) => row.slice(firstCol, lastCol + 1));

  // Sort the extracted section
  sortSection.sort((a, b) => {
    const sortIndex = sortCol - firstCol; // Adjust sortCol to be relative to the extracted section

    if (a[sortIndex] === "" && b[sortIndex] === "") return 0; // Compare based on the sortCol column
    if (a[sortIndex] === "") return 1;
    if (b[sortIndex] === "") return -1;
    return a[sortIndex].localeCompare(b[sortIndex]);
  });

  // Place the sorted section back into the original array
  for (let r = 0; r < sortSection.length; r++) {
    for (let c = 0; c < sortSection[r].length; c++) {
      arr[firstRow + r][firstCol + c] = sortSection[r][c];
    }
  }
} // End gSortArraySection

// gSaveArraySectionToSheet //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Saves a specific section of an array to a Google Sheet, adjusting for 0-based array indices and 1-based sheet indices
function gSaveArraySectionToSheet(
  sheetRef,
  entireSheetArr,
  top_R,
  bottom_R_orArrLenthMinus1,
  first_C,
  last_C
) {
  // Calculate the number of rows and columns to update
  const bottom_R = bottom_R_orArrLenthMinus1;
  const numRows = bottom_R - top_R + 1;
  const numCols = last_C - first_C + 1;

  // Extract the relevant section of the array (adjust for 0-based index)
  // Adjust the slice operation to handle multiple columns (if applicable)
  const sectionToSave = entireSheetArr
    .slice(top_R, bottom_R + 1)
    .map((row) =>
      row ? row.slice(first_C, last_C + 1) : Array(numCols).fill("")
    );

  // Get the range in the sheet and set the values (adjust for 1-based index)
  const range = sheetRef.getRange(top_R + 1, first_C + 1, numRows, numCols);

  range.setValues(sectionToSave);
} // END gSaveArraySectionToSheet

// gHideAll //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Hides all sheets from sheet "Hide>" and to the right, then on remaining sheets, parses text in cell A1's notes and hides listed rows and columns (e.g. Hide: A-C,F,1,4,8-12)
function gHideAll(ss) {
  const ssRef = gSSRef(ss);

  // Hide all sheets starting with the sheet named "Hide>" and to the right
  const sheets = ssRef.getSheets();

  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() === "Hide>") {
      // Hide 'Hide>' and all subsequent sheets
      for (let j = i; j < sheets.length; j++) {
        sheets[j].hideSheet();
      }
      break; // Exit the loop once 'Hide>' is found and processed
    }
  }

  // Process remaining sheets
  sheets.forEach((sheet) => {
    if (!sheet.isSheetHidden()) {
      // Load the notes for cell A1
      const hideStr = sheet.getRange("A1").getNote();

      // Extract the relevant line and clean it
      const hideLine = hideStr
        .split("\n")
        .find((line) => line.startsWith("Hide:"));
      if (hideLine) {
        let cleanedHideStr = hideLine
          .replace("Hide:", "")
          .replace(/\s+/g, "")
          .replace(/^,|,$/g, ""); // Remove "Hide:", remove whitespace, remove trailing/ending ","
        cleanedHideStr.toUpperCase();
        const hideArr = cleanedHideStr.split(",");

        // Process each element in hideArr
        hideArr.forEach((element) => {
          if (/^\d+$/.test(element)) {
            // Single integer - hide that row
            sheet.hideRows(parseInt(element));
          } else if (/^\d+-\d+$/.test(element)) {
            // Integer range - hide all rows in that range
            const [start, end] = element.split("-").map(Number);
            sheet.hideRows(start, end - start + 1);
          } else if (/^[A-Z]+$/.test(element)) {
            // Single letter - hide that column
            const col = element.charCodeAt(0) - 64; // Convert letter to column number
            sheet.hideColumns(col);
          } else if (/^[A-Z]+-[A-Z]+$/.test(element)) {
            // Letter range - hide all columns in that range
            const [startCol, endCol] = element
              .split("-")
              .map((letter) => letter.charCodeAt(0) - 64);
            sheet.hideColumns(startCol, endCol - startCol + 1);
          }
        });
      }
    }
  });
} // END gHideAll

// gUn_HideAll //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Un-Hides all sheets rows and columns
function gUn_HideAll(ss) {
  const ssRef = gSSRef(ss);

  // Unhide all sheets
  const sheets = ssRef.getSheets();
  sheets.forEach((sheet) => {
    if (sheet.isSheetHidden()) sheet.showSheet();

    // Unhide all rows and columns in the sheet
    const maxRows = sheet.getMaxRows();
    const maxColumns = sheet.getMaxColumns();

    // Check if there are any hidden rows or columns to avoid unnecessary API calls
    if (sheet.isRowHiddenByUser(1)) sheet.showRows(1, maxRows);
    if (sheet.isColumnHiddenByUser(1)) sheet.showColumns(1, maxColumns);
  });
} // END gUn_HideAll

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                               CacheServices    (end Global Functions)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// gPutDocCache //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Saves a key and an object to the document's cache service by stringifying the object
function gPutDocCache(key, value) {
  // Stringify the value (object) before saving
  const stringValue = JSON.stringify(value);

  // Ensure the stringified value does not exceed the cache limit
  if (stringValue.length > 99000)
    throw new Error(
      "In gPutDocCache, the stringified value length exceeds 99,000 characters."
    );

  const cache = CacheService.getDocumentCache();
  cache.put(key, stringValue, 21600); // 21,600 seconds = 6 hours
} // end gPutDocCache

// gGetDocCache //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Retrieves a value by key from the document's cache service and parses it as an object
function gGetDocCache(key) {
  const cache = CacheService.getDocumentCache();
  const stringValue = cache.get(key);

  // If the value doesn't exist, return null or a default value
  if (!stringValue) return null;

  // Parse the stringified object back into a JavaScript object
  const value = JSON.parse(stringValue);

  return value;
} // end gGetDocCache
