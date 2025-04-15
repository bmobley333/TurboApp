// turbo.gs //
// 04.14.2025 //



// ==========================================================================
// === Server-Side Global Variables ===
// ==========================================================================



const gSrv = { // Using gSrv prefix for server-side globals
  ids: {
    sheets: {
      db: '1m7-VRDey6rPFDm_ZHQy4GOGNRxR0Th2IVY9TUVQ5DdU',
      mastercs: '1zVjGg2KKzEAdX6WezkatGPV0cahWXGIdlZASIUlhf_o',
      masterkl: '1Gcbnggc9UnQrzAcGwSkKdb_jtwxF3u8x1Fb-BdsECZ0',
      ps: '1cu4tsBeQg4l2czraYDjdVpn1zsIKPhYi5czbGf9ZeZg',
      sg: '190vk3Dcqdux_pFjuryvmMx2dWTnJyl23cZvTGhM3kgU'
      // Player-specific mycs/mykl will be handled dynamically
    },
    docs: {
      cm: '1UuS-329gRi012k5nVtpmDltvrCO0yP04hjDIVoyoFwo',
      em: '1YnTTuuOVaHRk2o_Wad2IPcIGaYQ-FYbK5H5GKuMDr6A',
      rb: '1Ug7AvfuF12iPuCSMsQA_JfxtRHaEfY88-KvrNJK58PU',
      sg: '190vk3Dcqdux_pFjuryvmMx2dWTnJyl23cZvTGhM3kgU'
    }
  },
  // Configuration values
  DATA_TAB_NAME: 'Data', // Name of the CS tab holding player-specific IDs
  MYKL_ID_CELL_A1: 'F8' // A1 notation for the cell containing MyKL ID in the Data tab
  // Add other server-side constants if needed
};



// ==========================================================================
// === Server Side Functions (turbo.gs) ===
// ==========================================================================


// fColToA1 ///////////////////////////////////////////////////////////////////
// Purpose -> Helper function to convert 0-based column index to A1 notation.
// Inputs  -> col (Number): The 0-based column index.
// Outputs -> (String): The A1 notation label (e.g., A, Z, AA).
function fColToA1(col) {
    let label = '';
    let c = col; // Use local variable

    // Loop through column value
    while (c >= 0) {
        label = String.fromCharCode((c % 26) + 65) + label;
        c = Math.floor(c / 26) - 1;
    }
    return label;
} // END fColToA1




// doGet //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Auto launches when Turbo web app is 1st called - it then loads index.html.
//            All based on a specific sheetId passed in the URL.
// Inputs  -> e (Event Object): Google Apps Script web app event object containing parameters.
// Outputs -> (HtmlOutput): The HTML page to be served.
function doGet(e) {

    // Get sheetId from URL parameter
    const sheetId = e.parameter.sheetId;
    if (!sheetId) {
        return HtmlService.createHtmlOutput("❌ sheetId parameter not provided in URL.");
    }

    // Attempt to load and serve the main HTML template
    try {
        // Optional: Verify sheetId is valid before rendering (can be slow)
        // SpreadsheetApp.openById(sheetId);

        // Create template from index.html
        const template = HtmlService.createTemplateFromFile("index");
        template.sheetId = sheetId; // Pass sheetId to the template client-side

        // Evaluate and return the HTML
        return template.evaluate()
            .setTitle("Turbo")
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    } catch (error) {
        // Log detailed error and return user-friendly error message
        console.error(`doGet Error for sheetId ${sheetId}: ${error} ${error.stack ? '\n' + error.stack : ''}`);
        return HtmlService.createHtmlOutput(`❌ Error loading app: ${error.message}. Is sheetId '${sheetId}' valid and accessible?`);
    }
} // END doGet




// ==========================================================================
// === Get Game Sheet data/styles                          (End of doGet) ===
// ==========================================================================



////////////////////////////////////////////////////////////////////////////////////////////////////////// START fCSGetGameSheet and helpers




// fCSGetGameSheet5 /////////////////////////////////////////////////////////////////////////////////
// Loads full data, format, and notes from 'Game' sheet. Returns structured object.
function fCSGetGameSheet5(targetSheetId) {
    try {
        const sh = fOpenGameSheet(targetSheetId);
        return fExtractSheetData(sh);
    } catch (e) {
        const context = e.stack?.includes("fOpenGameSheet") ? "opening spreadsheet" : "processing sheet";
        const msg = `fCSGetGameSheet5: Error ${context}: ${e.message}`;
        console.error(msg + "\nStack:\n" + e.stack);
        throw new Error(msg);
    }
} // END fCSGetGameSheet5




// fOpenGameSheet ///////////////////////////////////////////////////////////////////////////////////
// Opens spreadsheet and retrieves the "Game" sheet, or throws with clear message.
function fOpenGameSheet(targetSheetId) {
    if (!targetSheetId || typeof targetSheetId !== 'string') {
        throw new Error("Invalid or missing Sheet ID provided.");
    }

    const ss = SpreadsheetApp.openById(targetSheetId);
    const sh = ss.getSheetByName("Game");

    if (!sh) {
        throw new Error(`Sheet named "Game" not found.`);
    }

    return sh;
} // END fOpenGameSheet



// fIsSheetTrulyEmpty ////////////////////////////////////////////////////////////////////////////////
// Determines if the sheet has zero real content, returns true if it's blank.
function fIsSheetTrulyEmpty(sh, arr) {
    const numRows = arr.length;
    const numCols = arr[0]?.length || 0;
    const onlyEmpty = numRows === 1 && numCols === 1 && arr[0][0] === "";
    return (numRows === 0 || numCols === 0 || onlyEmpty) &&
           (sh.getLastRow() === 0 && sh.getLastColumn() === 0);
} // END fIsSheetTrulyEmpty



// fBuildFormatObject ////////////////////////////////////////////////////////////////////////////////
// Constructs formatting object from raw arrays and merge info.
function fBuildFormatObject(sh, rngData, numCols) {
    const fontColorObjects = rngData.getFontColorObjects();
    const fontColorHex = fontColorObjects.map(row =>
      row.map(obj => obj?.asRgbColor?.()?.asHexString?.()?.replace(/^#ff/, '#') ?? null)
    );

    const colWidths = Array.from({ length: numCols }, (_, c) =>
        sh.getColumnWidth(c + 1)
    );

    const mergedRanges = rngData.getMergedRanges();
    const merges = mergedRanges.map(r => ({
        row: r.getRow() - 1,
        col: r.getColumn() - 1,
        rowspan: r.getNumRows(),
        colspan: r.getNumColumns()
    })).filter(m => Number.isInteger(m.row) && Number.isInteger(m.col) && (m.rowspan > 1 || m.colspan > 1));

    return {
        bg: rngData.getBackgrounds(),
        fontColorHex: fontColorHex,
        weight: rngData.getFontWeights(),
        fontSize: rngData.getFontSizes(),
        fontStyle: rngData.getFontStyles(),
        fontFamily: rngData.getFontFamilies(),
        wrap: rngData.getWraps(),
        merges: merges,
        colWidths: colWidths,
        borders: [] // client-side handles borders
    };
} // END fBuildFormatObject




// fBuildReturnObject ////////////////////////////////////////////////////////////////////////////////
// Combines the final return structure.
function fBuildReturnObject(arr, format, notesArr) {
    return {
        arr: arr,
        format: format,
        notesArr: notesArr
    };
} // END fBuildReturnObject




// fExtractSheetData ////////////////////////////////////////////////////////////////////////////////
// Reads data, formats, and notes from the sheet and returns structured result.
function fExtractSheetData(sh) {
    const rngData = sh.getDataRange();
    const arr = rngData.getValues();

    if (fIsSheetTrulyEmpty(sh, arr)) {
        return fBuildReturnObject(
            [[]],
            {
                bg: [[]], fontColorHex: [[]], weight: [[]], fontSize: [[]],
                fontStyle: [[]], fontFamily: [[]], wrap: [[]], merges: [],
                colWidths: [], borders: []
            },
            [[]]
        );
    }

    const numCols = arr[0]?.length || 0;
    const format = fBuildFormatObject(sh, rngData, numCols);
    const notesArr = rngData.getNotes();

    return fBuildReturnObject(arr, format, notesArr);
} // END fExtractSheetData




////////////////////////////////////////////////////////////////////////////////////////////////////////// END fCSGetGameSheet and helpers

// ==========================================================================
// === Read/Write to Google Sheets/Docs                        (End of Get Game Sheet data/styles) ===
// ==========================================================================




// fGetMyKlId //////////////////////////////////////////////////////////////////
// Purpose -> Retrieves the MyKL ID from the player's Character Sheet's <Data> tab.
// Inputs  -> myCsId (String): The Sheet ID of the player's Character Sheet (mycs).
// Outputs -> (String): The player's MyKL Sheet ID.
// Throws  -> (Error): If IDs cannot be retrieved or sheets/tabs not found.
function fGetMyKlId(myCsId) {
  // Validate input
  if (!myCsId || typeof myCsId !== 'string') {
    console.error("fGetMyKlId Error: Character Sheet ID (myCsId) was not provided or invalid.");
    throw new Error("getMyKlId: Character Sheet ID (myCsId) was not provided.");
  }
  Logger.log(`fGetMyKlId: Attempting to get MyKL ID from CS ID: ${myCsId}`);

  try {
    // Open the Character Sheet using the provided ID
    const csSpreadsheet = SpreadsheetApp.openById(myCsId);
    if (!csSpreadsheet) {
        throw new Error(`Could not open Character Sheet with ID: ${myCsId}. Check permissions and ID validity.`);
    }

    // Get the specific sheet named in gSrv.DATA_TAB_NAME
    const dataSheet = csSpreadsheet.getSheetByName(gSrv.DATA_TAB_NAME);
    if (!dataSheet) {
      throw new Error(`Sheet named "${gSrv.DATA_TAB_NAME}" not found in Character Sheet ID: ${myCsId}.`);
    }

    // Get the MyKL ID from the specified cell
    const myKlId = dataSheet.getRange(gSrv.MYKL_ID_CELL_A1).getValue();

    // Validate the retrieved MyKL ID
    if (!myKlId || typeof myKlId !== 'string' || myKlId.trim() === '') {
       throw new Error(`Could not retrieve a valid MyKL ID from cell ${gSrv.MYKL_ID_CELL_A1} in sheet "${gSrv.DATA_TAB_NAME}". Value was: "${myKlId}"`);
    }

    Logger.log(`fGetMyKlId: Successfully retrieved MyKL ID: ${myKlId}`);
    return myKlId.trim(); // Return the trimmed ID

  } catch (e) {
    // Log detailed error and re-throw a user-friendly error
    console.error(`Error in fGetMyKlId for CS ID ${myCsId}: ${e.message}\nStack: ${e.stack}`);
    // Include specific error message if available, otherwise generic
    throw new Error(`Server error getting MyKL ID: ${e.message || e}`);
  }

} // END fGetMyKlId



// fGetServerSheetData /////////////////////////////////////////////////////////
// Purpose -> Reads data from a specified range in a given Google Sheet.
//            Can read from a specific sheet name or the first/active sheet.
// Inputs  -> fileId (String): The ID of the Google Sheet.
//            rangeA1 (String): The A1 notation of the range to read.
//            sheetName (String, Optional): The name of the specific sheet to read from.
//                                          If omitted/null/empty, reads from the first sheet.
// Outputs -> (Array[][]): The values from the specified range.
// Throws  -> (Error): If inputs are invalid or sheet/range access fails.
function fGetServerSheetData(fileId, rangeA1, sheetName = null) {
  // === 1. Validate Inputs ===
  if (!fileId || typeof fileId !== 'string') {
    console.error("fGetServerSheetData Error: Invalid or missing Sheet ID provided.");
    throw new Error("getServerSheetData: Invalid or missing Sheet ID provided.");
  }
  if (!rangeA1 || typeof rangeA1 !== 'string' || !rangeA1.includes(':')) {
    console.error(`fGetServerSheetData Error: Invalid or missing A1 notation: "${rangeA1}"`);
    throw new Error("getServerSheetData: Invalid or missing A1 notation provided.");
  }
  // Optional sheetName validation (e.g., check if string if provided)
  const targetSheetLog = sheetName ? `sheet "${sheetName}"` : "first/active sheet";
  Logger.log(`fGetServerSheetData: Reading range "${rangeA1}" from ${targetSheetLog} in Sheet ID: ${fileId}`);

  try {
    // === 2. Open Spreadsheet ===
    const ss = SpreadsheetApp.openById(fileId);
    if (!ss) {
        throw new Error(`Could not open spreadsheet with ID: ${fileId}. Check permissions and ID validity.`);
    }

    // === 3. Get Target Sheet ===
    let sh;
    if (sheetName && typeof sheetName === 'string' && sheetName.trim() !== '') {
        sh = ss.getSheetByName(sheetName.trim());
        if (!sh) {
            throw new Error(`Sheet named "${sheetName.trim()}" not found in SheetID: ${fileId}.`);
        }
    } else {
        // Default to the first sheet if no valid name provided
        sh = ss.getSheets()[0];
        if (!sh) {
            // Should be rare unless the sheet is completely empty
             throw new Error(`Could not get the first sheet in SheetID: ${fileId}.`);
        }
    }

    // === 4. Get Range and Values ===
    const range = sh.getRange(rangeA1); // Might throw if A1 notation is invalid
    if (!range) {
        // Redundant check, but safe practice
        throw new Error(`Could not find range "${rangeA1}" in sheet "${sh.getName()}" of ID: ${fileId}.`);
    }
    const values = range.getValues();

    Logger.log(`fGetServerSheetData: Successfully retrieved ${values?.length ?? 0} rows for range ${rangeA1} from sheet "${sh.getName()}".`);
    return values; // Return the 2D array of values

  } catch (e) {
    // Log detailed error and re-throw a user-friendly error
    console.error(`Error in fGetServerSheetData for ID ${fileId}, Range ${rangeA1}, Sheet "${sheetName}": ${e.message}\nStack: ${e.stack}`);
    throw new Error(`Server error reading sheet data: ${e.message || e}`);
  }

} // END fGetServerSheetData




// fSetServerSheetData /////////////////////////////////////////////////////////
// Purpose -> Writes data to a specified range in a given Google Sheet.
// Inputs  -> fileId (String): The ID of the Google Sheet.
//            rangeA1 (String): The A1 notation of the range to write to.
//            values (Array[][]): The 2D array of values to write.
// Outputs -> (Boolean): True on success.
// Throws  -> (Error): If inputs are invalid, sheet/range access fails, or
//                     value dimensions don't match range dimensions.
function fSetServerSheetData(fileId, rangeA1, values) {
  // === 1. Validate Inputs ===
  if (!fileId || typeof fileId !== 'string') {
    console.error("fSetServerSheetData Error: Invalid or missing Sheet ID provided.");
    throw new Error("setServerSheetData: Invalid or missing Sheet ID provided.");
  }
  if (!rangeA1 || typeof rangeA1 !== 'string' || !rangeA1.includes(':')) {
    console.error(`fSetServerSheetData Error: Invalid or missing A1 notation: "${rangeA1}"`);
    throw new Error("setServerSheetData: Invalid or missing A1 notation provided.");
  }
  // Validate that 'values' is a non-empty 2D array
  if (!Array.isArray(values) || values.length === 0 || !Array.isArray(values[0])) {
     console.error(`fSetServerSheetData Error: Invalid 'values' format for range ${rangeA1}. Expected non-empty 2D array. Received:`, values);
     throw new Error("setServerSheetData: Invalid 'values' format. Expected non-empty 2D array.");
  }
  Logger.log(`fSetServerSheetData: Writing ${values.length}x${values[0].length} data to range "${rangeA1}" in Sheet ID: ${fileId}`);

  try {
    // === 2. Open Sheet and Get Range ===
    const ss = SpreadsheetApp.openById(fileId);
    if (!ss) {
        throw new Error(`Could not open spreadsheet with ID: ${fileId}. Check permissions and ID validity.`);
    }
    // Assumes first/active sheet. Add sheet name logic if needed.
    const range = ss.getRange(rangeA1);
    if (!range) {
        throw new Error(`Could not find range "${rangeA1}" in the first sheet of ID: ${fileId}.`);
    }

    // === 3. Validate Dimensions ===
    const rangeNumRows = range.getNumRows();
    const rangeNumCols = range.getNumColumns();
    const valuesNumRows = values.length;
    const valuesNumCols = values[0].length;

    if (rangeNumRows !== valuesNumRows || rangeNumCols !== valuesNumCols) {
      throw new Error(`Dimension mismatch: Range "${rangeA1}" is ${rangeNumRows}x${rangeNumCols}, but provided data is ${valuesNumRows}x${valuesNumCols}.`);
    }

    // === 4. Write Values ===
    range.setValues(values);

    Logger.log(`fSetServerSheetData: Successfully wrote data to range ${rangeA1}.`);
    return true; // Indicate success

  } catch (e) {
    // Log detailed error and re-throw a user-friendly error
    console.error(`Error in fSetServerSheetData for ID ${fileId}, Range ${rangeA1}: ${e.message}\nStack: ${e.stack}`);
    throw new Error(`Server error writing sheet data: ${e.message || e}`);
  }

} // END fSetServerSheetData



// fGetDocumentText ////////////////////////////////////////////////////////////
// Purpose -> Reads the text content from a given Google Doc.
// Inputs  -> docId (String): The ID of the Google Doc.
// Outputs -> (String): The text content of the Doc's body.
// Throws  -> (Error): If docId is invalid or document access fails.
function fGetDocumentText(docId) {
  // Validate input
  if (!docId || typeof docId !== 'string') {
    console.error("fGetDocumentText Error: Invalid or missing Document ID provided.");
    throw new Error("getDocumentText: Invalid or missing Document ID provided.");
  }
  Logger.log(`fGetDocumentText: Reading text from Doc ID: ${docId}`);

  try {
    // Open the document
    const doc = DocumentApp.openById(docId);
    if (!doc) {
        throw new Error(`Could not open document with ID: ${docId}. Check permissions and ID validity.`);
    }

    // Get the text content from the body
    const textContent = doc.getBody().getText();

    Logger.log(`fGetDocumentText: Successfully read ${textContent?.length ?? 0} characters from Doc ID: ${docId}.`);
    return textContent; // Return the text content

  } catch (e) {
    // Log detailed error and re-throw a user-friendly error
    console.error(`Error in fGetDocumentText for ID ${docId}: ${e.message}\nStack: ${e.stack}`);
    throw new Error(`Server error reading document: ${e.message || e}`);
  }

} // END fGetDocumentText


