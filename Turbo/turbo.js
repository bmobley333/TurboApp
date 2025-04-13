// turbo.gs //


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
// === Get Specific Range Data                         (End of Get Game Sheet data/styles) ===
// ==========================================================================


// fCSGetRangeData /////////////////////////////////////////////////////////////////////////////////
// Purpose -> Fetches values from a specified A1 range within the "Game" sheet of a given spreadsheet.
// Inputs  -> targetSheetId (String): The ID of the target Google Spreadsheet.
//         -> a1Notation (String): The standard A1 notation for the range (e.g., "C5:F10").
// Outputs -> (Array[][]|Error): A 2D array of values from the specified range, or throws an error on failure.
function fCSGetRangeData(targetSheetId, a1Notation) {
    Logger.log(`fCSGetRangeData called for SheetID: ${targetSheetId}, Range: ${a1Notation}`); // Removed debug
    try {
        // Validate inputs
        if (!targetSheetId || typeof targetSheetId !== 'string') {
            throw new Error("Invalid or missing Sheet ID provided.");
        }
        if (!a1Notation || typeof a1Notation !== 'string' || !a1Notation.includes(':')) {
            // Basic check for A1 format - improve if needed
            throw new Error("Invalid or missing A1 notation provided.");
        }

        // Open sheet (reuse existing helper if suitable, or inline logic)
        const ss = SpreadsheetApp.openById(targetSheetId);
        const sh = ss.getSheetByName("Game");
        if (!sh) {
            throw new Error(`Sheet named "Game" not found in SheetID: ${targetSheetId}.`);
        }

        // Get range and values
        const range = sh.getRange(a1Notation);
        const values = range.getValues();

        Logger.log(`fCSGetRangeData successfully retrieved ${values.length} rows.`); // Removed debug
        return values; // Return the 2D array

    } catch (e) {
        // Log detailed error and re-throw for client-side failure handler
        const errMsg = `fCSGetRangeData Error: Failed to get range "${a1Notation}" from sheet "${targetSheetId}". Details: ${e.message}`;
        console.error(errMsg + "\nStack:\n" + e.stack);
        throw new Error(errMsg); // Throw error back to client
    }
} // END fCSGetRangeData


