// turbo.gs //
// 04.29.2025 //
// Now using VS Code Post Mega Bad Branch 10 //
// v29.0 //



// ==========================================================================
// === Global Variables & doGet ===
// ==========================================================================


// NOTE gIndex must be passed into Turbo.gs from client side



const gSrv = { // Using gSrv prefix for server-side globals
  ids: {
    sheets: {
      db: '1W00TbYmt9cJpgBmRgWkIeek3G4TkXMJHMUHU1vBKkNU',
      mastercs: '19tPOKtbl5CioN0UVdZL76Ldg_xje2KYvhpgd9T56MxE',
      masterkl: '1ahFn32jaZGGggq6_QbSV_QBAcMGiUs5-TYtyNMDJ2d4',
      ps: '1EW5kF7k-_39bDPlAnKY-AbCgWdOE68d19BE6-FLbMmg',
      // Player-specific cs/kl will be handled dynamically via functions below
      cs: '',  // this is equal to the old myCS and is the ID from the player's CS (loaded via doGet)
      kl: ''  // this is equal to the old myKL and is the ID from the player's KL (loaded via doGet's call to fSrvGetMyKlId)
    },
    docs: {
      cm: '1V56MUMmSdGRGshieHZAVsvjMcx64e0_bz9okGz8XTfU',
      em: '1WJV_QZKqFqx83RDScFCEgL7trhpM7UVp16AIeVbgVDQ',
      rb: '1JZYdkvBDPtvGBhdk_FU2TRFUXD8CQ-cbJbPMRFs7Rm4',
      sg: '1sA749nIox0TWFMqka5sfGLpmHml_hEn8vBDxFpFVfC4'
    }
  },
  // Configuration values
  DATA_TAB_NAME: 'Data', // Name of the CS tab holding player-specific IDs
  MYKL_ID_CELL_A1: 'F8', // A1 notation for the cell containing MyKL ID in the Data tab
  designerPassword: 'Ogluck',
  // Add other server-side constants if needed
};




// ==========================================================================
// === Global Constants & Simple Utilities ===
// ==========================================================================


// fSrvConvertIndicesToA1 //////////////////////////////////////////////////////////
// Purpose -> Converts 0-based row/column indices to standard A1 or 'A1:B2' or 'A1' (if r1=r2 and c1=c2) notation string
//            on the server-side.
// Inputs  -> r1, c1, r2, c2 (Number): 0-based start/end row and column indices.
// Outputs -> (String | null): A1 notation string (e.g., "C5:F10"), or null on invalid input.
function fSrvConvertIndicesToA1(r1, c1, r2, c2) {
    // Validate inputs are non-negative numbers
    if ([r1, c1, r2, c2].some(idx => typeof idx !== 'number' || idx < 0 || isNaN(idx))) {
        console.error(`fSrvConvertIndicesToA1: Invalid indices provided (${r1},${c1},${r2},${c2})`);
        return null; // Return null for invalid indices
    }

    // Convert column indices to A1 letters using the existing helper
    const startColA1 = fSrvColToA1(c1);
    const endColA1 = fSrvColToA1(c2);

    // Add 1 to row indices for 1-based A1 notation
    const startRowA1 = r1 + 1;
    const endRowA1 = r2 + 1;

    // Determine if it's a single cell or a range
    if (r1 === r2 && c1 === c2) {
        // Single cell notation
        return `${startColA1}${startRowA1}`;
    } else {
        // Range notation: Top-left cell : Bottom-right cell
        return `${startColA1}${startRowA1}:${endColA1}${endRowA1}`;
    }
} // END fSrvConvertIndicesToA1




// fSrvColToA1 ///////////////////////////////////////////////////////////////////
// Purpose -> Helper function to convert 0-based column index to A1 notation.
// Inputs  -> col (Number): The 0-based column index.
// Outputs -> (String): The A1 notation label (e.g., A, Z, AA).
function fSrvColToA1(col) {
    let label = '';
    let c = col; // Use local variable

    // Loop through column value
    while (c >= 0) {
        label = String.fromCharCode((c % 26) + 65) + label;
        c = Math.floor(c / 26) - 1;
    }
    return label;
} // END fSrvColToA1




// ==========================================================================
// === Initialization / Bootstrapping ===
// ==========================================================================




// doGet ///////////////////////////////////////////////////////////////////////////
// Purpose -> Standard Apps Script function triggered when the web app URL is accessed.
//            Serves as the main server-side entry point.
// Process -> 1. Receives the request event object 'e', containing URL parameters.
//            2. Extracts and validates 'csID' (e.parameter.csID), 'userEmail' (e.parameter.userEmail), and 'gameVer' (e.parameter.gameVer).
//            3. Stores 'csID' in the server-side global gSrv.ids.sheets.cs. (Others not stored server-side).
//            4. Calls fSrvGetMyKlId(csID) to retrieve the associated KL ID ('klId') and store it in gSrv.ids.sheets.kl.
//            5. Creates an HTML template object from the 'index.html' file.
//            6. **Injects Data**: Assigns 'csID' to template.csID, 'userEmail' to template.userEmail,and 'gameVer' to template.gameVer.
//            7. Evaluates the 'index.html' template. During this evaluation:
//               - The scriptlet '<?= csID ?>' assigns the csID value to the client-side global variable 'gIndexCSID'.
//               - The scriptlet '<?= userEmail ?>' assigns the userEmail value to the client-side global variable 'gIndexEmail'.
//               - The scriptlet '<?= gameVer ?>' assigns the gameVer value to the client-side global variable 'gIndexGameVer'.
//               - All globals are then available to all .html files BUT NOT to turbo.gs.
//            8. Sets the browser window title to "MetaScape".
//            9. Returns the fully rendered HTML page to the client's browser.
// Inputs  -> e (Event Object): Contains request details, including URL parameters (e.g., e.parameter.csID, e.parameter.userEmail, e.parameter.gameVer).
// Outputs -> (HtmlOutput): The fully rendered HTML page to be displayed in the user's browser.
// Side Effects -> Populates gSrv.ids.sheets.cs and gSrv.ids.sheets.kl on the server-side.
//            -> Causes creation of gIndexCSID, gIndexEmail, and gIndexGameVer client-side globals.
// Availability -> 'gIndexCSID', 'gIndexEmail', 'gIndexGameVer' become available to *all* client-side JavaScript code
//                 (e.g. within index.html, scripts.html, gamelogic.html).
//            -> They are *not* directly accessible by server-side .gs code (like this function) but can be passed
//               from .html to .gs via parameters of google.script.run.
function doGet(e) {
    let csId = null; // Use let to allow modification in error handling
    let klId = null; // Use let for KL ID
    let userEmail = null; // Variable for user email
    let gameVer = null; // Variable for game version

    try {
        // --- 1. Get and Validate Parameters ---
        csId = e?.parameter?.csID; // Use optional chaining
        userEmail = e?.parameter?.userEmail; // Get userEmail parameter
        gameVer = e?.parameter?.gameVer; // Get gameVer parameter

        if (!csId || typeof csId !== 'string') {
            console.error("doGet Error: csID parameter not provided or invalid in URL.");
            return HtmlService.createHtmlOutput("❌ csID parameter missing or invalid in URL.");
        }
        // Optional: Validate userEmail format if necessary, but allow it to be potentially empty/null if not critical
        if (!userEmail || typeof userEmail !== 'string') {
            console.warn("doGet Warning: userEmail parameter not provided or invalid in URL.");
            userEmail = ''; // Default to empty string if missing or invalid
        }
        // Allow gameVer to be empty/null
        if (gameVer === null || gameVer === undefined) {
             console.warn("doGet Warning: gameVer parameter not provided in URL.");
             gameVer = ''; // Default to empty string if missing
        } else if (typeof gameVer !== 'string'){
             gameVer = String(gameVer); // Ensure it's a string
        }
        Logger.log(`doGet: Received CS ID: ${csId}, User Email: ${userEmail}, Game Ver: ${gameVer}`);

        // --- 2. Populate gSrv with CS ID ---
        gSrv.ids.sheets.cs = csId;
        Logger.log(`doGet: Assigned gSrv.ids.sheets.cs = ${gSrv.ids.sheets.cs}`);

        // --- 3. Get and Populate KL ID ---
        klId = fSrvGetMyKlId(csId); // This function handles its own internal errors/throws
        gSrv.ids.sheets.kl = klId;
        Logger.log(`doGet: Assigned gSrv.ids.sheets.kl = ${gSrv.ids.sheets.kl}`);

        // --- 4. Serve HTML ---
        const template = HtmlService.createTemplateFromFile("index");
        template.csID = csId; // Pass csId to the template client-side (becomes gIndexCSID)
        template.userEmail = userEmail; // Pass userEmail to the template client-side (becomes gIndexEmail)
        template.gameVer = gameVer; // Pass gameVer to the template client-side (becomes gIndexGameVer)

        // Evaluate and return the HTML
        return template.evaluate()
            .setTitle("MetaScape") // Updated Title
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    } catch (error) {
        // Log detailed error and return user-friendly error message
        const errorContext = `CS ID: ${csId || 'Unknown'}, Email: ${userEmail || 'Unknown'}, GameVer: ${gameVer || 'Unknown'}, KL ID Fetch Attempted: ${klId !== null || 'No (failed before KL fetch)'}`; // Added GameVer to context
        console.error(`doGet Error (${errorContext}): ${error.message}${error.stack ? '\n' + error.stack : ''}`);

        // Provide specific feedback based on error type
        let userErrorMessage = `❌ Error loading app: ${error.message}.`;
        if (error.message.includes("getMyKlId")) {
             userErrorMessage += ` Could not retrieve KL ID from CS ID '${csId}'. Check Data tab setup.`;
        } else if (error.message.includes("openById")) {
             userErrorMessage += ` Is csID '${csId}' valid and accessible?`;
        } else if (error.message.includes("getSheetByName")) {
             userErrorMessage += ` Ensure required tabs exist in sheet '${csId}'.`;
        }
        return HtmlService.createHtmlOutput(userErrorMessage);
    }
} // END doGet




// fSrvGetMyKlId //////////////////////////////////////////////////////////////////
// Purpose -> Retrieves the MyKL ID from the player's Character Sheet's <Data> tab.
// Inputs  -> myCsId (String): The Sheet ID of the player's Character Sheet (mycs).
// Outputs -> (String): The player's MyKL Sheet ID.
// Throws  -> (Error): If IDs cannot be retrieved or sheets/tabs not found.
function fSrvGetMyKlId(myCsId) {
  // Validate input
  if (!myCsId || typeof myCsId !== 'string') {
    console.error("fSrvGetMyKlId Error: Character Sheet ID (myCsId) was not provided or invalid.");
    throw new Error("getMyKlId: Character Sheet ID (myCsId) was not provided.");
  }
  Logger.log(`fSrvGetMyKlId: Attempting to get MyKL ID from CS ID: ${myCsId}`);

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

    Logger.log(`fSrvGetMyKlId: Successfully retrieved MyKL ID: ${myKlId}`);
    return myKlId.trim(); // Return the trimmed ID

  } catch (e) {
    // Log detailed error and re-throw a user-friendly error
    console.error(`Error in fSrvGetMyKlId for CS ID ${myCsId}: ${e.message}\nStack: ${e.stack}`);
    // Include specific error message if available, otherwise generic
    throw new Error(`Server error getting MyKL ID: ${e.message || e}`);
  }

} // END fSrvGetMyKlId




// ==========================================================================
// === Designer Access ===
// ==========================================================================




// fSrvValidateDesignerPassword /////////////////////////////////////////////////
// Purpose -> Validates a password attempt against the stored designer password.
// Inputs  -> passwordAttempt (String): The password entered by the user.
// Outputs -> (Boolean): True if the password matches, false otherwise.
function fSrvValidateDesignerPassword(passwordAttempt) {
    const funcName = "fSrvValidateDesignerPassword";
    const storedPassword = gSrv.designerPassword; // Assuming password is in gSrv

    if (!storedPassword) {
        console.error(`${funcName}: Designer password is not set in server globals (gSrv.designerPassword).`);
        Logger.log(`${funcName}: ERROR - Designer password not configured on server.`);
        // Return false if the server doesn't have a password configured
        return false;
    }

    // Basic check to ensure we have a string attempt
    if (typeof passwordAttempt !== 'string') {
        Logger.log(`${funcName}: Invalid password attempt type received: ${typeof passwordAttempt}`);
        return false;
    }

    // Perform case-sensitive comparison
    const isValid = (passwordAttempt === storedPassword);

    // Optional: Log the attempt result (avoid logging passwords themselves in production)
    Logger.log(`${funcName}: Password validation attempt. Result: ${isValid ? 'Success' : 'Failure'}`);

    return isValid;
} // END fSrvValidateDesignerPassword




// ==========================================================================
// === Low-Level Tag & Range Helpers ===
// ==========================================================================




// fSrvBuildTagMaps /////////////////////////////////////////////////////////
// Purpose -> Builds rowTag and colTag maps from a 2D array (Row 0 for colTags, Col 0 for rowTags).
// Inputs  -> fullData (Array[][]): The 2D array of data
// Outputs -> (Object): An object { rowTag: {}, colTag: {} } containing the maps.
function fSrvBuildTagMaps(fullData) {
    const rowTagMap = {};
    const colTagMap = {};

    // Validate data array structure
    if (!Array.isArray(fullData) || fullData.length === 0) {
        console.warn("fSrvBuildTagMaps: Input data array is empty or invalid.");
        return { rowTag: rowTagMap, colTag: colTagMap }; // Return empty maps
    }

    const numRows = fullData.length;
    const numCols = fullData[0]?.length || 0; // Handle potentially empty first row

    // Process Row Tags (from Column 0)
    for (let r = 0; r < numRows; r++) {
        const rowHeaderCell = fullData[r]?.[0]; // Value in the first column of this row
        if (typeof rowHeaderCell === 'string' && rowHeaderCell.trim()) {
            // Split potentially comma-separated tags
            rowHeaderCell.split(',').forEach(tag => {
                const trimmedTag = tag.trim();
                if (trimmedTag) {
                    // Map the trimmed tag to the current 0-based row index (r)
                    // Warn if tag is being overwritten (optional)
                    // if (rowTagMap.hasOwnProperty(trimmedTag)) {
                    //   console.warn(`fSrvBuildTagMaps: Overwriting row tag "${trimmedTag}" (Old: ${rowTagMap[trimmedTag]}, New: ${r})`);
                    // }
                    rowTagMap[trimmedTag] = r;
                }
            });
        }
    }

    // Process Column Tags (from Row 0)
    const colHeaderRow = fullData[0];
    if (Array.isArray(colHeaderRow)) {
        for (let c = 0; c < numCols; c++) {
            const colHeaderCell = colHeaderRow[c]; // Value in the first row of this column
            if (typeof colHeaderCell === 'string' && colHeaderCell.trim()) {
                // Split potentially comma-separated tags
                colHeaderCell.split(',').forEach(tag => {
                    const trimmedTag = tag.trim();
                    if (trimmedTag) {
                        // Map the trimmed tag to the current 0-based column index (c)
                        // Warn if tag is being overwritten (optional)
                        // if (colTagMap.hasOwnProperty(trimmedTag)) {
                        //   console.warn(`fSrvBuildTagMaps: Overwriting col tag "${trimmedTag}" (Old: ${colTagMap[trimmedTag]}, New: ${c})`);
                        // }
                        colTagMap[trimmedTag] = c;
                    }
                });
            }
        }
    } else {
        console.warn("fSrvBuildTagMaps: Header row (row 0) is missing or invalid. Cannot build column tags.");
    }

    return { rowTag: rowTagMap, colTag: colTagMap };
} // END fSrvBuildTagMaps




// fSrvResolveTag ///////////////////////////////////////////////////////////
// Purpose -> Resolves a single row or column tag/index using the provided tag maps.
// Inputs  -> tagOrIndex (String | Number): The tag string or 0-based index.
//         -> tagMap (Object): The corresponding tag map (rowTag or colTag).
//         -> type (String): 'row' or 'col' for logging purposes.
// Outputs -> (Number): The resolved 0-based index, or NaN if resolution fails.
function fSrvResolveTag(tagOrIndex, tagMap, type = 'unknown') {
    if (typeof tagOrIndex === 'number' && !isNaN(tagOrIndex) && tagOrIndex >= 0) {
        // Input is already a valid numeric index
        return tagOrIndex;
    } else if (typeof tagOrIndex === 'string' && tagOrIndex.trim()) {
        // Input is a string tag
        const trimmedTag = tagOrIndex.trim();
        if (tagMap.hasOwnProperty(trimmedTag)) {
            // Tag found in the map
            return tagMap[trimmedTag];
        } else {
            // Tag not found
            console.warn(`fSrvResolveTag: Could not resolve ${type} tag "${trimmedTag}"`);
            return NaN;
        }
    } else {
        // Input is invalid type or empty
        console.warn(`fSrvResolveTag: Invalid input provided for ${type} resolution:`, tagOrIndex);
        return NaN;
    }
} // END fSrvResolveTag




// ==========================================================================
// === Core Sheet Loader – “Game” Tab ===
// ==========================================================================




// fSrvReadCSGameSheet /////////////////////////////////////////////////////////////////////////////////
// Purpose -> Loads full data, format, and notes from the 'Game' sheet of a given spreadsheet ID.
// Inputs  -> gIndex.CSID (String): The ID of the spreadsheet to read from.
// Outputs -> (Object): Structured object { arr, format, notesArr } containing sheet data.
// Throws  -> (Error): If sheet ID is invalid, spreadsheet/sheet cannot be opened, or data extraction fails.
function fSrvReadCSGameSheet(gIndex) {
    const funcName = 'fSrvReadCSGameSheet'; // Added funcName for better logging/errors
    try {
        // Validate gIndex.CSID
        if (!gIndex.CSID || typeof gIndex.CSID !== 'string') {
            throw new Error("Invalid or missing Sheet ID provided.");
        }
        Logger.log(`${funcName}: Opening Spreadsheet ID: ${gIndex.CSID}`); // Log opening attempt

        // Open Spreadsheet
        const ss = SpreadsheetApp.openById(gIndex.CSID);
        if (!ss) { // Check if spreadsheet opened successfully
             throw new Error(`Could not open Spreadsheet with ID: ${gIndex.CSID}. Check permissions and ID validity.`);
        }

        // Get "Game" Sheet
        const sh = ss.getSheetByName("Game");
        if (!sh) { // Check if sheet was found
            throw new Error(`Sheet named "Game" not found in Spreadsheet ID: ${gIndex.CSID}.`);
        }
        Logger.log(`${funcName}: Successfully opened sheet "Game" in ID: ${gIndex.CSID}`);

        // Call fSrvExtractSheetData with the obtained sheet object
        Logger.log(`${funcName}: Extracting data from sheet "Game"...`);
        return fSrvExtractSheetData(sh);

    } catch (e) {
        // Determine context based on error message content for better reporting
        const context = e.message.includes("openById") || e.message.includes("Spreadsheet with ID") ? "opening spreadsheet"
                      : e.message.includes("getSheetByName") || e.message.includes("not found") ? "getting 'Game' sheet"
                      : e.stack?.includes("fSrvExtractSheetData") ? "processing sheet data" // Check stack for fSrvExtractSheetData context
                      : "during operation"; // Generic fallback context

        const msg = `${funcName}: Error ${context}: ${e.message}`;
        console.error(msg + "\nStack:\n" + e.stack); // Log detailed error
        throw new Error(msg); // Re-throw a consolidated error message
    }
} // END fSrvReadCSGameSheet




// fSrvExtractSheetData ////////////////////////////////////////////////////////////////////////////////
// Reads data, formats, and notes from the sheet and returns structured result.
function fSrvExtractSheetData(sh) {
    const rngData = sh.getDataRange();
    const arr = rngData.getValues();

    if (fSrvIsSheetTrulyEmpty(sh, arr)) {
        return fSrvBuildReturnObject(
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
    const format = fSrvBuildFormatObject(sh, rngData, numCols);
    const notesArr = rngData.getNotes();

    return fSrvBuildReturnObject(arr, format, notesArr);
} // END fSrvExtractSheetData




// fSrvIsSheetTrulyEmpty ////////////////////////////////////////////////////////////////////////////////
// Determines if the sheet has zero real content, returns true if it's blank.
function fSrvIsSheetTrulyEmpty(sh, arr) {
    const numRows = arr.length;
    const numCols = arr[0]?.length || 0;
    const onlyEmpty = numRows === 1 && numCols === 1 && arr[0][0] === "";
    return (numRows === 0 || numCols === 0 || onlyEmpty) &&
           (sh.getLastRow() === 0 && sh.getLastColumn() === 0);
} // END fSrvIsSheetTrulyEmpty




// fSrvBuildFormatObject ////////////////////////////////////////////////////////////////////////////////
// Constructs formatting object from raw arrays and merge info.
function fSrvBuildFormatObject(sh, rngData, numCols) {
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
} // END fSrvBuildFormatObject




// fSrvBuildReturnObject ////////////////////////////////////////////////////////////////////////////////
// Combines the final return structure.
function fSrvBuildReturnObject(arr, format, notesArr) {
    return {
        arr: arr,
        format: format,
        notesArr: notesArr
    };
} // END fSrvBuildReturnObject



// ==========================================================================
// === Generic Sheet Operations ===
// ==========================================================================




// fSrvLoadFullGoogleSheetAndTags /////////////////////////////////////////////////////
// Purpose -> Loads column tags (Row 0), row tags (Col 0), and all cell text values
//            from a specified sheet within a specified Google Workbook (DB, MasterCS, etc.).
//            Validates tag uniqueness (case-insensitive).
// Inputs  -> workbookAbr (String): Abbreviation ('db', 'mastercs', 'masterkl', 'mycs', 'mykl').
//         -> sheetName (String): The name of the specific sheet (tab) to read from.
//         -> csId (String): The Character Sheet ID (used for 'mycs', and to find 'mykl').
// Outputs -> (Object): { ColTags: Object, RowTags: Object, sheetText2D: Array[][] } on success.
// Throws  -> (Error): If inputs are invalid, workbook ID not found/derived, sheet not found,
//                     sheet is empty, or duplicate tags are found (case-insensitive).
function fSrvLoadFullGoogleSheetAndTags(workbookAbr, sheetName, csId) {
    const funcName = "fSrvLoadFullGoogleSheetAndTags";
    Logger.log(`${funcName}: Loading Tags & Data for Workbook: "${workbookAbr}", Sheet: "${sheetName}", CSID: "${csId}".`);

    // --- 1. Validate Inputs ---
    if (!workbookAbr || typeof workbookAbr !== 'string') {
        throw new Error(`${funcName}: Invalid or missing workbookAbr provided.`);
    }
    if (!sheetName || typeof sheetName !== 'string' || sheetName.trim() === '') {
        throw new Error(`${funcName}: Invalid or empty sheetName provided.`);
    }
    const trimmedSheetName = sheetName.trim();
    const lowerWorkbookAbr = workbookAbr.toLowerCase();

    // --- 2. Determine Workbook ID ---
    let workbookID = null;
    try {
        switch (lowerWorkbookAbr) {
            case 'db':       workbookID = gSrv.ids.sheets.db; break;
            case 'mastercs': workbookID = gSrv.ids.sheets.mastercs; break;
            case 'masterkl': workbookID = gSrv.ids.sheets.masterkl; break;
            case 'mycs':
                if (!csId) { throw new Error("CSID is required to identify 'MyCS' workbook."); }
                workbookID = csId; // MyCS ID is the CSID itself
                break;
            case 'mykl':
                if (!csId) { throw new Error("CSID is required to look up 'MyKL' workbook ID."); }
                workbookID = fSrvGetMyKlId(csId); // Get MyKL ID using helper
                if (!workbookID) { throw new Error(`Could not find linked 'MyKL' workbook ID for CSID: ${csId}`); }
                break;
            default:
                throw new Error(`Unsupported workbook abbreviation: "${workbookAbr}"`);
        }
        if (!workbookID) { throw new Error(`Workbook ID for "${workbookAbr}" resolved to null or empty.`); }
        Logger.log(`   -> Resolved Workbook ID for "${workbookAbr}": ${workbookID}`);
    } catch (e) {
         console.error(`${funcName} Error resolving Workbook ID: ${e.message}`);
         throw new Error(`Failed to resolve Workbook ID for "${workbookAbr}": ${e.message}`); // Re-throw
    }


    // --- 3. Open Target Workbook & Get Target Tab ---
    let ss, sh;
    try {
        ss = SpreadsheetApp.openById(workbookID);
        sh = ss.getSheetByName(trimmedSheetName);
        if (!sh) {
            throw new Error(`Sheet named "${trimmedSheetName}" not found in Workbook ID: ${workbookID} (Abr: ${workbookAbr}).`);
        }
        Logger.log(`   -> Successfully opened sheet "${trimmedSheetName}" in Workbook "${workbookAbr}".`);
    } catch (e) {
        console.error(`Error accessing sheet "${trimmedSheetName}" in Workbook "${workbookAbr}" (ID: ${workbookID}): ${e.message}`);
        throw new Error(`Failed to access sheet "${trimmedSheetName}" in Workbook "${workbookAbr}": ${e.message}`);
    }

    // --- 4. Get Full Data Range & Check if Empty ---
    const dataRange = sh.getDataRange();
    const sheetText2D = dataRange.getValues();
    const numRows = sheetText2D.length;
    const numCols = numRows > 0 ? (sheetText2D[0]?.length || 0) : 0;

    if (numRows === 0 || numCols === 0 || (numRows === 1 && numCols === 1 && sheetText2D[0][0] === '')) {
        throw new Error(`${funcName}: Sheet "${trimmedSheetName}" in Workbook "${workbookAbr}" exists but appears empty.`);
    }
    Logger.log(`   -> Read ${numRows}x${numCols} cells from "${trimmedSheetName}".`);


    // --- 5. Process Column Tags (Row 0) ---
    const colTagsMap = {};
    const colTagsSeen = new Set();
    const headerRow = sheetText2D[0];

    for (let c = 0; c < numCols; c++) {
        const cellValue = headerRow[c];
        if (typeof cellValue === 'string' && cellValue.trim() !== '') {
            const tags = cellValue.split(',').map(tag => tag.trim()).filter(Boolean);
            for (const tag of tags) {
                const lowerTag = tag.toLowerCase();
                if (colTagsSeen.has(lowerTag)) {
                    throw new Error(`${funcName}: Duplicate Column Tag found (case-insensitive): "${tag}" in Row 0, Column ${c + 1}. Sheet: "${trimmedSheetName}", Workbook: "${workbookAbr}".`);
                }
                if (colTagsMap.hasOwnProperty(tag)) {
                     Logger.log(`   -> WARNING: Column Tag "${tag}" reused in Row 0. Mapping to last occurrence (Col ${c}).`);
                }
                colTagsMap[tag] = c;
                colTagsSeen.add(lowerTag);
            }
        }
    }
    Logger.log(`   -> Processed ${Object.keys(colTagsMap).length} unique column tags.`);

    // --- 6. Process Row Tags (Col 0) ---
    const rowTagsMap = {};
    const rowTagsSeen = new Set();

    for (let r = 0; r < numRows; r++) {
        const cellValue = (sheetText2D[r] && sheetText2D[r].length > 0) ? sheetText2D[r][0] : undefined;
        if (typeof cellValue === 'string' && cellValue.trim() !== '') {
            const tags = cellValue.split(',').map(tag => tag.trim()).filter(Boolean);
            for (const tag of tags) {
                const lowerTag = tag.toLowerCase();
                if (rowTagsSeen.has(lowerTag)) {
                    throw new Error(`${funcName}: Duplicate Row Tag found (case-insensitive): "${tag}" in Column 0, Row ${r + 1}. Sheet: "${trimmedSheetName}", Workbook: "${workbookAbr}".`);
                }
                 if (rowTagsMap.hasOwnProperty(tag)) {
                     Logger.log(`   -> WARNING: Row Tag "${tag}" reused in Col 0. Mapping to last occurrence (Row ${r}).`);
                }
                rowTagsMap[tag] = r;
                rowTagsSeen.add(lowerTag);
            }
        }
    }
    Logger.log(`   -> Processed ${Object.keys(rowTagsMap).length} unique row tags.`);


    // --- 7. Return Result ---
    Logger.log(`${funcName}: Successfully loaded tags and data for Workbook "${workbookAbr}", Sheet "${trimmedSheetName}".`);
    return {
        ColTags: colTagsMap,
        RowTags: rowTagsMap,
        sheetText2D: sheetText2D
    };

} // END fSrvLoadFullGoogleSheetAndTags




// fSrvGetSheetRangeDataNTags /////////////////////////////////////////////////////////
// Purpose -> Reads data from a specified sheet, accepting a sheet key OR fileId.
//            Optionally accepts a rangeObject to extract a slice; otherwise returns full sheet.
//            Reads the full sheet once for tag mapping. Handles 'Calc_LastRow'.
//            Builds relative tags for slices, absolute tags for full sheet returns.
// Inputs  -> sheetKeyOrId (String): Key from gSrv.ids.sheets OR a direct Sheet fileId.
//         -> sheetName (String): The name of the specific sheet (tab) to read from.
//         -> rangeObject (Object | null | undefined): Optional. {r1, c1, r2, c2} using tags,
//                                                      0-based indices, or 'Calc_LastRow' for r2.
//                                                      If omitted/null/undefined, returns full sheet data.
// Outputs -> (Object): { data: Array[][]|Array[]|Any|null, colTags: Object, rowTags: Object }
//            'data' contains values, formatted correctly (single, 1D, or 2D).
//            colTags/rowTags contain mappings relative to 'data' (absolute if full sheet returned).
//            Returns null for 'data' on critical errors (though usually throws).
// Throws  -> (Error): If inputs are invalid (except optional rangeObject), sheet/range access fails, etc.
function fSrvGetSheetRangeDataNTags(sheetKeyOrId, sheetName, rangeObject = null) {
    // Log received parameters, handling optional rangeObject
    const rangeLog = rangeObject ? JSON.stringify(rangeObject) : 'Not Provided (Full Sheet)';
    Logger.log(`fSrvGetSheetRangeDataNTags: Request received. Key/ID: "${sheetKeyOrId}", Sheet: "${sheetName}", Range: ${rangeLog}`);

    let fileId = null;
    let identifiedBy = '';
    let absoluteRowTagMap = {};
    let absoluteColTagMap = {};
    let relativeRowTagMap = {}; // Only used if rangeObject is provided
    let relativeColTagMap = {}; // Only used if rangeObject is provided

    // --- 1. Resolve File ID ---
    if (!sheetKeyOrId || typeof sheetKeyOrId !== 'string') {
        const errorMsg = `Invalid or missing sheetKeyOrId parameter.`;
        console.error(`fSrvGetSheetRangeDataNTags Error: ${errorMsg}`);
        throw new Error(`getServerSheetData: ${errorMsg}`);
    }
    const isLikelySheetId = sheetKeyOrId.length > 30 && /^[a-z0-9_-]+$/i.test(sheetKeyOrId);
    if (isLikelySheetId) {
        fileId = sheetKeyOrId;
        identifiedBy = `Direct ID: ${fileId}`;
        Logger.log(`   -> Interpreted first argument as Direct File ID: ${fileId}`);
    } else {
        const keyLower = sheetKeyOrId.toLowerCase();
        if (!gSrv.ids.sheets[keyLower]) {
            const errorMsg = `Invalid sheetKey: "${sheetKeyOrId}". Not found in gSrv.ids.sheets.`;
            console.error(`fSrvGetSheetRangeDataNTags Error: ${errorMsg}`);
            throw new Error(`getServerSheetData: ${errorMsg}`);
        }
        fileId = gSrv.ids.sheets[keyLower];
        identifiedBy = `Key: ${sheetKeyOrId} -> ID: ${fileId}`;
        Logger.log(`   -> Interpreted first argument as Key: "${sheetKeyOrId}", resolved to ID: ${fileId}`);
    }

    // --- 2. Validate Required Inputs (Sheet ID, Sheet Name) ---
     if (!fileId) { throw new Error(`getServerSheetData: Could not determine File ID.`); }
    if (!sheetName || typeof sheetName !== 'string') { throw new Error(`getServerSheetData: Invalid or missing sheetName.`); }

    // --- Validate Optional rangeObject Structure (only if provided) ---
    let isRangeProvidedAndValid = false;
    if (rangeObject && typeof rangeObject === 'object') {
        if (rangeObject.r1 !== undefined && rangeObject.c1 !== undefined && rangeObject.r2 !== undefined && rangeObject.c2 !== undefined) {
            isRangeProvidedAndValid = true;
            Logger.log(`   -> Valid rangeObject structure provided.`);
        } else {
             // Provided object is incomplete - treat as invalid range, will default to full sheet
             Logger.log(`   -> rangeObject provided but incomplete: ${JSON.stringify(rangeObject)}. Defaulting to full sheet.`);
             rangeObject = null; // Nullify to trigger full sheet logic
        }
    } else if (rangeObject) {
         // Provided but not an object - treat as invalid range
         Logger.log(`   -> rangeObject provided but not an object: ${typeof rangeObject}. Defaulting to full sheet.`);
         rangeObject = null; // Nullify to trigger full sheet logic
    } else {
        // rangeObject is null or undefined - default to full sheet (intended)
        Logger.log(`   -> rangeObject not provided. Defaulting to full sheet.`);
    }

    try {
        // --- 3. Open Sheet & Read Full Data ---
        const ss = SpreadsheetApp.openById(fileId);
        const sh = ss.getSheetByName(sheetName);
        if (!sh) { throw new Error(`Sheet named "${sheetName}" not found in Sheet (${identifiedBy}).`); }
        const fullData = sh.getDataRange().getValues();
        const numRows = fullData.length;
        const numCols = fullData[0]?.length || 0;
        Logger.log(`fSrvGetSheetRangeDataNTags: Read ${numRows}x${numCols} cells from "${sheetName}".`);

        // --- 4. Build *Absolute* Tag Maps ---
        const absoluteTagMaps = fSrvBuildTagMaps(fullData);
        absoluteRowTagMap = absoluteTagMaps.rowTag;
        absoluteColTagMap = absoluteTagMaps.colTag;
        // Logger.log(`fSrvGetSheetRangeDataNTags: Built absolute tag maps. Rows: ${Object.keys(absoluteRowTagMap).length}, Cols: ${Object.keys(absoluteColTagMap).length}`);

        // --- Handle Empty Sheet Case ---
        if (numRows === 0 || numCols === 0) {
            console.warn(`fSrvGetSheetRangeDataNTags: Sheet "${sheetName}" (${identifiedBy}) appears empty.`);
            return { data: [[]], colTags: {}, rowTags: {} }; // Return empty structure
        }

        // === Branch Logic: Full Sheet vs. Range Slice ===

        if (isRangeProvidedAndValid) {
            // --- BRANCH A: Process Specific Range ---
            Logger.log(`   -> Processing provided range...`);

            // --- 5. Resolve Input Range Object Tags to *Absolute* Indices ---
            const r1_abs = fSrvResolveTag(rangeObject.r1, absoluteRowTagMap, 'row');
            const c1_abs = fSrvResolveTag(rangeObject.c1, absoluteColTagMap, 'col');
            const c2_abs = fSrvResolveTag(rangeObject.c2, absoluteColTagMap, 'col');
            let r2_abs; // Declare r2_abs here

            // --- Handle 'Calc_LastRow' for r2 ---
            if (typeof rangeObject.r2 === 'string' && rangeObject.r2.toLowerCase() === 'calc_lastrow') {
                 if (isNaN(c1_abs)) { throw new Error(`Cannot calculate last row: Column 'c1' (${rangeObject.c1}) could not be resolved.`); }
                 Logger.log(`   -> Calculating last row for column index ${c1_abs}...`);
                 const lastSheetRow = sh.getLastRow(); // 1-based index
                 r2_abs = -1; // Initialize as not found
                 for (let r = lastSheetRow - 1; r >= 0; r--) {
                     const cellValue = fullData[r]?.[c1_abs];
                     if (cellValue !== undefined && cellValue !== null && String(cellValue).trim() !== '') {
                         r2_abs = r; // Found the last non-empty row (0-based index)
                         Logger.log(`   -> Found last non-empty cell at row index ${r2_abs}.`);
                         break;
                     }
                 }
                 if (r2_abs === -1) {
                      if(!isNaN(r1_abs)) {
                          r2_abs = r1_abs; // Fallback to r1 if calc fails but r1 is valid
                          Logger.log(`   -> Warning: Could not find last non-empty row in column ${c1_abs}. Using r1 index ${r1_abs} as fallback.`);
                      } else {
                           throw new Error(`Cannot calculate last row: Column ${c1_abs} appears empty and r1 ('${rangeObject.r1}') is also invalid.`);
                      }
                 }
            } else {
                 // Standard tag/index resolution for r2
                 r2_abs = fSrvResolveTag(rangeObject.r2, absoluteRowTagMap, 'row');
            }
            // --- End 'Calc_LastRow' Handling ---

            // --- Validate All Resolved Indices ---
            if ([r1_abs, c1_abs, r2_abs, c2_abs].some(isNaN)) {
                 let failedTags = [];
                 if (isNaN(r1_abs)) failedTags.push(`r1: ${rangeObject.r1}`);
                 if (isNaN(c1_abs)) failedTags.push(`c1: ${rangeObject.c1}`);
                 if (isNaN(r2_abs)) failedTags.push(`r2: ${rangeObject.r2} (resolved: ${r2_abs})`);
                 if (isNaN(c2_abs)) failedTags.push(`c2: ${rangeObject.c2}`);
                 throw new Error(`Could not resolve absolute tags/indices in rangeObject: ${JSON.stringify(rangeObject)}. Failed Tags: ${failedTags.join(', ')}`);
            }
            Logger.log(`fSrvGetSheetRangeDataNTags: Resolved range to absolute indices: r1=${r1_abs}, c1=${c1_abs}, r2=${r2_abs}, c2=${c2_abs}`);

            // --- 6. Define Extraction Boundaries ---
            const rStart = Math.min(r1_abs, r2_abs);
            const rEnd = Math.max(r1_abs, r2_abs);
            const cStart = Math.min(c1_abs, c2_abs);
            const cEnd = Math.max(c1_abs, c2_abs);

            // --- 7. Extract Data Slice ---
            if (rStart >= numRows || cStart >= numCols) {
                  console.warn(`fSrvGetSheetRangeDataNTags: Resolved range start [${rStart}, ${cStart}] is outside the bounds of the sheet data [${numRows}, ${numCols}]. Returning empty data.`);
                  return { data: [[]], colTags: {}, rowTags: {} };
            }
            const extractedData = fullData
                 .slice(rStart, rEnd + 1)
                 .map(row => row.slice(cStart, cEnd + 1));
            const extractedRows = extractedData.length;
            const extractedCols = extractedData[0]?.length || 0;

            // --- 8. Build *Relative* Tag Maps ---
            // Adjust Column Tags
            for (const tag in absoluteColTagMap) {
                 const absoluteIndex = absoluteColTagMap[tag];
                 if (absoluteIndex >= cStart && absoluteIndex <= cEnd) {
                     const relativeIndex = absoluteIndex - cStart;
                     relativeColTagMap[tag] = relativeIndex;
                 }
            }
            // Adjust Row Tags
            for (const tag in absoluteRowTagMap) {
                 const absoluteIndex = absoluteRowTagMap[tag];
                 if (absoluteIndex >= rStart && absoluteIndex <= rEnd) {
                     const relativeIndex = absoluteIndex - rStart;
                     relativeRowTagMap[tag] = relativeIndex;
                 }
            }
            Logger.log(`fSrvGetSheetRangeDataNTags: Built relative tag maps for slice. Rel Rows: ${Object.keys(relativeRowTagMap).length}, Rel Cols: ${Object.keys(relativeColTagMap).length}`);

            // --- 9. Format Return Data (for Slice) ---
            let returnData;
            if (extractedRows === 1 && extractedCols === 1) {
                 Logger.log(`fSrvGetSheetRangeDataNTags: Formatting sliced data as single value.`);
                 returnData = extractedData[0][0];
            } else if (extractedRows === 1) {
                 Logger.log(`fSrvGetSheetRangeDataNTags: Formatting sliced data as 1D array (single row).`);
                 returnData = extractedData[0];
            } else if (extractedCols === 1 && extractedData.every(row => row.length === 1)) {
                  Logger.log(`fSrvGetSheetRangeDataNTags: Formatting sliced data as 1D array (single column).`);
                  returnData = extractedData.map(row => row[0]);
            } else {
                 Logger.log(`fSrvGetSheetRangeDataNTags: Formatting sliced data as 2D array (${extractedRows}x${extractedCols}).`);
                 returnData = extractedData;
            }

            // --- 10. Return Final Object (for Slice) ---
            return { data: returnData, colTags: relativeColTagMap, rowTags: relativeRowTagMap };

        } else {
            // --- BRANCH B: Return Full Sheet ---
            Logger.log(`   -> No range provided or range invalid. Returning full sheet data...`);

            // --- 9. Format Return Data (for Full Sheet) ---
            let returnData;
            if (numRows === 1 && numCols === 1) {
                Logger.log(`fSrvGetSheetRangeDataNTags: Formatting full sheet data as single value.`);
                returnData = fullData[0][0];
            } else if (numRows === 1) {
                Logger.log(`fSrvGetSheetRangeDataNTags: Formatting full sheet data as 1D array (single row).`);
                returnData = fullData[0];
            } else if (numCols === 1 && fullData.every(row => row.length === 1)) {
                 Logger.log(`fSrvGetSheetRangeDataNTags: Formatting full sheet data as 1D array (single column).`);
                 returnData = fullData.map(row => row[0]);
            } else {
                Logger.log(`fSrvGetSheetRangeDataNTags: Formatting full sheet data as 2D array (${numRows}x${numCols}).`);
                returnData = fullData;
            }

            // --- 10. Return Final Object (for Full Sheet) ---
            // Use the absolute tag maps when returning the full sheet
            return { data: returnData, colTags: absoluteColTagMap, rowTags: absoluteRowTagMap };
        }

    } catch (e) {
        console.error(`Error in fSrvGetSheetRangeDataNTags for Input "${sheetKeyOrId}", Sheet "${sheetName}", Range ${rangeLog}: ${e.message}\nStack: ${e.stack}`);
        throw new Error(`Server error processing sheet data: ${e.message || e}`);
    }

} // END fSrvGetSheetRangeDataNTags




// fSrvSaveDataToSheetRange /////////////////////////////////////////////////////////
// Purpose -> Writes data to a specified range in a given Google Sheet,
//            accepting a sheet key, sheet name, a rangeObject with tags/indices,
//            and a value (single or 1D/2D array).
// Inputs  -> sheetKey (String): Key corresponding to a Sheet ID in gSrv.ids.sheets.
//         -> sheetName (String): The name of the specific sheet (tab) to write to.
//         -> rangeObject (Object): {r1, c1, r2, c2} using tags or 0-based indices.
//         -> valueOrValues (Any | Array[] | Array[][]): The single value or array to write.
// Outputs -> (Boolean): True on success.
// Throws  -> (Error): If inputs are invalid, sheet/range access fails, or
//                     value dimensions don't match range dimensions (for arrays).
function fSrvSaveDataToSheetRange(sheetKey, sheetName, rangeObject, valueOrValues) {
    Logger.log(`fSrvSaveDataToSheetRange: Requesting write to key "${sheetKey}", sheet "${sheetName}", range: ${JSON.stringify(rangeObject)}, type: ${Array.isArray(valueOrValues) ? (Array.isArray(valueOrValues[0]) ? '2D Array' : '1D Array') : typeof valueOrValues}`);

    // === 1. Validate Inputs (Sheet Key, Name, Range Object) ===
    if (!sheetKey || typeof sheetKey !== 'string' || !gSrv.ids.sheets[sheetKey.toLowerCase()]) {
        const errorMsg = `Invalid or missing sheetKey: "${sheetKey}". Not found in gSrv.ids.sheets.`;
        console.error(`fSrvSaveDataToSheetRange Error: ${errorMsg}`);
        throw new Error(`setServerSheetData: ${errorMsg}`);
    }
    const fileId = gSrv.ids.sheets[sheetKey.toLowerCase()];
    if (!sheetName || typeof sheetName !== 'string') {
        const errorMsg = `Invalid or missing sheetName: "${sheetName}".`;
        console.error(`fSrvSaveDataToSheetRange Error: ${errorMsg}`);
        throw new Error(`setServerSheetData: ${errorMsg}`);
    }
    if (!rangeObject || typeof rangeObject !== 'object' ||
        rangeObject.r1 === undefined || rangeObject.c1 === undefined ||
        rangeObject.r2 === undefined || rangeObject.c2 === undefined) {
        const errorMsg = `Invalid or incomplete rangeObject: ${JSON.stringify(rangeObject)}.`;
        console.error(`fSrvSaveDataToSheetRange Error: ${errorMsg}`);
        throw new Error(`setServerSheetData: ${errorMsg}`);
    }

    try {
        // === 2. Open Sheet & Read Full Data (for Tag Mapping) ===
        const ss = SpreadsheetApp.openById(fileId);
        const sh = ss.getSheetByName(sheetName);
        if (!sh) {
            throw new Error(`Sheet named "${sheetName}" not found in Sheet ID: ${fileId} (Key: ${sheetKey}).`);
        }
        const fullData = sh.getDataRange().getValues();
        if (fullData.length === 0 || fullData[0]?.length === 0) {
           console.warn(`fSrvSaveDataToSheetRange: Sheet "${sheetName}" appears empty. Tag resolution might fail if tags are used.`);
        }

        // === 3. Build Tag Maps (from in-memory data) ===
        const { rowTag, colTag } = fSrvBuildTagMaps(fullData);

        // === 4. Resolve Input Range Object Tags ===
        const r1 = fSrvResolveTag(rangeObject.r1, rowTag, 'row');
        const c1 = fSrvResolveTag(rangeObject.c1, colTag, 'col');
        const r2 = fSrvResolveTag(rangeObject.r2, rowTag, 'row');
        const c2 = fSrvResolveTag(rangeObject.c2, colTag, 'col');
        if ([r1, c1, r2, c2].some(isNaN)) {
            throw new Error(`Could not resolve one or more tags/indices in rangeObject: ${JSON.stringify(rangeObject)}`);
        }
        Logger.log(`fSrvSaveDataToSheetRange: Resolved range to indices: r1=${r1}, c1=${c1}, r2=${r2}, c2=${c2}`);

        // === 5. Get Target Range & Dimensions ===
        const rangeNumRows = Math.abs(r1 - r2) + 1;
        const rangeNumCols = Math.abs(c1 - c2) + 1;
        const a1Notation = fSrvConvertIndicesToA1(r1, c1, r2, c2);
        if (!a1Notation) {
             throw new Error(`Could not convert resolved indices [${r1},${c1},${r2},${c2}] to A1 notation.`);
        }
        const targetRange = sh.getRange(a1Notation);
        if (!targetRange) {
             throw new Error(`Could not get target range "${a1Notation}" in sheet "${sheetName}".`);
        }
        Logger.log(`fSrvSaveDataToSheetRange: Target A1 notation: ${a1Notation} (${rangeNumRows}x${rangeNumCols})`);

        // === 6. Write Value(s) based on Type and Range Size ===
        if (rangeNumRows === 1 && rangeNumCols === 1 && !Array.isArray(valueOrValues)) {
            // --- Single Cell, Single Value ---
            Logger.log(`   -> Using setValue for single cell with value: ${String(valueOrValues).substring(0,100)}...`);
            targetRange.setValue(valueOrValues);
        } else {
            // --- Range or Array Input ---
            let finalValuesArray; // This must be a 2D array for setValues

            if (!Array.isArray(valueOrValues)) {
                // --- Fill Range with Single Value ---
                Logger.log(`   -> Filling ${rangeNumRows}x${rangeNumCols} range with single value: ${valueOrValues}`);
                finalValuesArray = Array(rangeNumRows).fill(null).map(() => Array(rangeNumCols).fill(valueOrValues));
            } else if (Array.isArray(valueOrValues) && !Array.isArray(valueOrValues[0])) {
                // --- Handle 1D Array Input ---
                Logger.log(`   -> Handling 1D array input (length ${valueOrValues.length}) for ${rangeNumRows}x${rangeNumCols} range.`);
                if (rangeNumRows === 1 && rangeNumCols === valueOrValues.length) {
                    // Match: Single row range, array length matches column count
                    finalValuesArray = [valueOrValues]; // Wrap 1D array in another array
                } else if (rangeNumCols === 1 && rangeNumRows === valueOrValues.length) {
                    // Match: Single column range, array length matches row count
                    finalValuesArray = valueOrValues.map(v => [v]); // Convert each value to a single-element array
                } else {
                    // Mismatch
                    throw new Error(`Dimension mismatch: Target range is ${rangeNumRows}x${rangeNumCols}, but provided 1D array has length ${valueOrValues.length}.`);
                }
            } else {
                // --- Handle 2D Array Input ---
                Logger.log(`   -> Handling 2D array input (${valueOrValues.length}x${valueOrValues[0]?.length || 0}) for ${rangeNumRows}x${rangeNumCols} range.`);
                const valuesNumRows = valueOrValues.length;
                const valuesNumCols = valueOrValues[0]?.length || 0;
                // Check for consistent column lengths
                if (!valueOrValues.every(row => Array.isArray(row) && row.length === valuesNumCols)) {
                    throw new Error(`Input 2D array has inconsistent column lengths.`);
                }
                // Validate dimensions
                if (rangeNumRows !== valuesNumRows || rangeNumCols !== valuesNumCols) {
                    throw new Error(`Dimension mismatch: Resolved range is ${rangeNumRows}x${rangeNumCols}, but provided 2D array is ${valuesNumRows}x${valuesNumCols}.`);
                }
                finalValuesArray = valueOrValues; // Use the validated 2D array directly
            }

            // Use setValues for arrays or fill operations
            Logger.log(`   -> Using setValues for ${finalValuesArray.length}x${finalValuesArray[0]?.length || 0} array.`);
            targetRange.setValues(finalValuesArray);
        }

        Logger.log(`fSrvSaveDataToSheetRange: Successfully wrote data to range ${a1Notation}.`);
        return true; // Indicate success

    } catch (e) {
        // Log detailed error and re-throw a user-friendly error
        console.error(`Error in fSrvSaveDataToSheetRange for Key ${sheetKey}, Sheet "${sheetName}", Range ${JSON.stringify(rangeObject)}: ${e.message}\nStack: ${e.stack}`);
        throw new Error(`Server error writing sheet data: ${e.message || e}`);
    }

} // END fSrvSaveDataToSheetRange




// ==========================================================================
// === Player / GM Screen Logic ===
// ==========================================================================




// fSrvGetURLToPlayerCharForGMScreen /////////////////////////////////////////////////////
// Purpose -> Reads specific static data (Race/Class, Level, Player/Char Names, Slot)
//            from the player's Character Sheet ('RaceClass' tab).
// Inputs  -> csId (String): The ID of the player's Character Sheet.
// Outputs -> (Object): { raceClass, level, playerName, charName, slotNum } on success.
//            'slotNum' will be the cleaned tag (e.g., "Slot3") or null if validation fails.
// Throws  -> (Error): If sheet/tab/tags not found or critical values are missing.
function fSrvGetURLToPlayerCharForGMScreen(csId) {
    const funcName = "fSrvGetURLToPlayerCharForGMScreen";
    // Logger.log(`${funcName}: Fetching static header data from CS ID: ${csId}`);

    try {
        // --- 1. Validate CS ID ---
        if (!csId) {
            throw new Error("Character Sheet ID argument was not provided or is empty.");
        }

        // --- 2. Open Sheet & Tab ---
        const ss = SpreadsheetApp.openById(csId);
        const raceClassSheet = ss.getSheetByName('RaceClass');
        if (!raceClassSheet) {
            throw new Error(`Sheet named 'RaceClass' not found in CS ID: ${csId}.`);
        }

        // --- 3. Read Full Data & Build Tags ---
        const fullData = raceClassSheet.getDataRange().getValues();
        if (fullData.length === 0 || fullData[0]?.length === 0) {
            throw new Error(`Sheet 'RaceClass' in CS ID: ${csId} appears empty.`);
        }
        const { rowTag, colTag } = fSrvBuildTagMaps(fullData);

        // --- 4. Resolve Target Cell Indices ---
        const colValIndex = fSrvResolveTag('Val', colTag, 'col');
        const rowRaceClassIndex = fSrvResolveTag('RC', rowTag, 'row');
        const rowLevelIndex = fSrvResolveTag('level', rowTag, 'row');
        const rowPlayerNameIndex = fSrvResolveTag('playerName', rowTag, 'row');
        const rowCharNameIndex = fSrvResolveTag('charName', rowTag, 'row');
        const rowSlotIndex = fSrvResolveTag('slot', rowTag, 'row');

        // Validate all required tags resolved (including 'slot')
        if ([colValIndex, rowRaceClassIndex, rowLevelIndex, rowPlayerNameIndex, rowCharNameIndex, rowSlotIndex].some(isNaN)) {
            let missing = [];
            if (isNaN(colValIndex)) missing.push("Column 'Val'");
            if (isNaN(rowRaceClassIndex)) missing.push("Row 'RC'");
            if (isNaN(rowLevelIndex)) missing.push("Row 'level'");
            if (isNaN(rowPlayerNameIndex)) missing.push("Row 'playerName'");
            if (isNaN(rowCharNameIndex)) missing.push("Row 'charName'");
            if (isNaN(rowSlotIndex)) missing.push("Row 'slot'"); 
            throw new Error(`Could not resolve required tags in 'RaceClass' sheet: ${missing.join(', ')}.`);
        }

        // --- 5. Extract Values ---
        const raceClass = fullData[rowRaceClassIndex]?.[colValIndex] ?? '';
        const level = fullData[rowLevelIndex]?.[colValIndex] ?? '';
        const playerNameRaw = fullData[rowPlayerNameIndex]?.[colValIndex] ?? '';
        const charNameRaw = fullData[rowCharNameIndex]?.[colValIndex] ?? '';
        const slotRaw = fullData[rowSlotIndex]?.[colValIndex] ?? ''; 

        // --- 6. Process Player/Char Names (Get first word) ---
        const playerName = String(playerNameRaw).trim().split(' ')[0] || '';
        const charName = String(charNameRaw).trim().split(' ')[0] || '';

        // --- 7. Validate and Clean Slot Number --- <<< ADDED SECTION
        let cleanedSlotNum = null; // Default to null
        const slotString = String(slotRaw).trim();
        const slotRegex = /^Slot\s+([1-9])$/i; // Matches "Slot #[1-9]" case-insensitive
        const match = slotString.match(slotRegex);

        if (match && match[1]) {
            // Valid format found, clean it (remove space)
            cleanedSlotNum = `Slot${match[1]}`;
            Logger.log(`${funcName}: Valid Slot found: "${slotString}", Cleaned: "${cleanedSlotNum}"`);
        } else {
            // Invalid format or empty
            Logger.log(`${funcName}: Invalid Slot format found: "${slotString}". Setting slotNum to null.`);
            // cleanedSlotNum remains null
        }
        // --- END Slot Number Processing --- <<< END ADDED SECTION

        // --- 8. Log and Return ---
        const result = {
            raceClass: String(raceClass),
            level: String(level),
            playerName: playerName,
            charName: charName,
            slotNum: cleanedSlotNum // Include cleaned slot number (or null)
        };
        Logger.log(`${funcName}: Success. Returning: ${JSON.stringify(result)}`);
        return result;

    } catch (e) {
        // Log detailed error and re-throw
        console.error(`Error in ${funcName} for CS ID ${csId}: ${e.message}\nStack: ${e.stack}`);
        // Ensure specific error message is propagated
        throw new Error(`Server error getting character header data: ${e.message || e}`);
    }

} // END fSrvGetURLToPlayerCharForGMScreen




// fSrvSaveURLtoNamesAndLogToDBandPS ////////////////////////////////////////////////////////
// Purpose -> Writes bundled log and header data to target sheets (DB/GMScreen
//            and PS/PartyLog) using relative offsets from a base cell identified by
//            'Log' row tag and a dynamic slot column tag (e.g., 'Slot3') passed in the bundle.
// Inputs  -> dataBundle (Object): { log, vit, nish, url, raceClass, level, playerChar, slotNum }
// Outputs -> (Boolean): True if both writes succeeded, false otherwise.
// Throws  -> (Error): If critical errors occur (e.g., opening sheets, invalid bundle).
function fSrvSaveURLtoNamesAndLogToDBandPS(dataBundle) {
    const funcName = "fSrvSaveURLtoNamesAndLogToDBandPS";
    Logger.log(`${funcName}: Received data bundle. Preparing to write to DB and PS.`);

    // --- 1. Validate Input Bundle ---
    const requiredKeys = ['log', 'vit', 'nish', 'url', 'raceClass', 'level', 'playerChar', 'slotNum']; // Added slotNum
    if (!dataBundle || typeof dataBundle !== 'object' || !requiredKeys.every(key => dataBundle.hasOwnProperty(key))) {
        const missing = requiredKeys.filter(key => !dataBundle?.hasOwnProperty(key));
        const errorMsg = `Invalid or incomplete dataBundle received. Missing keys: ${missing.join(', ')}`;
        console.error(`${funcName} Error: ${errorMsg}`);
        throw new Error(`${funcName}: ${errorMsg}`);
    }

    // === 1a. Validate Slot Number <<< NEW SECTION ===
    const slotNumTag = dataBundle.slotNum;
    if (!slotNumTag || typeof slotNumTag !== 'string' || !slotNumTag.startsWith('Slot')) {
         const errorMsg = `Invalid slotNum ("${slotNumTag}") received in dataBundle. Must be a valid Slot tag (e.g., 'Slot3').`;
         console.error(`${funcName} Error: ${errorMsg}`);
         // Return false here instead of throwing, as the client call might succeed otherwise
         // Let the client handle the 'false' return value.
         return false;
    }
    Logger.log(`${funcName}: Using Slot Tag: ${slotNumTag}`);
    // === END NEW SECTION ===

    // --- 2. Define Targets and Base Row ---
    const targets = [
        { key: 'db', sheetName: 'GMScreen' },
        { key: 'ps', sheetName: 'PartyLog' }
    ];
    const baseCellTagR = 'Log'; // Keep base row tag fixed
    const baseCellTagC = slotNumTag; // Use dynamic slot tag
    const numHeaderRows = 6;

    // --- 3. Prepare Data Array (in correct vertical order) ---
    const dataToWrite = [
        [dataBundle.url],
        [dataBundle.raceClass],
        [dataBundle.level],
        [dataBundle.vit],
        [dataBundle.nish],
        [dataBundle.playerChar],
        [dataBundle.log]
    ];

    let overallSuccess = true;

    // --- 4. Loop Through Targets and Write Data ---
    for (const target of targets) {
        Logger.log(`${funcName}: Processing target: Key='${target.key}', Sheet='${target.sheetName}'`);
        let success = false;
        try {
            const fileId = gSrv.ids.sheets[target.key.toLowerCase()];
            if (!fileId) { throw new Error(`Could not find File ID for key '${target.key}'`); }

            const ss = SpreadsheetApp.openById(fileId);
            const sh = ss.getSheetByName(target.sheetName);
            if (!sh) { throw new Error(`Sheet named "${target.sheetName}" not found in Sheet ID: ${fileId} (Key: ${target.key}).`); }

            // Read full data *for this sheet* to build tag maps
            const fullData = sh.getDataRange().getValues();
            if (fullData.length === 0 || fullData[0]?.length === 0) {
                console.warn(`${funcName}: Target sheet "${target.sheetName}" appears empty. Cannot resolve base cell.`);
                throw new Error(`Target sheet "${target.sheetName}" is empty.`);
            }
            const { rowTag, colTag } = fSrvBuildTagMaps(fullData);

            // Resolve the *base cell* using fixed Row ('Log') and dynamic Column (baseCellTagC)
            const baseRowIndex = fSrvResolveTag(baseCellTagR, rowTag, 'row');
            const baseColIndex = fSrvResolveTag(baseCellTagC, colTag, 'col'); // Use dynamic tag

            if (isNaN(baseRowIndex) || isNaN(baseColIndex)) {
                throw new Error(`Could not resolve base cell tags ('${baseCellTagR}', '${baseCellTagC}') in sheet "${target.sheetName}".`);
            }
            Logger.log(`   -> Resolved Base Cell ('${baseCellTagR}', '${baseCellTagC}') to [${baseRowIndex}, ${baseColIndex}] in "${target.sheetName}".`);

            // Calculate the top-left cell of the 7-row range
            const startRowIndex = baseRowIndex - numHeaderRows;
            const startColIndex = baseColIndex; // Only writing to one column
            const numRowsToWrite = dataToWrite.length; // Should be 7
            const numColsToWrite = 1;

            // Validate start row index
            if (startRowIndex < 0) {
                throw new Error(`Calculated start row index (${startRowIndex}) is invalid (must be >= 0). Base cell ('${baseCellTagR}') might be too high.`);
            }

            // Get the target range using row/column indices (1-based for getRange)
            const targetRange = sh.getRange(startRowIndex + 1, startColIndex + 1, numRowsToWrite, numColsToWrite);
            const targetA1 = targetRange.getA1Notation();
            Logger.log(`   -> Target range calculated: ${targetA1} (${numRowsToWrite}x${numColsToWrite})`);

            // Write the prepared 2D array
            targetRange.setValues(dataToWrite);
            Logger.log(`   -> Successfully wrote data to ${targetA1} in sheet "${target.sheetName}".`);
            success = true;

        } catch (e) {
            // Log error for this specific target but continue to the next target
            console.error(`Error writing to target ${target.key}/${target.sheetName}: ${e.message}\nStack: ${e.stack}`);
            overallSuccess = false; // Mark that at least one target failed
        }
    } // End loop through targets

    // --- 5. Return Overall Success Status ---
    Logger.log(`${funcName}: Finished processing all targets. Overall Success: ${overallSuccess}`);
    if (!overallSuccess) {
       Logger.log(`${funcName}: Write failed for at least one target.`);
    }

    return overallSuccess; // Return true only if BOTH writes succeeded

} // END fSrvSaveURLtoNamesAndLogToDBandPS




// ==========================================================================
// === Firestore Integration ===
// ==========================================================================




// fSrvGetFirestoreInstance ///////////////////////////////////////////////////////
// Purpose -> Initializes and returns an authenticated Firestore instance using
//            credentials stored in PropertiesService. [Processes Key String]
// Inputs  -> None.
// Outputs -> (Object | null): Authenticated Firestore instance or null on error.
function fSrvGetFirestoreInstance() {
    const funcName = "fSrvGetFirestoreInstance";
    Logger.log(`${funcName}: Attempting to initialize Firestore...`); // <<< KEPT: Entry point log
    let clientEmail, privateKeyRaw, projectId, processedKey; // Declare vars
    try {
        // Retrieve credentials from Script Properties
        const scriptProperties = PropertiesService.getScriptProperties();
        clientEmail = scriptProperties.getProperty('FIRESTORE_CLIENT_EMAIL');
        privateKeyRaw = scriptProperties.getProperty('FIRESTORE_PRIVATE_KEY'); // Get the raw string
        projectId = scriptProperties.getProperty('FIRESTORE_PROJECT_ID');

        // Validate credentials & Log Status
        let missingCred = false;
        if (!clientEmail) { Logger.log(`   -> ${funcName} Error: FIRESTORE_CLIENT_EMAIL not found or empty.`); missingCred = true; } // <<< KEPT: Critical error
        if (!privateKeyRaw) { Logger.log(`   -> ${funcName} Error: FIRESTORE_PRIVATE_KEY not found or empty.`); missingCred = true; } // <<< KEPT: Critical error
        if (!projectId) { Logger.log(`   -> ${funcName} Error: FIRESTORE_PROJECT_ID not found or empty.`); missingCred = true; } // <<< KEPT: Critical error

        if (missingCred) {
            console.error(`${funcName} Error: Missing Firestore credentials in Script Properties.`);
            return null; // Exit if any credential is fundamentally missing
        }

        // --- Process the Private Key String ---
        processedKey = privateKeyRaw;
        // 1. Remove surrounding quotes if present (handle copy-paste variations)
        if (processedKey.startsWith('"') && processedKey.endsWith('"')) {
            processedKey = processedKey.substring(1, processedKey.length - 1);
        }
        // 2. Replace literal "\\n" sequences with actual newline characters "\n"
        processedKey = processedKey.replaceAll('\\n', '\n');
        // --- End Key Processing ---

        // Initialize Firestore using the library and *processed* key
        Logger.log(`   -> Calling FirestoreApp.getFirestore for project ${projectId}...`); // <<< KEPT: Status log
        const firestore = FirestoreApp.getFirestore(clientEmail, processedKey, projectId); // Use processedKey here

        // Check if firestore object was created
        if (!firestore) {
           Logger.log(`   -> ${funcName} Error: FirestoreApp.getFirestore returned null/undefined.`); // <<< KEPT: Critical error
           console.error(`${funcName} Error: FirestoreApp.getFirestore failed to return an instance.`);
           return null;
        }

        Logger.log(`${funcName}: Firestore instance initialized successfully for project ${projectId}.`); // <<< KEPT: Success log
        return firestore;

    } catch (e) {
        // Catch errors during property retrieval or initialization
        console.error(`Error caught in ${funcName}: ${e.message}\nStack: ${e.stack}`);
        Logger.log(`   -> ❌ Exception during Firestore initialization: ${e.message}`); // <<< KEPT: Exception log
        // Log potentially relevant details if available before the error (AVOID logging key)
        Logger.log(`   -> Details at time of error: ProjectID=${projectId || 'N/A'}, Email=${clientEmail || 'N/A'}`); // <<< MODIFIED: Removed key snippet
        return null;
    }
} // END fSrvGetFirestoreInstance




// fSrvSaveTurboDataToFirestore ////////////////////////////////////////////////////
// Purpose -> Saves grid text data (gUI.arr) to a Firestore document within the user's
//            email collection. Document: Collection = userEmail, Doc = Turbo_Game_<csId>.
//            Processes gUI.arr into Array of Row Objects format for Firestore compatibility.
//            Uses updateDocument which creates if not exists.
// Inputs  -> gIndex.Email (String): The user's email address (used as the collection name).
//         -> gIndex.CSID (String): The Character Sheet ID (used for document naming).
//         -> fullArrData (Array[][]): The complete gUI.arr from the client.
//         -> charInfo (Object): DEPRECATED/UNUSED. Character metadata - no longer saved here.
// Outputs -> (Object): { success: Boolean, message?: String } Reflects success/failure of saving grid data ONLY.
function fSrvSaveTurboDataToFirestore(gIndex, fullArrData, charInfo) { // Keep charInfo for compatibility if needed, but ignore
    const funcName = "fSrvSaveTurboDataToFirestore";
    Logger.log(`${funcName}: Saving Grid data for User: ${gIndex.Email}, CS ID: ${gIndex.CSID}...`);

    // === 1. Validate Inputs ===
    if (!gIndex.Email || typeof gIndex.Email !== 'string' || gIndex.Email.indexOf('@') === -1) { // Basic email check
        return { success: false, message: "Invalid User Email provided." };
    }
    if (!gIndex.CSID || typeof gIndex.CSID !== 'string') {
        return { success: false, message: "Invalid Character Sheet ID provided." };
    }
    if (!Array.isArray(fullArrData) || (fullArrData.length > 0 && !Array.isArray(fullArrData[0]))) {
         return { success: false, message: "Invalid fullArrData provided (must be 2D array)." };
    }
    // No longer need to validate charInfo

    // === 2. Get Firestore Instance ===
    const firestore = fSrvGetFirestoreInstance();
    if (!firestore) {
        const msg = "Failed to initialize Firestore instance.";
        Logger.log(`${funcName} Error: ${msg}`);
        return { success: false, message: "Server configuration error (Firestore)." };
    }

    // === 3. Define Path ===
    const collectionPath = gIndex.Email; // User's email is the collection
    const gameDocId = `Turbo_Game_${gIndex.CSID}`;
    // Removed charInfoDocId
    const gameDocPath = `${collectionPath}/${gameDocId}`;
    // Removed charInfoDocPath
    Logger.log(`   -> Target Firestore Game Path: ${gameDocPath}`);
    // Removed CharInfo Path Log

    // === 4. Process Array into Array of Row Objects ===
    const arrayOfRowObjects = [];
    const numRows = fullArrData.length;
    for (let r = 0; r < numRows; r++) {
        const rowData = fullArrData[r] || [];
        const rowKey = `row${r}`;
        const rowObject = {};
        rowObject[rowKey] = rowData;
        arrayOfRowObjects.push(rowObject);
    }

    // === 5. Prepare Data Payloads ===
    const gameDataToSave = {
        gUIarr: arrayOfRowObjects,
        _lastUpdated: new Date()
    };
    // Removed charInfoDataToSave

    // === 6. Save to Firestore (Attempt Game Data Only) ===
    let gameSaveSuccess = false;
    // Removed charInfoSaveSuccess
    let gameSaveError = null;
    // Removed charInfoError

    try {
        Logger.log(`   -> Calling firestore.updateDocument for Game Data: ${gameDocPath}...`);
        firestore.updateDocument(gameDocPath, gameDataToSave, false); // update/create
        gameSaveSuccess = true;
        Logger.log(`      -> ✅ Successfully saved Game Data.`);
    } catch (e) {
        gameSaveError = e;
        console.error(`Exception saving Game Data to ${gameDocPath}: ${e.message}\nStack: ${e.stack}`);
        Logger.log(`   -> ❌ Exception during Game Data save: ${e.message}`);
    }

    // Removed try...catch block for charInfo save

    // === 7. Return Overall Result (Based on Game Data Only) ===
    if (gameSaveSuccess) { // Check only gameSaveSuccess
        Logger.log(`   -> ✅ Successfully saved Game Data document for ${gIndex.CSID} to collection ${gIndex.Email}.`);
        return { success: true };
    } else {
        // Construct detailed error message based only on game save error
        let finalMessage = `Firestore save failed. Game Data Error: ${gameSaveError?.message || 'Unknown'}.`;
        Logger.log(`   -> ❌ Firestore Game Data save failed.`);
        return { success: false, message: finalMessage.trim() };
    }

} // END fSrvSaveTurboDataToFirestore




// fSrvConvertFirestoreTypesToJS //////////////////////////////////////////////////
// Purpose -> Recursively converts Firestore's typed value objects (mapValue,
//            arrayValue, stringValue, etc.) into standard JavaScript types
//            (objects, arrays, strings, numbers, booleans).
// Inputs  -> firestoreValue (Object): A value object from Firestore (e.g.,
//            doc.fields.someProperty or an element within an arrayValue/mapValue).
// Outputs -> (Any): The corresponding standard JavaScript value or type.
function fSrvConvertFirestoreTypesToJS(firestoreValue) {
  if (!firestoreValue) return firestoreValue; // Handle null/undefined cases safely

  // Check for primitive types
  if (firestoreValue.stringValue !== undefined) return firestoreValue.stringValue;
  if (firestoreValue.integerValue !== undefined) return parseInt(firestoreValue.integerValue, 10); // Parse as integer
  if (firestoreValue.doubleValue !== undefined) return parseFloat(firestoreValue.doubleValue); // Parse as float
  if (firestoreValue.booleanValue !== undefined) return firestoreValue.booleanValue;
  if (firestoreValue.nullValue !== undefined) return null;
  if (firestoreValue.timestampValue !== undefined) return new Date(firestoreValue.timestampValue); // Convert to JS Date

  // Check for bytesValue, geoPointValue if you use them, otherwise ignore or return placeholder

  // Check for array type
  if (firestoreValue.arrayValue && firestoreValue.arrayValue.values) {
    // It's an array, recursively convert its elements
    return firestoreValue.arrayValue.values.map(element => fSrvConvertFirestoreTypesToJS(element));
  }

  // Check for map/object type
  if (firestoreValue.mapValue && firestoreValue.mapValue.fields) {
    // It's an object/map, recursively convert its properties
    const jsObject = {};
    for (const key in firestoreValue.mapValue.fields) {
      jsObject[key] = fSrvConvertFirestoreTypesToJS(firestoreValue.mapValue.fields[key]);
    }
    return jsObject;
  }

  // If it's none of the known Firestore types (e.g., already a JS primitive passed in), return as is.
  // Or potentially log a warning if an unexpected structure is encountered.
  Logger.log(`fSrvConvertFirestoreTypesToJS: Encountered unexpected value structure: ${JSON.stringify(firestoreValue).substring(0,100)}... Returning as is.`);
  return firestoreValue;
} // END fSrvConvertFirestoreTypesToJS




// fSrvUnpackFirestoreArrayTo2D //////////////////////////////////////////////
// Purpose -> Converts the Firestore array-of-row-objects format into a standard
//            2D JavaScript array server-side. Handles potential sparse arrays.
// Inputs  -> firestoreArr (Array): Array from Firestore after type conversion,
//            e.g., [ {"row0":[...]}, {"row1":[...]}, ... ]
// Outputs -> (Array[][]): A standard 2D JavaScript array. Returns empty array on error.
function fSrvUnpackFirestoreArrayTo2D(firestoreArr) {
    const funcName = "fSrvUnpackFirestoreArrayTo2D";
    if (!Array.isArray(firestoreArr)) {
        Logger.log(`${funcName}: Input is not an array. Returning empty array.`);
        return [];
    }

    const new2DArray = [];
    let maxRow = -1;
    let maxCols = 0;

    // First pass: Populate based on keys, find max row/col
    for (const rowObject of firestoreArr) {
        if (typeof rowObject !== 'object' || rowObject === null) continue; // Skip non-objects

        const key = Object.keys(rowObject)[0];
        if (!key || !key.startsWith('row')) continue; // Skip invalid keys

        const rowNum = parseInt(key.substring(3), 10);
        if (isNaN(rowNum)) continue; // Skip if key is not like "rowX"

        const rowData = rowObject[key];
        if (!Array.isArray(rowData)) {
             Logger.log(`${funcName}: Value for key ${key} is not an array. Skipping.`);
             continue; // Ensure the value is actually an array
        }

        new2DArray[rowNum] = rowData;
        if (rowNum > maxRow) maxRow = rowNum;
        if (rowData.length > maxCols) maxCols = rowData.length; // Track max column length
    }

    // Second pass: Fill potential sparse gaps and ensure consistent column length
    for (let r = 0; r <= maxRow; r++) {
        if (typeof new2DArray[r] === 'undefined') {
            new2DArray[r] = Array(maxCols).fill(''); // Initialize sparse rows with empty strings
        } else {
            // Ensure existing rows have the correct length
            while (new2DArray[r].length < maxCols) {
                new2DArray[r].push(''); // Pad with empty strings
            }
        }
    }
    // Ensure the final array itself isn't sparse (if maxRow was > initial length)
    while(new2DArray.length <= maxRow) {
        new2DArray.push(Array(maxCols).fill(''));
    }


    // Handle case where firestoreArr was empty
    if (maxRow === -1) {
        return [[]]; // Return array with one empty row if input was empty
    }


    return new2DArray;
} // END fSrvUnpackFirestoreArrayTo2D




// fSrvCheckAndLoadFirestoreGUIarrAs2D //////////////////////////////////////////////////
// Purpose -> Checks Firestore for a document containing saved gUI.arr data based on CS ID
//            within the specified user's email collection. If found, extracts the gUIarr
//            field, converts types, unpacks it into a 2D array, and returns it.
//            Handles document not found and errors gracefully.
// Inputs  -> gIndex.Email (String): The user's email (Firestore collection name).
//         -> gIndex.CSID (String): The Character Sheet ID (part of document name).
// Outputs -> (Object): { success: Boolean, firestoreArr?: Array[][], message?: String }
//            'firestoreArr' is the standard 2D array if success is true.
function fSrvCheckAndLoadFirestoreGUIarrAs2D(gIndex) {
    const funcName = "fSrvCheckAndLoadFirestoreGUIarrAs2D";
    Logger.log(`${funcName}: Checking Firestore for gUI.arr data for User: ${gIndex.Email}, CS ID: ${gIndex.CSID}...`);

    // === 1. Validate Inputs ===
    if (!gIndex.Email || typeof gIndex.Email !== 'string' || gIndex.Email.indexOf('@') === -1) {
        const msg = "Invalid User Email provided.";
        Logger.log(`${funcName} Error: ${msg}`);
        return { success: false, message: msg };
    }
    if (!gIndex.CSID || typeof gIndex.CSID !== 'string') {
        const msg = "Invalid Character Sheet ID provided.";
        Logger.log(`${funcName} Error: ${msg}`);
        return { success: false, message: msg };
    }

    // === 2. Get Firestore Instance ===
    const firestore = fSrvGetFirestoreInstance();
    if (!firestore) {
        const msg = "Failed to initialize Firestore instance.";
        Logger.log(`${funcName} Error: ${msg}`);
        return { success: false, message: "Server configuration error (Firestore)." };
    }

    // === 3. Define Path and Fetch Data ===
    const collectionPath = gIndex.Email;
    const documentId = `Turbo_Game_${gIndex.CSID}`; // Specific document for gUI.arr
    const documentPath = `${collectionPath}/${documentId}`;
    Logger.log(`   -> Target Firestore Path: ${documentPath}`);

    try {
        // Attempt to get the document
        const doc = firestore.getDocument(documentPath);

        // === 4. Check if Document Exists & Has Data ===
        // Check specifically for the gUIarr field within fields
        if (!doc || !doc.fields || !doc.fields.gUIarr) {
            const msg = `Document not found or missing 'gUIarr' field at path: ${documentPath}.`;
            Logger.log(`   -> ${funcName}: ${msg}`);
            return { success: false, message: "No saved grid data found in Firestore for this character." };
        }
        Logger.log(`   -> Document found. Processing 'gUIarr' field...`);

        // === 5. Convert Firestore Types for gUIarr ===
        const arrDataRaw = doc.fields.gUIarr;
        const arrDataConverted = fSrvConvertFirestoreTypesToJS(arrDataRaw);

        // === 6. Validate Converted Data (Should be Array of Objects) ===
        if (!Array.isArray(arrDataConverted)) {
            const msg = `Invalid data type for gUIarr after conversion. Expected array, got ${typeof arrDataConverted}. Path: ${documentPath}`;
            Logger.log(`   -> ${funcName} Error: ${msg}`);
            return { success: false, message: "Invalid grid data format retrieved from Firestore." };
        }

        // === 7. Unpack Array of Objects into 2D Array ===
        const unpackedArr = fSrvUnpackFirestoreArrayTo2D(arrDataConverted);

        // === 8. Return Success with Unpacked Data ===
        Logger.log(`   -> ✅ Successfully fetched and unpacked Firestore gUI.arr data for ${gIndex.CSID}.`);
        return {
            success: true,
            firestoreArr: unpackedArr // Return the standard 2D array
        };

    } catch (e) {
        // Handle potential errors during Firestore operations
        const safeErrorMessage = e.message?.includes("permission")
                               ? "Permission denied accessing Firestore."
                               : e.message?.includes("NOT_FOUND")
                               ? "No saved grid data found in Firestore for this character."
                               : `Server error during Firestore read: ${e.message || e}`;

        console.error(`Exception caught in ${funcName} fetching from path ${documentPath}: ${e.message}\nStack: ${e.stack}`);
        Logger.log(`   -> ❌ Exception during Firestore read for ${gIndex.CSID}: ${safeErrorMessage}`);
        return { success: false, message: safeErrorMessage };
    }
} // END fSrvCheckAndLoadFirestoreGUIarrAs2D




// fSrvSaveFullSheetTextAndTagsToFirestore //////////////////////////////////////
// Purpose -> Saves loaded sheet tags (ColTags, RowTags) and the full sheet text data
//            (sheetText2D, converted to array-of-row-objects) to Firestore, potentially
//            slicing the data into multiple documents if it exceeds size estimates.
// Inputs  -> gIndex (Object): Contains Email, CSID, GameVer.
//         -> workbookAbr (String): Abbreviation ('db', 'mastercs', 'masterkl', 'mycs', 'mykl').
//         -> sheetName (String): The name of the sheet that was loaded.
//         -> data (Object): The object returned by fSrvLoadFullGoogleSheetAndTags, containing
//                             { ColTags, RowTags, sheetText2D }.
// Outputs -> (Object): { success: Boolean, message?: String }
function fSrvSaveFullSheetTextAndTagsToFirestore(gIndex, workbookAbr, sheetName, data) {
    const funcName = "fSrvSaveFullSheetTextAndTagsToFirestore";
    const MAX_CHUNK_SIZE_ESTIMATE = 500000; // Target ~500KB per data chunk
    Logger.log(`${funcName}: Saving data for Workbook: "${workbookAbr}", Sheet: "${sheetName}", Version: ${gIndex?.GameVer}, Email: ${gIndex?.Email}, CSID: ${gIndex?.CSID}...`);

    // --- 1. Validate Inputs ---
    const lowerWorkbookAbr = workbookAbr?.toLowerCase() || '';
    const trimmedSheetName = sheetName?.trim() || '';
    if (!data || typeof data !== 'object' || !data.ColTags || !data.RowTags || !data.sheetText2D || !Array.isArray(data.sheetText2D)) {
        return { success: false, message: "Invalid data object provided (missing ColTags, RowTags, or sheetText2D array)." };
    }
    if (!trimmedSheetName) {
        return { success: false, message: "Invalid or empty sheetName provided." };
    }

    // --- 2. Get Firestore Instance ---
    const firestore = fSrvGetFirestoreInstance();
    if (!firestore) {
        const msg = "Failed to initialize Firestore instance.";
        Logger.log(`${funcName} Error: ${msg}`);
        return { success: false, message: "Server configuration error (Firestore)." };
    }

    // --- 3. Determine Firestore Path using Helper ---
    let baseCollectionName;
    let baseDocumentId;
    let documentPathBase; // For logging clarity
    try {
        // Path calculation requires valid gIndex properties, will throw if invalid
        const pathInfo = fSrvCalcFirestorePath(workbookAbr, trimmedSheetName, gIndex);
        baseCollectionName = pathInfo.collectionName;
        baseDocumentId = pathInfo.documentId;
        documentPathBase = `${baseCollectionName}/${baseDocumentId}`; // For logging
        Logger.log(`   -> Base Firestore Path Calculated: ${documentPathBase}`);
    } catch (pathError) {
        Logger.log(`   -> ❌ Error determining Firestore path: ${pathError.message}`);
        return { success: false, message: pathError.message }; // Return error from helper
    }

    // --- 4. Convert sheetText2D to Array of Row Objects ---
    const sheetTextArray = data.sheetText2D;
    const arrayOfRowObjects = [];
    const numRows = sheetTextArray.length;
    for (let r = 0; r < numRows; r++) {
        const rowData = sheetTextArray[r] || [];
        const rowKey = `row${r}`;
        const rowObject = {};
        rowObject[rowKey] = rowData;
        arrayOfRowObjects.push(rowObject);
    }
    Logger.log(`   -> Converted ${numRows} rows to array-of-row-objects format.`);

    // --- 5. Slicing Logic ---
    const chunks = [];
    let currentChunk = [];
    let currentChunkSizeEstimate = 0;

    Logger.log(`   -> Slicing data based on estimated size (Target: ${MAX_CHUNK_SIZE_ESTIMATE} bytes)...`);
    for (let i = 0; i < arrayOfRowObjects.length; i++) {
        const rowObject = arrayOfRowObjects[i];
        let rowObjectSizeEstimate = 0;
        try {
            // Estimate size of the single row object
            rowObjectSizeEstimate = JSON.stringify(rowObject).length;
        } catch (e) {
            Logger.log(`   -> Warning: Could not estimate size for row object at index ${i}. Assuming small size (0). Error: ${e.message}`);
            // Proceed cautiously if stringify fails for a single row
        }

        // Check if adding this row would exceed the limit for the current chunk
        if (currentChunk.length > 0 && (currentChunkSizeEstimate + rowObjectSizeEstimate > MAX_CHUNK_SIZE_ESTIMATE)) {
            // Current chunk is full (or adding next row exceeds limit), push it and start new
            chunks.push(currentChunk);
            Logger.log(`      -> Chunk ${chunks.length} finalized with ${currentChunk.length} rows (Estimated size: ${currentChunkSizeEstimate} bytes).`);
            currentChunk = [rowObject]; // Start new chunk with current row object
            currentChunkSizeEstimate = rowObjectSizeEstimate; // Reset size estimate
        } else {
            // Add to current chunk
            currentChunk.push(rowObject);
            currentChunkSizeEstimate += rowObjectSizeEstimate;
        }
    }
    // Add the last remaining chunk if it has data
    if (currentChunk.length > 0) {
        chunks.push(currentChunk);
        Logger.log(`      -> Chunk ${chunks.length} finalized with ${currentChunk.length} rows (Estimated size: ${currentChunkSizeEstimate} bytes).`);
    }

    // Determine total chunks (must be at least 1 if there was data)
    const totalChunks = (arrayOfRowObjects.length > 0) ? Math.max(1, chunks.length) : 0;
    // If input array was empty, chunks will be empty, totalChunks=0.
    // If input array had data but was small, chunks.length will be 1, totalChunks=1.
    Logger.log(`   -> Total data chunks determined: ${totalChunks}`);

    // --- 6. Prepare and Save Metadata Document ---
    const metadataDocId = `${baseDocumentId}_metadata`;
    const metadataPath = `${baseCollectionName}/${metadataDocId}`;
    const metadataObject = {
        ColTags: data.ColTags,
        RowTags: data.RowTags,
        totalChunks: totalChunks, // Save the calculated number of chunks
        _lastUpdated: new Date()
    };
    let metadataSaveSuccess = false;
    try {
        Logger.log(`   -> Saving Metadata Document to: ${metadataPath}`);
        firestore.updateDocument(metadataPath, metadataObject, false); // update/create
        metadataSaveSuccess = true;
        Logger.log(`      -> ✅ Successfully saved Metadata Document.`);
    } catch (e) {
        const errorMsg = `Failed to save Metadata Document (${metadataPath}): ${e.message || e}`;
        console.error(`${funcName} Error: ${errorMsg}\nStack: ${e.stack}`);
        Logger.log(`   -> ❌ ${errorMsg}`);
        return { success: false, message: errorMsg }; // Critical failure if metadata can't save
    }

    // --- 7. Save Data Chunk Documents ---
    let allChunksSaved = true; // Assume success until a chunk fails
    if (totalChunks > 0) {
        for (let i = 0; i < totalChunks; i++) {
            const chunkIndex = i + 1; // 1-based index for naming
            const chunkDocId = `${baseDocumentId}_${chunkIndex}of${totalChunks}`;
            const chunkPath = `${baseCollectionName}/${chunkDocId}`;
            const chunkData = {
                rowDataChunk: chunks[i], // The actual slice of arrayOfRowObjects
                _lastUpdated: new Date()
            };

            try {
                Logger.log(`   -> Saving Data Chunk ${chunkIndex}/${totalChunks} to: ${chunkPath}`);
                firestore.updateDocument(chunkPath, chunkData, false); // update/create
                Logger.log(`      -> ✅ Successfully saved Data Chunk ${chunkIndex}/${totalChunks}.`);
            } catch (e) {
                const errorMsg = `Failed to save Data Chunk ${chunkIndex}/${totalChunks} (${chunkPath}): ${e.message || e}`;
                console.error(`${funcName} Error: ${errorMsg}\nStack: ${e.stack}`);
                Logger.log(`   -> ❌ ${errorMsg}`);
                allChunksSaved = false; // Mark failure but continue trying others
                // Optional: Collect individual error messages if needed
            }
        }
    } else {
        Logger.log(`   -> No data chunks to save (source data likely empty).`);
        // If there were no rows, metadata still saved, consider this overall success.
    }

    // --- 8. Return Overall Result ---
    if (metadataSaveSuccess && allChunksSaved) {
        Logger.log(`   -> ✅ Successfully saved Metadata and all ${totalChunks} Data Chunk(s).`);
        return { success: true };
    } else {
        const finalMessage = `Firestore save partially failed. Metadata saved: ${metadataSaveSuccess}. All data chunks saved: ${allChunksSaved}. Check logs for details.`;
        Logger.log(`   -> ❌ ${finalMessage}`);
        return { success: false, message: finalMessage };
    }

} // END fSrvSaveFullSheetTextAndTagsToFirestore




// fSrvCalcFirestorePath //////////////////////////////////////////////////////////
// Purpose -> Determines the Firestore collection and document ID based on the
//            workbook type and other parameters.
// Inputs  -> workbookAbr (String): Abbreviation ('db', 'mastercs', 'masterkl', 'mycs', 'mykl').
//         -> sheetName (String): The name of the sheet.
//         -> gIndex.GameVer (String): The game version (e.g., "28.3"), required for 'db'/'master*' types.
//         -> gIndex.Email (String): The user's Email, required for 'mycs'/'mykl' types.
//         -> gIndex.CSID (String): The Character Sheet ID, required for 'mycs'/'mykl' types.
// Outputs -> (Object): { collectionName: String, documentId: String } on success.
// Throws  -> (Error): If inputs are invalid or workbook abbreviation is unsupported.
function fSrvCalcFirestorePath(workbookAbr, sheetName, gIndex) {
    const funcName = "fSrvCalcFirestorePath";

    // --- 1. Validate Inputs ---
    if (!workbookAbr || typeof workbookAbr !== 'string') {
        throw new Error(`${funcName}: Invalid or missing workbookAbr provided.`);
    }
    if (!sheetName || typeof sheetName !== 'string' || sheetName.trim() === '') {
        throw new Error(`${funcName}: Invalid or empty sheetName provided.`);
    }
    const lowerWorkbookAbr = workbookAbr.toLowerCase();
    const trimmedSheetName = sheetName.trim();
    const validWorkbooks = ['db', 'mastercs', 'masterkl', 'mycs', 'mykl'];
    if (!validWorkbooks.includes(lowerWorkbookAbr)) {
        throw new Error(`${funcName}: Unsupported workbook abbreviation: "${workbookAbr}"`);
    }

    // --- 2. Initialize Variables ---
    let collectionName = '';
    let documentId = '';

    // --- 3. Determine Collection and Document ID ---
    switch (lowerWorkbookAbr) {
        case 'db':
        case 'mastercs':
        case 'masterkl':
            if (!gIndex.GameVer || typeof gIndex.GameVer !== 'string' || gIndex.GameVer.trim() === '') {
                throw new Error(`${funcName}: Game Version is required for workbook type '${workbookAbr}'.`);
            }
            // Assign collection name based on abbreviation, ensuring proper casing
            if (lowerWorkbookAbr === 'db') collectionName = 'DB';
            else if (lowerWorkbookAbr === 'mastercs') collectionName = 'MasterCS';
            else collectionName = 'MasterKL'; // Must be masterkl
            documentId = `v${gIndex.GameVer.trim()} ${trimmedSheetName}`; // Format: vVERSION SHEETNAME
            break;
        case 'mycs':
        case 'mykl':
            if (!gIndex.Email || typeof gIndex.Email !== 'string' || gIndex.Email.indexOf('@') === -1) {
                throw new Error(`${funcName}: User Email is required and must be valid for workbook type '${workbookAbr}'.`);
            }
            if (!gIndex.CSID || typeof gIndex.CSID !== 'string') {
                throw new Error(`${funcName}: Character Sheet ID is required for workbook type '${workbookAbr}'.`);
            }
            collectionName = gIndex.Email; // User's gIndex.Email is the collection
            if (lowerWorkbookAbr === 'mycs') {
                documentId = `MyCS_${trimmedSheetName}_${gIndex.CSID}`;
            } else { // Must be 'mykl'
                documentId = `MyKL_${trimmedSheetName}_OfMyCS_${gIndex.CSID}`;
            }
            break;
        // Default case is handled by the initial validation check for validWorkbooks
    }

    // --- 4. Final Validation and Return ---
    if (!collectionName || !documentId) {
        // This should theoretically not happen if logic above is correct, but as a safeguard:
        throw new Error(`${funcName}: Failed to determine collectionName or documentId for workbook '${workbookAbr}'.`);
    }

    // Logger.log(`${funcName}: Determined Path - Collection: "${collectionName}", Document: "${documentId}"`); // Optional log
    return { collectionName, documentId };

} // END fSrvCalcFirestorePath




// fSrvGetFirestoreFSData ///////////////////////////////////////////////////////////
// Purpose -> Reads data from a Firestore document (potentially sliced across multiple
//            documents) previously saved by fSrvSaveFullSheetTextAndTagsToFirestore.
//            Reassembles sliced data and returns the full sheet content and absolute tag maps.
// Inputs  -> workbookAbr (String): Workbook abbreviation ('db', 'mycs', etc.).
//         -> sheetName (String): The sheet name associated with the data.
//         -> gIndex (Object): Object containing CSID, GameVer, Email.
// Outputs -> (Object): On success: { success: true, FSData: { colTagsMap, rowTagsMap, text } }
//                     On failure: { success: false, message: String }
function fSrvGetFirestoreFSData(workbookAbr, sheetName, gIndex) {
    const funcName = "fSrvGetFirestoreFSData";
    Logger.log(`${funcName}: Reading document(s) for Workbook: "${workbookAbr}", Sheet: "${sheetName}", Ver: ${gIndex?.GameVer}, Email: ${gIndex?.Email}, CSID: ${gIndex?.CSID}...`);

    let firestore;
    let baseCollectionName;
    let baseDocumentId;
    let metadataPath;
    let absoluteColTagMap = {}; // Initialize in case of zero chunks
    let absoluteRowTagMap = {}; // Initialize in case of zero chunks

    try {
        // --- 1. Get Firestore Instance ---
        firestore = fSrvGetFirestoreInstance();
        if (!firestore) {
            return { success: false, message: "Server configuration error (Firestore)." };
        }

        // --- 2. Calculate Firestore Path ---
        try {
            const pathInfo = fSrvCalcFirestorePath(workbookAbr, sheetName, gIndex);
            baseCollectionName = pathInfo.collectionName;
            baseDocumentId = pathInfo.documentId;
        } catch (pathError) {
            Logger.log(`   -> ❌ Error determining Firestore path: ${pathError.message}`);
            return { success: false, message: pathError.message };
        }
        metadataPath = `${baseCollectionName}/${baseDocumentId}_metadata`;
        Logger.log(`   -> Target Metadata Path: ${metadataPath}`);

        // --- 3. Fetch Metadata Document ---
        let metadataDoc;
        try {
            metadataDoc = firestore.getDocument(metadataPath);
        } catch (e) {
            // Catch potential "NOT_FOUND" or permission errors specifically from getDocument
            const isNotFoundError = e.message?.toUpperCase().includes("NOT_FOUND");
            const errorMsg = isNotFoundError
                           ? `Metadata document not found at path: ${metadataPath}. Data may be missing or not yet saved.`
                           : `Error fetching metadata document (${metadataPath}): ${e.message || e}`;
            Logger.log(`   -> ${funcName}: ${errorMsg}`);
            return { success: false, message: errorMsg };
        }

        // --- 4. Validate Metadata & Extract Info ---
        if (!metadataDoc || !metadataDoc.fields) {
            const msg = `Metadata document not found or empty at path: ${metadataPath}.`;
            Logger.log(`   -> ${funcName}: ${msg}`);
            return { success: false, message: msg };
        }

        const colTagsRaw = metadataDoc.fields.ColTags;
        const rowTagsRaw = metadataDoc.fields.RowTags;
        const totalChunksRaw = metadataDoc.fields.totalChunks;

        if (!colTagsRaw || typeof colTagsRaw.mapValue === 'undefined' ||
            !rowTagsRaw || typeof rowTagsRaw.mapValue === 'undefined' ||
            !totalChunksRaw || typeof totalChunksRaw.integerValue === 'undefined') {
            const msg = "Invalid metadata document structure found (missing/invalid ColTags, RowTags, or totalChunks).";
            Logger.log(`   -> ${funcName} Error: ${msg}`);
            return { success: false, message: msg };
        }

        absoluteColTagMap = fSrvConvertFirestoreTypesToJS(colTagsRaw);
        absoluteRowTagMap = fSrvConvertFirestoreTypesToJS(rowTagsRaw);
        const totalChunks = parseInt(totalChunksRaw.integerValue, 10);

        if (typeof absoluteColTagMap !== 'object' || absoluteColTagMap === null || Array.isArray(absoluteColTagMap) ||
            typeof absoluteRowTagMap !== 'object' || absoluteRowTagMap === null || Array.isArray(absoluteRowTagMap) ||
            isNaN(totalChunks) || totalChunks < 0) {
            const msg = "Invalid data types found in metadata after conversion (ColTags/RowTags not objects, or totalChunks not integer >= 0).";
            Logger.log(`   -> ${funcName} Error: ${msg}`);
            return { success: false, message: msg };
        }
        Logger.log(`   -> Metadata validated. Total Chunks: ${totalChunks}. ColTags: ${Object.keys(absoluteColTagMap).length}, RowTags: ${Object.keys(absoluteRowTagMap).length}`);

        // --- 5. Handle Zero Chunks ---
        if (totalChunks === 0) {
            Logger.log(`   -> Total chunks is 0. Returning empty data structure.`);
            return { success: true, FSData: { colTagsMap: absoluteColTagMap, rowTagsMap: absoluteRowTagMap, text: [[]] } }; // Return empty 2D array
        }

        // --- 6. Fetch Data Chunks ---
        const fetchedChunkDocs = [];
        const missingChunks = [];
        Logger.log(`   -> Attempting to fetch ${totalChunks} data chunk(s)...`);
        for (let i = 1; i <= totalChunks; i++) {
            const chunkDocId = `${baseDocumentId}_${i}of${totalChunks}`;
            const chunkPath = `${baseCollectionName}/${chunkDocId}`;
            try {
                const chunkDoc = firestore.getDocument(chunkPath);
                if (chunkDoc && chunkDoc.fields && chunkDoc.fields.rowDataChunk) {
                    fetchedChunkDocs.push(chunkDoc); // Store the whole doc for now
                    // Logger.log(`      -> Successfully fetched chunk ${i}/${totalChunks}.`); // Can be noisy
                } else {
                    missingChunks.push(i);
                    Logger.log(`      -> ❌ Failed to fetch or find valid 'rowDataChunk' in chunk ${i}/${totalChunks} at ${chunkPath}.`);
                }
            } catch (e) {
                const isNotFoundError = e.message?.toUpperCase().includes("NOT_FOUND");
                Logger.log(`      -> ❌ Exception fetching chunk ${i}/${totalChunks} at ${chunkPath}: ${e.message || e}${isNotFoundError ? ' (NOT_FOUND)' : ''}`);
                missingChunks.push(i); // Mark as missing on error too
            }
        }

        // --- 7. Error Check: Ensure All Chunks Were Fetched ---
        if (missingChunks.length > 0) {
            const errorMsg = `Failed to load all required data chunks. Missing chunk(s): ${missingChunks.join(', ')} of ${totalChunks}. Data is incomplete.`;
            Logger.log(`   -> ${funcName} Error: ${errorMsg}`);
            return { success: false, message: errorMsg };
        }
        Logger.log(`   -> Successfully fetched all ${totalChunks} data chunk(s).`);

        // --- 8. Reassemble Data ---
        const combinedRowObjects = [];
        Logger.log(`   -> Reassembling data from chunks...`);
        for (let i = 0; i < fetchedChunkDocs.length; i++) {
            const chunkDoc = fetchedChunkDocs[i];
            const chunkIndex = i + 1; // 1-based for logging
            const rowDataChunkRaw = chunkDoc.fields.rowDataChunk;
            const rowDataChunkConverted = fSrvConvertFirestoreTypesToJS(rowDataChunkRaw);

            if (!Array.isArray(rowDataChunkConverted)) {
                const errorMsg = `Invalid rowDataChunk format found in chunk ${chunkIndex} after conversion (expected array).`;
                Logger.log(`   -> ${funcName} Error: ${errorMsg}`);
                return { success: false, message: errorMsg };
            }
            combinedRowObjects.push(...rowDataChunkConverted); // Concatenate arrays
        }
        Logger.log(`   -> Reassembled ${combinedRowObjects.length} total row objects.`);

        // --- 9. Unpack and Format Final FSData ---
        const fullData2D = fSrvUnpackFirestoreArrayTo2D(combinedRowObjects);
        const numRowsFinal = fullData2D.length;
        const numColsFinal = fullData2D[0]?.length || 0;
        Logger.log(`   -> Unpacked reassembled data into final 2D array (${numRowsFinal}x${numColsFinal}).`);

        const assembledFSDataObject = {
            colTagsMap: absoluteColTagMap, // Use the absolute tags from metadata
            rowTagsMap: absoluteRowTagMap, // Use the absolute tags from metadata
            text: fullData2D              // The fully reassembled 2D data array
        };

        // --- 10. Return Success ---
        Logger.log(`   -> ✅ Successfully read and formatted sliced Firestore data.`);
        return {
            success: true,
            FSData: assembledFSDataObject
        };

    } catch (e) {
        // Catch errors from Firestore calls, path calculation, tag resolution, etc.
        const safeErrorMessage = e.message?.includes("permission") ? "Permission denied accessing Firestore."
                               : `Server error during Firestore read/process: ${e.message || e}`;

        console.error(`Exception caught in ${funcName} accessing path ${metadataPath || 'Unknown'}: ${e.message}\nStack: ${e.stack}`);
        Logger.log(`   -> ❌ Exception during Firestore read/process for ${metadataPath || 'Unknown'}: ${safeErrorMessage}`);
        return { success: false, message: safeErrorMessage };
    }

} // END fSrvGetFirestoreFSData




// fSrvResolveRangeTagsForFirestore //////////////////////////////////////////////////
// Purpose -> Resolve a range object (tags/indices) into absolute 0-based numeric indices
//            using provided Firestore tag maps.
// Inputs  -> rangeObject (Object): { r1, c1, r2, c2 } with tag strings or 0-based indices.
//         -> rowTagMap (Object): Map of row tags to indices from Firestore.
//         -> colTagMap (Object): Map of column tags to indices from Firestore.
// Outputs -> (Object): { r1, c1, r2, c2 } with resolved 0-based numeric indices.
// Throws  -> (Error): If validation fails or tags cannot be resolved.
function fSrvResolveRangeTagsForFirestore(rangeObject, rowTagMap, colTagMap) {
    const funcName = "fSrvResolveRangeTagsForFirestore";

    // --- 1. Validate Inputs ---
    if (!rangeObject || typeof rangeObject !== 'object') {
        throw new Error(`${funcName}: Invalid rangeObject provided.`);
    }
    if (typeof rowTagMap !== 'object' || rowTagMap === null) {
        throw new Error(`${funcName}: Invalid rowTagMap provided.`);
    }
     if (typeof colTagMap !== 'object' || colTagMap === null) {
        throw new Error(`${funcName}: Invalid colTagMap provided.`);
    }
     if (rangeObject.r1 === undefined || rangeObject.c1 === undefined || rangeObject.r2 === undefined || rangeObject.c2 === undefined) {
         throw new Error(`${funcName}: rangeObject must contain r1, c1, r2, and c2 properties.`);
    }

    // --- 2. Resolve Tags/Indices using fSrvResolveTag ---
    const r1 = fSrvResolveTag(rangeObject.r1, rowTagMap, 'row');
    const c1 = fSrvResolveTag(rangeObject.c1, colTagMap, 'col');
    const r2 = fSrvResolveTag(rangeObject.r2, rowTagMap, 'row');
    const c2 = fSrvResolveTag(rangeObject.c2, colTagMap, 'col');

    // --- 3. Check for Resolution Errors ---
    const errors = [];
    if (isNaN(r1)) errors.push(`r1 ('${rangeObject.r1}')`);
    if (isNaN(c1)) errors.push(`c1 ('${rangeObject.c1}')`);
    if (isNaN(r2)) errors.push(`r2 ('${rangeObject.r2}')`);
    if (isNaN(c2)) errors.push(`c2 ('${rangeObject.c2}')`);

    if (errors.length > 0) {
        throw new Error(`${funcName}: Failed to resolve the following tags/indices: ${errors.join(', ')}.`);
    }

    // --- 4. Return Resolved Indices ---
    return { r1, c1, r2, c2 };

} // END fSrvResolveRangeTagsForFirestore




// fSrvVerifyFirestorePathExists ////////////////////////////////////////////////////
// Purpose -> Checks if the *metadata* Firestore document exists for a given
//            workbook/sheet combination, based on the calculated path.
// Inputs  -> workbookAbr (String): Workbook abbreviation ('db', 'mycs', etc.).
//         -> sheetName (String): The sheet name associated with the data.
//         -> gIndex.GameVer (String): Game version (required for 'db'/'master*').
//         -> gIndex.Email (String): User Email (required for 'mycs'/'mykl').
//         -> gIndex.CSID (String): Character Sheet ID (required for 'mycs'/'mykl').
// Outputs -> (Boolean): True if the metadata document exists, false otherwise (or on error).
function fSrvVerifyFirestorePathExists(workbookAbr, sheetName, gIndex) {
    const funcName = "fSrvVerifyFirestorePathExists";
    // Logger.log(`${funcName}: Verifying path for Workbook: "${workbookAbr}", Sheet: "${sheetName}", Version: ${gIndex.GameVer}, gIndex.Email: ${gIndex.Email}, gIndex.CSID: ${gIndex.CSID}...`); // Reduced logging

    let firestore;
    let metadataPath; // Changed variable name for clarity
    try {
        // --- 1. Get Firestore Instance ---
        firestore = fSrvGetFirestoreInstance();
        if (!firestore) {
            Logger.log(`   -> ${funcName}: Firestore initialization failed. Cannot verify path.`);
            return false; // Cannot verify if Firestore isn't available
        }

        // --- 2. Calculate Firestore Path for METADATA ---
        let baseCollectionName;
        let baseDocumentId;
        // Use try-catch here as fSrvCalcFirestorePath throws errors on invalid inputs
        try {
            const pathInfo = fSrvCalcFirestorePath(workbookAbr, sheetName, gIndex);
            baseCollectionName = pathInfo.collectionName;
            baseDocumentId = pathInfo.documentId; // Get the base ID
        } catch (pathError) {
            Logger.log(`   -> ${funcName}: Error calculating base path: ${pathError.message}. Assuming path does not exist.`);
            return false; // Cannot exist if path is invalid
        }

        // --- Construct path to the METADATA document ---
        metadataPath = `${baseCollectionName}/${baseDocumentId}_metadata`; // Append _metadata
        Logger.log(`   -> Calculated Firestore Metadata Path to check: ${metadataPath}`);

        // --- 3. Check Metadata Document Existence ---
        const doc = firestore.getDocument(metadataPath);

        // Check if the document object has an updateTime property (indicates existence)
        if (doc && doc.updateTime) {
             Logger.log(`   -> Metadata document found at path: ${metadataPath}. Exists: true.`);
             return true; // Document exists
        } else {
             Logger.log(`   -> Metadata document NOT found at path: ${metadataPath}. Exists: false.`);
             return false; // Document does not exist (or has no fields/metadata)
        }

    } catch (e) {
        // Catch other potential errors during Firestore getDocument call (e.g., permissions)
        const isNotFoundError = e.message && e.message.toUpperCase().includes("NOT_FOUND");
        if (isNotFoundError) {
             Logger.log(`   -> ${funcName}: Explicit NOT_FOUND error for metadata path ${metadataPath}. Exists: false.`);
             return false;
        } else {
            // Log other errors but return false as existence couldn't be confirmed
            console.error(`Exception caught in ${funcName} accessing metadata path ${metadataPath || 'Unknown'}: ${e.message}\nStack: ${e.stack}`);
            Logger.log(`   -> ❌ Exception during Firestore metadata check for ${metadataPath || 'Unknown'}: ${e.message}. Assuming path does not exist.`);
            return false;
        }
    }

} // END fSrvVerifyFirestorePathExists


