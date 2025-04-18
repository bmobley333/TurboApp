// turbo.gs //
// 04.14.2025 //



// ==========================================================================
// === Global Variables & doGet ===
// ==========================================================================



const gSrv = { // Using gSrv prefix for server-side globals
  ids: {
    sheets: {
      db: '1m7-VRDey6rPFDm_ZHQy4GOGNRxR0Th2IVY9TUVQ5DdU',
      mastercs: '1zVjGg2KKzEAdX6WezkatGPV0cahWXGIdlZASIUlhf_o',
      masterkl: '1Gcbnggc9UnQrzAcGwSkKdb_jtwxF3u8x1Fb-BdsECZ0',
      ps: '1cu4tsBeQg4l2czraYDjdVpn1zsIKPhYi5czbGf9ZeZg',
      sg: '190vk3Dcqdux_pFjuryvmMx2dWTnJyl23cZvTGhM3kgU',
      // Player-specific cs/kl will be handled dynamically via functions below
      cs: '',  // this is equal to the old myCS and is the ID from the player's CS (loaded via doGet)
      kl: ''  // this is equal to the old myKL and is the ID from the player's KL (loaded via doGet's call to fGetMyKlId)
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
  MYKL_ID_CELL_A1: 'F8', // A1 notation for the cell containing MyKL ID in the Data tab
  designerPassword: 'Ogluck',
  // Add other server-side constants if needed
};



// doGet //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Auto launches when Turbo web app is 1st called - it then loads index.html.
//            Populates dynamic sheet IDs (cs, kl) in gSrv.
//            All based on a specific csID passed in the URL.
// Inputs  -> e (Event Object): Google Apps Script web app event object containing parameters.
// Outputs -> (HtmlOutput): The HTML page to be served.
function doGet(e) {
    let csId = null; // Use let to allow modification in error handling
    let klId = null; // Use let for KL ID

    try {
        // --- 1. Get and Validate CS ID ---
        csId = e?.parameter?.csID; // Use optional chaining
        if (!csId || typeof csId !== 'string') {
            console.error("doGet Error: csID parameter not provided or invalid in URL.");
            return HtmlService.createHtmlOutput("❌ csID parameter missing or invalid in URL.");
        }
        Logger.log(`doGet: Received CS ID: ${csId}`);

        // --- 2. Populate gSrv with CS ID ---
        // Note: Modifying properties of a top-level const (gSrv). This works per execution instance.
        gSrv.ids.sheets.cs = csId;
        Logger.log(`doGet: Assigned gSrv.ids.sheets.cs = ${gSrv.ids.sheets.cs}`);

        // --- 3. Get and Populate KL ID ---
        // Call fGetMyKlId to retrieve the KL ID associated with this CS ID
        klId = fGetMyKlId(csId); // This function handles its own internal errors/throws
        gSrv.ids.sheets.kl = klId;
        Logger.log(`doGet: Assigned gSrv.ids.sheets.kl = ${gSrv.ids.sheets.kl}`);

        // --- 4. Serve HTML ---
        // Attempt to load and serve the main HTML template
        // Optional: Verify csID is valid before rendering (can be slow)
        // SpreadsheetApp.openById(csId);

        const template = HtmlService.createTemplateFromFile("index");
        template.csID = csId; // Pass csId to the template client-side (remains gIndexCSID there)

        // Evaluate and return the HTML
        return template.evaluate()
            .setTitle("Turbo")
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    } catch (error) {
        // Log detailed error and return user-friendly error message
        const errorContext = `CS ID: ${csId || 'Unknown'}, KL ID Fetch Attempted: ${klId !== null || 'No (failed before KL fetch)'}`;
        console.error(`doGet Error (${errorContext}): ${error.message}${error.stack ? '\n' + error.stack : ''}`);
        // Provide specific feedback based on whether the error likely came from fGetMyKlId or elsewhere
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




// ==========================================================================
// === Desinger Password  (END of Global Variables & doGet)===
// ==========================================================================



// fValidateDesignerPassword /////////////////////////////////////////////////
// Purpose -> Validates a password attempt against the stored designer password.
// Inputs  -> passwordAttempt (String): The password entered by the user.
// Outputs -> (Boolean): True if the password matches, false otherwise.
function fValidateDesignerPassword(passwordAttempt) {
    const funcName = "fValidateDesignerPassword";
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
} // END fValidateDesignerPassword



// ==========================================================================
// === Google Sheet/Doc Helper Functions  (END of Desinger Password)===
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



// fServerIndicesToA1 //////////////////////////////////////////////////////////
// Purpose -> Converts 0-based row/column indices to standard A1 notation string
//            on the server-side.
// Inputs  -> r1, c1, r2, c2 (Number): 0-based start/end row and column indices.
// Outputs -> (String | null): A1 notation string (e.g., "C5:F10"), or null on invalid input.
function fServerIndicesToA1(r1, c1, r2, c2) {
    // Validate inputs are non-negative numbers
    if ([r1, c1, r2, c2].some(idx => typeof idx !== 'number' || idx < 0 || isNaN(idx))) {
        console.error(`fServerIndicesToA1: Invalid indices provided (${r1},${c1},${r2},${c2})`);
        return null; // Return null for invalid indices
    }

    // Convert column indices to A1 letters using the existing helper
    const startColA1 = fColToA1(c1);
    const endColA1 = fColToA1(c2);

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
} // END fServerIndicesToA1




// fServerBuildTagMaps /////////////////////////////////////////////////////////
// Purpose -> Builds temporary rowTag and colTag maps from a 2D data array read
//            from a sheet (Row 0 for colTags, Col 0 for rowTags).
// Inputs  -> fullData (Array[][]): The 2D array of data from the sheet.
// Outputs -> (Object): An object { rowTag: {}, colTag: {} } containing the maps.
function fServerBuildTagMaps(fullData) {
    const rowTagMap = {};
    const colTagMap = {};

    // Validate data array structure
    if (!Array.isArray(fullData) || fullData.length === 0) {
        console.warn("fServerBuildTagMaps: Input data array is empty or invalid.");
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
                    //   console.warn(`fServerBuildTagMaps: Overwriting row tag "${trimmedTag}" (Old: ${rowTagMap[trimmedTag]}, New: ${r})`);
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
                        //   console.warn(`fServerBuildTagMaps: Overwriting col tag "${trimmedTag}" (Old: ${colTagMap[trimmedTag]}, New: ${c})`);
                        // }
                        colTagMap[trimmedTag] = c;
                    }
                });
            }
        }
    } else {
        console.warn("fServerBuildTagMaps: Header row (row 0) is missing or invalid. Cannot build column tags.");
    }

    return { rowTag: rowTagMap, colTag: colTagMap };
} // END fServerBuildTagMaps




// fServerResolveTag ///////////////////////////////////////////////////////////
// Purpose -> Resolves a single row or column tag/index using the provided tag maps.
// Inputs  -> tagOrIndex (String | Number): The tag string or 0-based index.
//         -> tagMap (Object): The corresponding tag map (rowTag or colTag).
//         -> type (String): 'row' or 'col' for logging purposes.
// Outputs -> (Number): The resolved 0-based index, or NaN if resolution fails.
function fServerResolveTag(tagOrIndex, tagMap, type = 'unknown') {
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
            console.warn(`fServerResolveTag: Could not resolve ${type} tag "${trimmedTag}"`);
            return NaN;
        }
    } else {
        // Input is invalid type or empty
        console.warn(`fServerResolveTag: Invalid input provided for ${type} resolution:`, tagOrIndex);
        return NaN;
    }
} // END fServerResolveTag




// ==========================================================================
// === Get Game Sheet data/styles                          (END of Google Sheet/Doc Helper Functions) ===
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




// fGetServerSheetData /////////////////////////////////////////////////////////
// Purpose -> Reads data from a specified range in a given Google Sheet,
//            accepting a sheet key OR a direct fileId, and a rangeObject.
//            Reads the full sheet once for tag mapping. Extracts the requested
//            data slice. Builds row/column tag maps with indices *relative*
//            to the extracted slice.
//            Returns an object containing the data array and the relative tag maps.
// Inputs  -> sheetKeyOrId (String): Key from gSrv.ids.sheets OR a direct Sheet fileId.
//         -> sheetName (String): The name of the specific sheet (tab) to read from.
//         -> rangeObject (Object): {r1, c1, r2, c2} using tags or 0-based indices.
// Outputs -> (Object): { data: Array[][]|Array[]|Any|null, colTags: Object, rowTags: Object }
//            'data' contains values from the range, formatted correctly (single, 1D, or 2D).
//            colTags/rowTags contain mappings with indices relative to the 'data' array structure.
//            Returns null for 'data' on critical errors.
// Throws  -> (Error): If inputs are invalid, sheet/range access fails, etc.
function fGetServerSheetData(sheetKeyOrId, sheetName, rangeObject) {
    Logger.log(`fGetServerSheetData: Request received. Key/ID: "${sheetKeyOrId}", Sheet: "${sheetName}", Range: ${JSON.stringify(rangeObject)}`);
    let fileId = null;
    let identifiedBy = '';
    let absoluteRowTagMap = {}; // Holds tags with absolute sheet indices initially
    let absoluteColTagMap = {}; // Holds tags with absolute sheet indices initially
    let relativeRowTagMap = {}; // Holds tags with indices relative to extracted data
    let relativeColTagMap = {}; // Holds tags with indices relative to extracted data

    // --- 1. Resolve File ID ---
    if (!sheetKeyOrId || typeof sheetKeyOrId !== 'string') {
        const errorMsg = `Invalid or missing sheetKeyOrId parameter.`;
        console.error(`fGetServerSheetData Error: ${errorMsg}`);
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
            console.error(`fGetServerSheetData Error: ${errorMsg}`);
            throw new Error(`getServerSheetData: ${errorMsg}`);
        }
        fileId = gSrv.ids.sheets[keyLower];
        identifiedBy = `Key: ${sheetKeyOrId} -> ID: ${fileId}`;
        Logger.log(`   -> Interpreted first argument as Key: "${sheetKeyOrId}", resolved to ID: ${fileId}`);
    }

    // --- 2. Validate Other Inputs ---
     if (!fileId) { throw new Error(`getServerSheetData: Could not determine File ID.`); }
    if (!sheetName || typeof sheetName !== 'string') { throw new Error(`getServerSheetData: Invalid or missing sheetName.`); }
    if (!rangeObject || typeof rangeObject !== 'object' || rangeObject.r1 === undefined || rangeObject.c1 === undefined || rangeObject.r2 === undefined || rangeObject.c2 === undefined) {
        throw new Error(`getServerSheetData: Invalid or incomplete rangeObject: ${JSON.stringify(rangeObject)}.`);
    }

    try {
        // --- 3. Open Sheet & Read Full Data ---
        const ss = SpreadsheetApp.openById(fileId);
        const sh = ss.getSheetByName(sheetName);
        if (!sh) { throw new Error(`Sheet named "${sheetName}" not found in Sheet (${identifiedBy}).`); }
        const fullData = sh.getDataRange().getValues();
        Logger.log(`fGetServerSheetData: Read ${fullData.length}x${fullData[0]?.length || 0} cells from "${sheetName}".`);

        // --- 4. Build *Absolute* Tag Maps ---
        const absoluteTagMaps = fServerBuildTagMaps(fullData); // Gets maps with absolute indices
        absoluteRowTagMap = absoluteTagMaps.rowTag;
        absoluteColTagMap = absoluteTagMaps.colTag;
        // Logger.log(`fGetServerSheetData: Built absolute tag maps. Rows: ${Object.keys(absoluteRowTagMap).length}, Cols: ${Object.keys(absoluteColTagMap).length}`);

        // --- Handle Empty Sheet Case ---
        if (fullData.length === 0 || fullData[0]?.length === 0) {
            console.warn(`fGetServerSheetData: Sheet "${sheetName}" (${identifiedBy}) appears empty.`);
            // Return empty data and empty relative maps
            return { data: [[]], colTags: {}, rowTags: {} };
        }

        // --- 5. Resolve Input Range Object Tags to *Absolute* Indices ---
        const r1_abs = fServerResolveTag(rangeObject.r1, absoluteRowTagMap, 'row');
        const c1_abs = fServerResolveTag(rangeObject.c1, absoluteColTagMap, 'col');
        const r2_abs = fServerResolveTag(rangeObject.r2, absoluteRowTagMap, 'row');
        const c2_abs = fServerResolveTag(rangeObject.c2, absoluteColTagMap, 'col');
        if ([r1_abs, c1_abs, r2_abs, c2_abs].some(isNaN)) {
            let failedTags = [];
            if (isNaN(r1_abs)) failedTags.push(`r1: ${rangeObject.r1}`);
            if (isNaN(c1_abs)) failedTags.push(`c1: ${rangeObject.c1}`);
            if (isNaN(r2_abs)) failedTags.push(`r2: ${rangeObject.r2}`);
            if (isNaN(c2_abs)) failedTags.push(`c2: ${rangeObject.c2}`);
            throw new Error(`Could not resolve absolute tags/indices in rangeObject: ${JSON.stringify(rangeObject)}. Failed Tags: ${failedTags.join(', ')}`);
        }
        Logger.log(`fGetServerSheetData: Resolved range to absolute indices: r1=${r1_abs}, c1=${c1_abs}, r2=${r2_abs}, c2=${c2_abs}`);

        // --- 6. Define Extraction Boundaries ---
        const rStart = Math.min(r1_abs, r2_abs);
        const rEnd = Math.max(r1_abs, r2_abs);
        const cStart = Math.min(c1_abs, c2_abs);
        const cEnd = Math.max(c1_abs, c2_abs);

        // --- 7. Extract Data Slice ---
        if (rStart >= fullData.length || cStart >= (fullData[0]?.length || 0)) {
             console.warn(`fGetServerSheetData: Resolved range start [${rStart}, ${cStart}] is outside the bounds of the sheet data [${fullData.length}, ${fullData[0]?.length||0}]. Returning empty data.`);
             return { data: [[]], colTags: {}, rowTags: {} }; // Return empty maps too
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
            // Check if the absolute index falls within the columns of the extracted slice
            if (absoluteIndex >= cStart && absoluteIndex <= cEnd) {
                const relativeIndex = absoluteIndex - cStart; // Calculate relative index
                relativeColTagMap[tag] = relativeIndex;
            }
        }
        // Adjust Row Tags
        for (const tag in absoluteRowTagMap) {
            const absoluteIndex = absoluteRowTagMap[tag];
            // Check if the absolute index falls within the rows of the extracted slice
            if (absoluteIndex >= rStart && absoluteIndex <= rEnd) {
                const relativeIndex = absoluteIndex - rStart; // Calculate relative index
                relativeRowTagMap[tag] = relativeIndex;
            }
        }
        Logger.log(`fGetServerSheetData: Built relative tag maps. Rel Rows: ${Object.keys(relativeRowTagMap).length}, Rel Cols: ${Object.keys(relativeColTagMap).length}`);

        // --- 9. Format Return Data ---
        let returnData;
        if (extractedRows === 1 && extractedCols === 1) {
            Logger.log(`fGetServerSheetData: Formatting data as single value.`);
            returnData = extractedData[0][0];
        } else if (extractedRows === 1) {
            Logger.log(`fGetServerSheetData: Formatting data as 1D array (single row).`);
            returnData = extractedData[0];
        } else if (extractedCols === 1 && extractedData.every(row => row.length === 1)) {
             Logger.log(`fGetServerSheetData: Formatting data as 1D array (single column).`);
             returnData = extractedData.map(row => row[0]);
        } else {
            Logger.log(`fGetServerSheetData: Formatting data as 2D array (${extractedRows}x${extractedCols}).`);
            returnData = extractedData;
        }

        // --- 10. Return Final Object ---
        return { data: returnData, colTags: relativeColTagMap, rowTags: relativeRowTagMap };

    } catch (e) {
        console.error(`Error in fGetServerSheetData for Input "${sheetKeyOrId}", Sheet "${sheetName}", Range ${JSON.stringify(rangeObject)}: ${e.message}\nStack: ${e.stack}`);
        throw new Error(`Server error processing sheet data: ${e.message || e}`);
    }

} // END fGetServerSheetData




// fSetServerSheetData /////////////////////////////////////////////////////////
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
function fSetServerSheetData(sheetKey, sheetName, rangeObject, valueOrValues) {
    Logger.log(`fSetServerSheetData: Requesting write to key "${sheetKey}", sheet "${sheetName}", range: ${JSON.stringify(rangeObject)}, type: ${Array.isArray(valueOrValues) ? (Array.isArray(valueOrValues[0]) ? '2D Array' : '1D Array') : typeof valueOrValues}`);

    // === 1. Validate Inputs (Sheet Key, Name, Range Object) ===
    if (!sheetKey || typeof sheetKey !== 'string' || !gSrv.ids.sheets[sheetKey.toLowerCase()]) {
        const errorMsg = `Invalid or missing sheetKey: "${sheetKey}". Not found in gSrv.ids.sheets.`;
        console.error(`fSetServerSheetData Error: ${errorMsg}`);
        throw new Error(`setServerSheetData: ${errorMsg}`);
    }
    const fileId = gSrv.ids.sheets[sheetKey.toLowerCase()];
    if (!sheetName || typeof sheetName !== 'string') {
        const errorMsg = `Invalid or missing sheetName: "${sheetName}".`;
        console.error(`fSetServerSheetData Error: ${errorMsg}`);
        throw new Error(`setServerSheetData: ${errorMsg}`);
    }
    if (!rangeObject || typeof rangeObject !== 'object' ||
        rangeObject.r1 === undefined || rangeObject.c1 === undefined ||
        rangeObject.r2 === undefined || rangeObject.c2 === undefined) {
        const errorMsg = `Invalid or incomplete rangeObject: ${JSON.stringify(rangeObject)}.`;
        console.error(`fSetServerSheetData Error: ${errorMsg}`);
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
           console.warn(`fSetServerSheetData: Sheet "${sheetName}" appears empty. Tag resolution might fail if tags are used.`);
        }

        // === 3. Build Tag Maps (from in-memory data) ===
        const { rowTag, colTag } = fServerBuildTagMaps(fullData);

        // === 4. Resolve Input Range Object Tags ===
        const r1 = fServerResolveTag(rangeObject.r1, rowTag, 'row');
        const c1 = fServerResolveTag(rangeObject.c1, colTag, 'col');
        const r2 = fServerResolveTag(rangeObject.r2, rowTag, 'row');
        const c2 = fServerResolveTag(rangeObject.c2, colTag, 'col');
        if ([r1, c1, r2, c2].some(isNaN)) {
            throw new Error(`Could not resolve one or more tags/indices in rangeObject: ${JSON.stringify(rangeObject)}`);
        }
        Logger.log(`fSetServerSheetData: Resolved range to indices: r1=${r1}, c1=${c1}, r2=${r2}, c2=${c2}`);

        // === 5. Get Target Range & Dimensions ===
        const rangeNumRows = Math.abs(r1 - r2) + 1;
        const rangeNumCols = Math.abs(c1 - c2) + 1;
        const a1Notation = fServerIndicesToA1(r1, c1, r2, c2);
        if (!a1Notation) {
             throw new Error(`Could not convert resolved indices [${r1},${c1},${r2},${c2}] to A1 notation.`);
        }
        const targetRange = sh.getRange(a1Notation);
        if (!targetRange) {
             throw new Error(`Could not get target range "${a1Notation}" in sheet "${sheetName}".`);
        }
        Logger.log(`fSetServerSheetData: Target A1 notation: ${a1Notation} (${rangeNumRows}x${rangeNumCols})`);

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

        Logger.log(`fSetServerSheetData: Successfully wrote data to range ${a1Notation}.`);
        return true; // Indicate success

    } catch (e) {
        // Log detailed error and re-throw a user-friendly error
        console.error(`Error in fSetServerSheetData for Key ${sheetKey}, Sheet "${sheetName}", Range ${JSON.stringify(rangeObject)}: ${e.message}\nStack: ${e.stack}`);
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




// fGetCharacterHeaderData /////////////////////////////////////////////////////
// Purpose -> Reads specific static data (Race/Class, Level, Player/Char Names, Slot)
//            from the player's Character Sheet ('RaceClass' tab).
// Inputs  -> csId (String): The ID of the player's Character Sheet.
// Outputs -> (Object): { raceClass, level, playerName, charName, slotNum } on success.
//            'slotNum' will be the cleaned tag (e.g., "Slot3") or null if validation fails.
// Throws  -> (Error): If sheet/tab/tags not found or critical values are missing.
function fGetCharacterHeaderData(csId) {
    const funcName = "fGetCharacterHeaderData";
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
        const { rowTag, colTag } = fServerBuildTagMaps(fullData);

        // --- 4. Resolve Target Cell Indices ---
        const colValIndex = fServerResolveTag('Val', colTag, 'col');
        const rowRaceClassIndex = fServerResolveTag('RC', rowTag, 'row');
        const rowLevelIndex = fServerResolveTag('level', rowTag, 'row');
        const rowPlayerNameIndex = fServerResolveTag('playerName', rowTag, 'row');
        const rowCharNameIndex = fServerResolveTag('charName', rowTag, 'row');
        const rowSlotIndex = fServerResolveTag('slot', rowTag, 'row');

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

} // END fGetCharacterHeaderData



// fSetLogAndHeaderData ////////////////////////////////////////////////////////
// Purpose -> Writes bundled log and header data to target sheets (DB/GMScreen
//            and PS/PartyLog) using relative offsets from a base cell identified by
//            'Log' row tag and a dynamic slot column tag (e.g., 'Slot3') passed in the bundle.
// Inputs  -> dataBundle (Object): { log, vit, nish, url, raceClass, level, playerChar, slotNum }
// Outputs -> (Boolean): True if both writes succeeded, false otherwise.
// Throws  -> (Error): If critical errors occur (e.g., opening sheets, invalid bundle).
function fSetLogAndHeaderData(dataBundle) {
    const funcName = "fSetLogAndHeaderData";
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
            const { rowTag, colTag } = fServerBuildTagMaps(fullData);

            // Resolve the *base cell* using fixed Row ('Log') and dynamic Column (baseCellTagC)
            const baseRowIndex = fServerResolveTag(baseCellTagR, rowTag, 'row');
            const baseColIndex = fServerResolveTag(baseCellTagC, colTag, 'col'); // Use dynamic tag

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

} // END fSetLogAndHeaderData



// ==========================================================================
// === Firestore                       (End of Read/Write to Google Sheets/Docs ) ===
// ==========================================================================



// fGetFirestoreInstance ///////////////////////////////////////////////////////
// Purpose -> Initializes and returns an authenticated Firestore instance using
//            credentials stored in PropertiesService. [Processes Key String]
// Inputs  -> None.
// Outputs -> (Object | null): Authenticated Firestore instance or null on error.
function fGetFirestoreInstance() {
    const funcName = "fGetFirestoreInstance";
    Logger.log(`${funcName}: Attempting to initialize Firestore...`);
    let clientEmail, privateKeyRaw, projectId, processedKey; // Declare vars
    try {
        // Retrieve credentials from Script Properties
        const scriptProperties = PropertiesService.getScriptProperties();
        clientEmail = scriptProperties.getProperty('FIRESTORE_CLIENT_EMAIL');
        privateKeyRaw = scriptProperties.getProperty('FIRESTORE_PRIVATE_KEY'); // Get the raw string
        projectId = scriptProperties.getProperty('FIRESTORE_PROJECT_ID');

        // Validate credentials & Log Status
        if (!clientEmail) Logger.log(`   -> ${funcName} Error: FIRESTORE_CLIENT_EMAIL not found or empty.`);
        if (!privateKeyRaw) Logger.log(`   -> ${funcName} Error: FIRESTORE_PRIVATE_KEY not found or empty.`);
        if (!projectId) Logger.log(`   -> ${funcName} Error: FIRESTORE_PROJECT_ID not found or empty.`);

        if (!clientEmail || !privateKeyRaw || !projectId) {
            console.error(`${funcName} Error: Missing Firestore credentials in Script Properties.`);
            return null; // Exit if any credential is fundamentally missing
        }

        // --- Process the Private Key String ---
        Logger.log(`   -> Raw Private Key from Props (Snippet): ${privateKeyRaw.substring(0, 20)}...${privateKeyRaw.substring(privateKeyRaw.length - 20)}`);
        processedKey = privateKeyRaw;
        // 1. Remove surrounding quotes if present (handle copy-paste variations)
        if (processedKey.startsWith('"') && processedKey.endsWith('"')) {
            processedKey = processedKey.substring(1, processedKey.length - 1);
            Logger.log(`   -> Removed surrounding quotes from key.`);
        }
        // 2. Replace literal "\\n" sequences with actual newline characters "\n"
        //    (Important: Use double backslash in replaceAll to match the literal sequence)
        processedKey = processedKey.replaceAll('\\n', '\n');
        Logger.log(`   -> Replaced \\\\n with \\n in key.`);
        Logger.log(`   -> Processed Private Key (Snippet): ${processedKey.substring(0, 40)}...${processedKey.substring(processedKey.length - 40)}`);
        // --- End Key Processing ---

        // Log other retrieved values
        Logger.log(`   -> Project ID: ${projectId}`);
        Logger.log(`   -> Client Email: ${clientEmail}`);


        // Initialize Firestore using the library and *processed* key
        Logger.log(`   -> Calling FirestoreApp.getFirestore with processed credentials...`);
        const firestore = FirestoreApp.getFirestore(clientEmail, processedKey, projectId); // Use processedKey here

        // Check if firestore object was created
        if (!firestore) {
           Logger.log(`   -> ${funcName} Error: FirestoreApp.getFirestore returned null/undefined after processing key.`);
           console.error(`${funcName} Error: FirestoreApp.getFirestore failed to return an instance.`);
           return null;
        }

        Logger.log(`${funcName}: Firestore instance appears initialized successfully for project ${projectId}.`);
        return firestore;

    } catch (e) {
        // Catch errors during property retrieval or initialization
        console.error(`Error caught in ${funcName}: ${e.message}\nStack: ${e.stack}`);
        Logger.log(`   -> ❌ Exception during Firestore initialization: ${e.message}`);
        // Log potentially relevant details if available before the error
        Logger.log(`   -> Details at time of error: ProjectID=${projectId || 'N/A'}, Email=${clientEmail || 'N/A'}, Key Raw loaded=${!!privateKeyRaw}, Key Processed Snippet=${processedKey ? processedKey.substring(0,20) + '...' : 'N/A'}`);
        return null;
    }
} // END fGetFirestoreInstance





// fSaveTextDataToFirestore ////////////////////////////////////////////////////
// Purpose -> Saves grid text data and character metadata to a Firestore document
//            in the 'GameTextData' collection. Document ID is GSID_<csId>.
//            Processes the incoming 2D array into an Array of Row Objects format
//            [ {"row0": [...]}, {"row1": [...]}, ... ] for Firestore compatibility.
//            Uses updateDocument which creates if not exists.
// Inputs  -> csId (String): The Character Sheet ID (used for document path).
//         -> fullArrData (Array[][]): The complete gUI.arr from the client.
//         -> charInfo (Object): Character metadata { slotNum, raceClass, playerName, etc. }.
// Outputs -> (Object): { success: Boolean, message?: String }
function fSaveTextDataToFirestore(csId, fullArrData, charInfo) {
    const funcName = "fSaveTextDataToFirestore";
    Logger.log(`${funcName}: Saving array-of-row-objects data for CS ID: ${csId}...`);

    // === 1. Validate Inputs ===
    if (!csId || typeof csId !== 'string') {
        return { success: false, message: "Invalid Character Sheet ID provided." };
    }
    // Validate fullArrData is a 2D array (basic check)
    if (!Array.isArray(fullArrData) || (fullArrData.length > 0 && !Array.isArray(fullArrData[0]))) {
         return { success: false, message: "Invalid fullArrData provided (must be 2D array)." };
    }
     // Validate characterInfo and specifically playerName needed for document ID (still needed for metadata part)
     if (!charInfo || typeof charInfo !== 'object' || !charInfo.playerName || typeof charInfo.playerName !== 'string' || charInfo.playerName.trim() === '') {
        const msg = "Invalid or missing characterInfo.playerName for Firestore document metadata.";
        Logger.log(`   -> ${funcName} Error: ${msg}`);
        return { success: false, message: msg };
    }
    Logger.log(`   -> Received fullArrData with dimensions: ${fullArrData.length}x${fullArrData[0]?.length || 0}`);
    Logger.log(`   -> Received charInfo: ${JSON.stringify(charInfo)}`);

    // === 2. Process Array into Array of Row Objects ===
    const arrayOfRowObjects = [];
    const numRows = fullArrData.length;
    Logger.log(`   -> Processing ${numRows} rows into array-of-row-objects format...`);
    for (let r = 0; r < numRows; r++) {
        const rowData = fullArrData[r] || []; // Get the row array or an empty array if row doesn't exist
        const rowKey = `row${r}`;
        const rowObject = {};
        rowObject[rowKey] = rowData; // Create object like {"row0": [val, val, ...]}
        arrayOfRowObjects.push(rowObject);
    }
    Logger.log(`   -> Created arrayOfRowObjects with ${arrayOfRowObjects.length} entries.`);


    // === 3. Get Firestore Instance ===
    const firestore = fGetFirestoreInstance();
    if (!firestore) {
        const msg = "Failed to initialize Firestore instance (check previous logs).";
        Logger.log(`${funcName} Error: ${msg}`);
        return { success: false, message: "Server configuration error (Firestore)." };
    }
    Logger.log(`   -> Firestore instance obtained.`);

    // === 4. Define Path and Data ===
    const documentId = `GSID_${csId}`;
    const collectionName = 'GameTextData';
    const documentPath = `${collectionName}/${documentId}`;

    const dataToSave = {
        gUIcharacterInfo: charInfo,       // Store character metadata
        gUIarr: arrayOfRowObjects,        // Stores gUI.arr as an array of row objects
        _lastUpdated: new Date()           // Add a server-side timestamp
    };
    Logger.log(`   -> Target Firestore Path: ${documentPath}`);
    // Logger.log(`   -> Data to Save (Snippet): ${JSON.stringify(dataToSave).substring(0, 200)}...`);

    // === 5. Save to Firestore ===
    try {
        Logger.log(`   -> Calling firestore.updateDocument...`);
        // Use updateDocument with mask=false to overwrite the entire document or create if needed.
        const writeResult = firestore.updateDocument(documentPath, dataToSave, false);
        Logger.log(`   -> Firestore write result: ${JSON.stringify(writeResult)}`); // Log raw result

        // Basic check on writeResult (can be adapted based on library specifics)
        if (writeResult && typeof writeResult === 'object') {
             Logger.log(`   -> ✅ Successfully saved array-of-row-objects data for ${csId} to ${documentPath}.`);
             return { success: true };
        } else {
             const msg = `Firestore updateDocument returned unexpected result.`;
             Logger.log(`   -> ❌ ${funcName} Error: ${msg}`);
             return { success: false, message: `Firestore write returned unexpected result: ${JSON.stringify(writeResult)}` };
        }

    } catch (e) {
        console.error(`Exception caught in ${funcName} saving to path ${documentPath}: ${e.message}\nStack: ${e.stack}`);
        Logger.log(`   -> ❌ Exception during Firestore save for ${csId}: ${e.message}`);
        // Check if the error is the nested array error again
        const nestedArrayError = "Nested arrays are not allowed";
        const safeErrorMessage = e.message.includes("permission") ? "Permission denied."
                               : e.message.includes(nestedArrayError) ? nestedArrayError + " (even with Array of Row Objects format)."
                               : "Server error during Firestore save.";
        return { success: false, message: safeErrorMessage };
    }
} // END fSaveTextDataToFirestore




// fGetTextDataFromFirestore ///////////////////////////////////////////////////
// Purpose -> Fetches the character info and grid data stored in Firestore for a given CS ID.
//            Converts Firestore types to standard JS types. Unpacks gUIarr server-side.
// Inputs  -> csId (String): The Character Sheet ID (used for document path).
// Outputs -> (Object): { success: Boolean, data?: { characterInfo: Object, arrData: Array[][] }, message?: String }
//            'arrData' is the standard 2D array if success is true.
function fGetTextDataFromFirestore(csId) {
    const funcName = "fGetTextDataFromFirestore";
    Logger.log(`${funcName}: Fetching data for CS ID: ${csId}...`);

    // === 1. Validate Input ===
    if (!csId || typeof csId !== 'string') {
        return { success: false, message: "Invalid Character Sheet ID provided." };
    }

    // === 2. Get Firestore Instance ===
    const firestore = fGetFirestoreInstance();
    if (!firestore) {
        const msg = "Failed to initialize Firestore instance (check previous logs).";
        Logger.log(`${funcName} Error: ${msg}`);
        return { success: false, message: "Server configuration error (Firestore)." };
    }
    Logger.log(`   -> Firestore instance obtained.`);

    // === 3. Define Path and Fetch Data ===
    const documentId = `GSID_${csId}`;
    const collectionName = 'GameTextData';
    const documentPath = `${collectionName}/${documentId}`;
    Logger.log(`   -> Target Firestore Path: ${documentPath}`);

    try {
        const doc = firestore.getDocument(documentPath);
        Logger.log(`   -> Raw doc object fetched: ${JSON.stringify(doc).substring(0, 500)}...`);

        // === 4. Validate Fetched Document (using fields check) ===
        if (!doc || !doc.fields) {
            const msg = `Document object retrieved, but 'fields' property is missing or document is empty. Path: ${documentPath}`;
            Logger.log(`   -> ${funcName} Warning: ${msg}`);
            Logger.log(`   -> Raw doc object for check: ${JSON.stringify(doc)}`);
            return { success: false, message: "No saved data found or data format error." };
        }
        const fetchedData = doc.fields;
        Logger.log(`   -> Document fields appear to exist. Proceeding to convert types...`);

        // === 5. Convert Firestore Types to JavaScript Types ===
        if (!fetchedData.gUIcharacterInfo || !fetchedData.gUIarr) {
            const msg = `Firestore document at ${documentPath} is missing expected top-level fields (gUIcharacterInfo or gUIarr) before type conversion.`;
            Logger.log(`   -> ${funcName} Error: ${msg}`);
            Logger.log(`   -> Found keys: ${Object.keys(fetchedData).join(', ')}`);
            return { success: false, message: "Saved data format is invalid or incomplete." };
        }

        const characterInfoRaw = fetchedData.gUIcharacterInfo;
        const arrDataRaw = fetchedData.gUIarr; // This is the array-of-row-objects structure
        Logger.log(`   -> Converting characterInfo...`);
        const characterInfo = fConvertFirestoreTypesToJS(characterInfoRaw);
        Logger.log(`   -> Converting arrData (array-of-objects)...`);
        const arrDataConverted = fConvertFirestoreTypesToJS(arrDataRaw); // Still array-of-objects

        // === 6. Validate Converted JavaScript Types ===
        if (typeof characterInfo !== 'object' || characterInfo === null || Array.isArray(characterInfo)) {
            const msg = `Invalid data type for characterInfo after conversion. Expected object, got ${typeof characterInfo}. Value: ${JSON.stringify(characterInfo)}`;
            Logger.log(`   -> ${funcName} Error: ${msg}`);
            return { success: false, message: "Invalid character info format retrieved." };
        }
        if (!Array.isArray(arrDataConverted)) { // Validate it's an array before unpacking
            const msg = `Invalid data type for arrData after conversion. Expected array (of objects), got ${typeof arrDataConverted}. Value: ${JSON.stringify(arrDataConverted).substring(0,200)}...`;
            Logger.log(`   -> ${funcName} Error: ${msg}`);
            return { success: false, message: "Invalid grid data format retrieved (expected array)." };
        }

        // === 7. Unpack Firestore Array ===
        Logger.log(`   -> Unpacking Firestore array data server-side...`);
        const unpackedArr = fUnpackFirestoreArray_Server(arrDataConverted);
        Logger.log(`   -> Unpacked into ${unpackedArr.length}x${unpackedArr[0]?.length || 0} array.`);


        // === 8. Return Successfully Converted and Unpacked Data ===
        Logger.log(`   -> ✅ Successfully fetched, converted, and unpacked data from ${documentPath}.`);
        return {
            success: true,
            data: {
                characterInfo: characterInfo, // Return plain JS object
                arrData: unpackedArr          // Return standard 2D array
            }
        };

    } catch (e) {
        console.error(`Exception caught in ${funcName} fetching/converting from path ${documentPath}: ${e.message}\nStack: ${e.stack}`);
        Logger.log(`   -> ❌ Exception during Firestore fetch/convert for ${csId}: ${e.message}`);
        const safeErrorMessage = e.message?.includes("permission") ? "Permission denied."
                               : e.message?.includes("NOT_FOUND") ? "No saved Firestore data found for this character." // More specific not found
                               : "Server error during Firestore operation.";
        return { success: false, message: safeErrorMessage };
    }
} // END fGetTextDataFromFirestore




// fConvertFirestoreTypesToJS //////////////////////////////////////////////////
// Purpose -> Recursively converts Firestore's typed value objects (mapValue,
//            arrayValue, stringValue, etc.) into standard JavaScript types
//            (objects, arrays, strings, numbers, booleans).
// Inputs  -> firestoreValue (Object): A value object from Firestore (e.g.,
//            doc.fields.someProperty or an element within an arrayValue/mapValue).
// Outputs -> (Any): The corresponding standard JavaScript value or type.
function fConvertFirestoreTypesToJS(firestoreValue) {
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
    return firestoreValue.arrayValue.values.map(element => fConvertFirestoreTypesToJS(element));
  }

  // Check for map/object type
  if (firestoreValue.mapValue && firestoreValue.mapValue.fields) {
    // It's an object/map, recursively convert its properties
    const jsObject = {};
    for (const key in firestoreValue.mapValue.fields) {
      jsObject[key] = fConvertFirestoreTypesToJS(firestoreValue.mapValue.fields[key]);
    }
    return jsObject;
  }

  // If it's none of the known Firestore types (e.g., already a JS primitive passed in), return as is.
  // Or potentially log a warning if an unexpected structure is encountered.
  Logger.log(`fConvertFirestoreTypesToJS: Encountered unexpected value structure: ${JSON.stringify(firestoreValue).substring(0,100)}... Returning as is.`);
  return firestoreValue;
} // END fConvertFirestoreTypesToJS



// fUnpackFirestoreArray_Server //////////////////////////////////////////////
// Purpose -> Converts the Firestore array-of-row-objects format into a standard
//            2D JavaScript array server-side. Handles potential sparse arrays.
// Inputs  -> firestoreArr (Array): Array from Firestore after type conversion,
//            e.g., [ {"row0":[...]}, {"row1":[...]}, ... ]
// Outputs -> (Array[][]): A standard 2D JavaScript array. Returns empty array on error.
function fUnpackFirestoreArray_Server(firestoreArr) {
    const funcName = "fUnpackFirestoreArray_Server";
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
} // END fUnpackFirestoreArray_Server



// fCheckAndLoadFirestoreData //////////////////////////////////////////////////
// Purpose -> Checks Firestore for a document based on CS ID. If found, extracts
//            the gUIarr field, converts types, unpacks it into a 2D array,
//            and returns it. Handles document not found and errors gracefully.
// Inputs  -> csId (String): The Character Sheet ID.
// Outputs -> (Object): { success: Boolean, firestoreArr?: Array[][], message?: String }
//            'firestoreArr' is the standard 2D array if success is true.
function fCheckAndLoadFirestoreData(csId) {
    const funcName = "fCheckAndLoadFirestoreData";
    Logger.log(`${funcName}: Checking Firestore for data for CS ID: ${csId}...`);

    // === 1. Validate Input ===
    if (!csId || typeof csId !== 'string') {
        const msg = "Invalid Character Sheet ID provided.";
        Logger.log(`${funcName} Error: ${msg}`);
        // Note: Don't throw, return failure object for client handling
        return { success: false, message: msg };
    }

    // === 2. Get Firestore Instance ===
    const firestore = fGetFirestoreInstance();
    if (!firestore) {
        const msg = "Failed to initialize Firestore instance.";
        Logger.log(`${funcName} Error: ${msg}`);
        return { success: false, message: "Server configuration error (Firestore)." };
    }
    Logger.log(`   -> Firestore instance obtained.`);

    // === 3. Define Path and Fetch Data ===
    const documentId = `GSID_${csId}`;
    const collectionName = 'GameTextData';
    const documentPath = `${collectionName}/${documentId}`;
    Logger.log(`   -> Target Firestore Path: ${documentPath}`);

    try {
        // Attempt to get the document
        const doc = firestore.getDocument(documentPath);

        // === 4. Check if Document Exists & Has Data ===
        if (!doc || !doc.fields || !doc.fields.gUIarr) {
            const msg = `Document not found or missing 'gUIarr' field at path: ${documentPath}.`;
            Logger.log(`   -> ${funcName}: ${msg}`);
            return { success: false, message: "No saved Firestore data found for this character." };
        }
        Logger.log(`   -> Document found. Processing 'gUIarr' field...`);

        // === 5. Convert Firestore Types for gUIarr ===
        const arrDataRaw = doc.fields.gUIarr;
        const arrDataConverted = fConvertFirestoreTypesToJS(arrDataRaw);

        // === 6. Validate Converted Data (Should be Array of Objects) ===
        if (!Array.isArray(arrDataConverted)) {
            const msg = `Invalid data type for gUIarr after conversion. Expected array, got ${typeof arrDataConverted}. Path: ${documentPath}`;
            Logger.log(`   -> ${funcName} Error: ${msg}`);
            return { success: false, message: "Invalid grid data format retrieved from Firestore." };
        }

        // === 7. Unpack Array of Objects into 2D Array ===
        Logger.log(`   -> Unpacking Firestore array data server-side...`);
        const unpackedArr = fUnpackFirestoreArray_Server(arrDataConverted);
        Logger.log(`   -> Unpacked into ${unpackedArr.length}x${unpackedArr[0]?.length || 0} array.`);

        // === 8. Return Success with Unpacked Data ===
        Logger.log(`   -> ✅ Successfully fetched and unpacked Firestore data for ${csId}.`);
        return {
            success: true,
            firestoreArr: unpackedArr // Return the standard 2D array
        };

    } catch (e) {
        // Handle potential errors during Firestore operations (e.g., permissions)
        const safeErrorMessage = e.message?.includes("permission")
                               ? "Permission denied accessing Firestore."
                               : e.message?.includes("NOT_FOUND")
                               ? "No saved Firestore data found for this character." // More specific not found
                               : `Server error during Firestore read: ${e.message || e}`;

        console.error(`Exception caught in ${funcName} fetching from path ${documentPath}: ${e.message}\nStack: ${e.stack}`);
        Logger.log(`   -> ❌ Exception during Firestore read for ${csId}: ${safeErrorMessage}`);
        return { success: false, message: safeErrorMessage };
    }
} // END fCheckAndLoadFirestoreData



