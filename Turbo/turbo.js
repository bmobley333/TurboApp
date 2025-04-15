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
      cs: '',  
      kl: ''  // this is equal to the old myKL and is the ID from the player's KL
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

    Logger.log(`fGetMyKlId: Successfully retrieved MyKL ID: ${gSrv.ids.kl}`);
    console.log("🧪 KLID param:", gSrv.ids.kl);
    return myKlId.trim(); // Return the trimmed ID

  } catch (e) {
    // Log detailed error and re-throw a user-friendly error
    console.error(`Error in fGetMyKlId for CS ID ${myCsId}: ${e.message}\nStack: ${e.stack}`);
    // Include specific error message if available, otherwise generic
    throw new Error(`Server error getting MyKL ID: ${e.message || e}`);
  }

} // END fGetMyKlId




// ==========================================================================
// === Google Sheet/Doc Helper Functions  (END of Global Variables & doGet)===
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
//            Reads the full sheet once for tag mapping and data extraction.
// Inputs  -> sheetKeyOrId (String): Key from gSrv.ids.sheets OR a direct Sheet fileId.
//         -> sheetName (String): The name of the specific sheet (tab) to read from.
//         -> rangeObject (Object): {r1, c1, r2, c2} using tags or 0-based indices.
// Outputs -> (Array[][]|Array[]|Any|null): Values from the range, formatted correctly.
//            Returns null on critical errors.
// Throws  -> (Error): If inputs are invalid, sheet/range access fails, etc.
function fGetServerSheetData(sheetKeyOrId, sheetName, rangeObject) {
    Logger.log(`fGetServerSheetData: Request received. Key/ID: "${sheetKeyOrId}", Sheet: "${sheetName}", Range: ${JSON.stringify(rangeObject)}`);

    // === 1. Determine File ID ===
    let fileId = null;
    let identifiedBy = ''; // For logging

    // Basic validation of the first argument
    if (!sheetKeyOrId || typeof sheetKeyOrId !== 'string') {
        const errorMsg = `Invalid or missing sheetKeyOrId parameter.`;
        console.error(`fGetServerSheetData Error: ${errorMsg}`);
        throw new Error(`getServerSheetData: ${errorMsg}`);
    }

    // Check if sheetKeyOrId looks like a Sheet ID (heuristic: length > 30 and alphanumeric/_)
    const isLikelySheetId = sheetKeyOrId.length > 30 && /^[a-z0-9_-]+$/i.test(sheetKeyOrId);

    if (isLikelySheetId) {
        fileId = sheetKeyOrId;
        identifiedBy = `Direct ID: ${fileId}`;
        Logger.log(`   -> Interpreted first argument as Direct File ID: ${fileId}`);
    } else {
        // Treat as a sheetKey, look up in gSrv
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

    // === 2. Validate Other Inputs ===
    if (!fileId) { // Should be caught above, but double-check
         const errorMsg = `Could not determine File ID from input: "${sheetKeyOrId}".`;
         console.error(`fGetServerSheetData Error: ${errorMsg}`);
         throw new Error(`getServerSheetData: ${errorMsg}`);
    }
    if (!sheetName || typeof sheetName !== 'string') {
        const errorMsg = `Invalid or missing sheetName: "${sheetName}".`;
        console.error(`fGetServerSheetData Error: ${errorMsg}`);
        throw new Error(`getServerSheetData: ${errorMsg}`);
    }
    if (!rangeObject || typeof rangeObject !== 'object' ||
        rangeObject.r1 === undefined || rangeObject.c1 === undefined ||
        rangeObject.r2 === undefined || rangeObject.c2 === undefined) {
        const errorMsg = `Invalid or incomplete rangeObject: ${JSON.stringify(rangeObject)}.`;
        console.error(`fGetServerSheetData Error: ${errorMsg}`);
        throw new Error(`getServerSheetData: ${errorMsg}`);
    }

    try {
        // === 3. Open Sheet & Read Full Data ===
        const ss = SpreadsheetApp.openById(fileId);
        const sh = ss.getSheetByName(sheetName);
        if (!sh) {
            throw new Error(`Sheet named "${sheetName}" not found in Sheet (${identifiedBy}).`);
        }
        const fullData = sh.getDataRange().getValues();
        Logger.log(`fGetServerSheetData: Read ${fullData.length}x${fullData[0]?.length || 0} cells from "${sheetName}".`);

        if (fullData.length === 0 || fullData[0]?.length === 0) {
            console.warn(`fGetServerSheetData: Sheet "${sheetName}" (${identifiedBy}) appears empty.`);
            return [[]];
        }

        // === 4. Build Tag Maps ===
        const { rowTag, colTag } = fServerBuildTagMaps(fullData);

        // === 5. Resolve Input Range Object Tags ===
        const r1 = fServerResolveTag(rangeObject.r1, rowTag, 'row');
        const c1 = fServerResolveTag(rangeObject.c1, colTag, 'col');
        const r2 = fServerResolveTag(rangeObject.r2, rowTag, 'row');
        const c2 = fServerResolveTag(rangeObject.c2, colTag, 'col');

        if ([r1, c1, r2, c2].some(isNaN)) {
            throw new Error(`Could not resolve one or more tags/indices in rangeObject: ${JSON.stringify(rangeObject)}`);
        }
        Logger.log(`fGetServerSheetData: Resolved range to indices: r1=${r1}, c1=${c1}, r2=${r2}, c2=${c2}`);

        // === 6. Extract Data from In-Memory Array ===
        const rStart = Math.min(r1, r2);
        const rEnd = Math.max(r1, r2);
        const cStart = Math.min(c1, c2);
        const cEnd = Math.max(c1, c2);

        if (rStart >= fullData.length || cStart >= (fullData[0]?.length || 0)) {
             console.warn(`fGetServerSheetData: Resolved range start [${rStart}, ${cStart}] is outside the bounds of the sheet data [${fullData.length}, ${fullData[0]?.length||0}]. Returning empty.`);
             return [[]];
        }

        const extractedData = fullData
            .slice(rStart, rEnd + 1)
            .map(row => row.slice(cStart, cEnd + 1));

        // === 7. Format Return Value ===
        if (extractedData.length === 1 && extractedData[0].length === 1) {
            Logger.log(`fGetServerSheetData: Returning single value.`);
            return extractedData[0][0];
        } else if (extractedData.length === 1) {
            Logger.log(`fGetServerSheetData: Returning 1D array (single row).`);
            return extractedData[0];
        } else if (extractedData[0].length === 1) {
             Logger.log(`fGetServerSheetData: Returning 1D array (single column).`);
             return extractedData.map(row => row[0]);
        } else {
            Logger.log(`fGetServerSheetData: Returning 2D array (${extractedData.length}x${extractedData[0]?.length}).`);
            return extractedData;
        }

    } catch (e) {
        console.error(`Error in fGetServerSheetData for Input "${sheetKeyOrId}", Sheet "${sheetName}", Range ${JSON.stringify(rangeObject)}: ${e.message}\nStack: ${e.stack}`);
        throw new Error(`Server error reading sheet data: ${e.message || e}`);
    }

} // END fGetServerSheetData




// fSetServerSheetData /////////////////////////////////////////////////////////
// Purpose -> Writes data to a specified range in a given Google Sheet,
//            accepting a sheet key, sheet name, a rangeObject with tags/indices,
//            and a 2D array of values to write.
// Inputs  -> sheetKey (String): Key corresponding to a Sheet ID in gSrv.ids.sheets.
//         -> sheetName (String): The name of the specific sheet (tab) to write to.
//         -> rangeObject (Object): {r1, c1, r2, c2} using tags or 0-based indices.
//         -> values (Array[][]): The 2D array of values to write.
// Outputs -> (Boolean): True on success.
// Throws  -> (Error): If inputs are invalid, sheet/range access fails, or
//                     value dimensions don't match range dimensions.
function fSetServerSheetData(sheetKey, sheetName, rangeObject, values) {
    Logger.log(`fSetServerSheetData: Requesting write to key "${sheetKey}", sheet "${sheetName}", range: ${JSON.stringify(rangeObject)}, data: ${values?.length}x${values?.[0]?.length}`);

    // === 1. Validate Inputs ===
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
    // Validate 'values' is a non-empty 2D array
    if (!Array.isArray(values) || values.length === 0 || !Array.isArray(values[0])) {
        const errorMsg = `Invalid 'values' format. Expected non-empty 2D array. Received: ${JSON.stringify(values).substring(0,100)}...`;
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
        // Read data primarily to get headers for tag mapping
        const fullData = sh.getDataRange().getValues();
        if (fullData.length === 0 || fullData[0]?.length === 0) {
           // Allow writing to an empty sheet, but tag resolution might fail if headers are needed
           console.warn(`fSetServerSheetData: Sheet "${sheetName}" appears empty. Tag resolution might fail if tags are used in rangeObject.`);
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

        // === 5. Get Target Range & Validate Dimensions ===
        const rangeNumRows = Math.abs(r1 - r2) + 1;
        const rangeNumCols = Math.abs(c1 - c2) + 1;
        const valuesNumRows = values.length;
        const valuesNumCols = values[0].length;

        if (rangeNumRows !== valuesNumRows || rangeNumCols !== valuesNumCols) {
            throw new Error(`Dimension mismatch: Resolved range is ${rangeNumRows}x${rangeNumCols}, but provided data is ${valuesNumRows}x${valuesNumCols}.`);
        }

        // Get A1 notation for the resolved range
        const a1Notation = fServerIndicesToA1(r1, c1, r2, c2);
        if (!a1Notation) { // Should be rare if indices are valid
             throw new Error(`Could not convert resolved indices [${r1},${c1},${r2},${c2}] to A1 notation.`);
        }
        Logger.log(`fSetServerSheetData: Target A1 notation: ${a1Notation}`);

        const targetRange = sh.getRange(a1Notation);
        if (!targetRange) { // Should be rare if A1 is valid
             throw new Error(`Could not get target range "${a1Notation}" in sheet "${sheetName}".`);
        }

        // === 6. Write Values ===
        targetRange.setValues(values);
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


