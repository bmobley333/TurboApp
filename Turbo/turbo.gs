// turbo.gs //
// 08.19.2025 //
// New Main for 30.0 //
// v30.0 //

// ==========================================================================
// === Global Variables & doGet ===
// ==========================================================================

// NOTE gIndex must be passed into Turbo.gs from client side

const gSrv = {
  // Using gSrv prefix for server-side globals
  ids: {
    sheets: {
      db: "1IOmKLX_WoU0xO4H3btrte5YrGiJZg-ESn8dcp3PVH98",
      mastercs: "1ALFdCHU1h78f_D4pRPFwNwJHRTlundxZqRc1MsgQZo4",
      masterkl: "10Lky3AYqosxAYmUH-ZBwX9MQ2G6LuMuNuPFE_a2cjHU",
      ps: "1oKvsWrXXu2agRqvaxHEl8_0klMeQomXsT6v2JPvPu3w",
      // Player-specific cs/kl will be handled dynamically via functions below
      cs: "", // this is equal to the old myCS and is the ID from the player's CS (loaded via doGet)
      kl: "", // this is equal to the old myKL and is the ID from the player's KL (loaded via doGet's call to fSrvGetMyKlId)
    },
    docs: {
      cm: "1X9spxfuNS84V2jnLHXUP8jQ1MviOx1ZHv5tdSeIfmGo",
      em: "1v7CzSaAzzhVonWBs68nTjuSfIax_ikHY7J9dC0RadKY",
      rb: "1TX2YYTu7jAuxnYYxybmS9ef7wA6t0vFCUjBs4roXxpE",
      sg: "18tistnwfbIT262cQB7fvyLOMBjATPEiSLr43p-K2wSk",
    },
  },
  // Configuration values
  DATA_TAB_NAME: "Data", // Name of the CS tab holding player-specific IDs
  MYKL_ID_CELL_A1: "F8", // A1 notation for the cell containing MyKL ID in the Data tab
  designerPassword: "Ogluck",
  // Add other server-side constants if needed
};

// ==========================================================================
// === Global Constants & Simple Utilities ===
// ==========================================================================

/** function fSrvConvertIndicesToA1
 * Purpose: Converts 0-based row/column indices to a standard A1 notation string.
 * Assumptions: Inputs are valid, non-negative numbers.
 * @param {number} r1 - 0-based starting row index.
 * @param {number} c1 - 0-based starting column index.
 * @param {number} r2 - 0-based ending row index.
 * @param {number} c2 - 0-based ending column index.
 * @returns {string | null} A1 notation string (e.g., "C5:F10" or "A1"), or null on invalid input.
 */
function fSrvConvertIndicesToA1(r1, c1, r2, c2) {
  if (
    [r1, c1, r2, c2].some(
      (idx) => typeof idx !== "number" || idx < 0 || isNaN(idx)
    )
  ) {
    console.error(
      `fSrvConvertIndicesToA1: Invalid indices provided (${r1},${c1},${r2},${c2})`
    );
    return null;
  }
  const startColA1 = fSrvColToA1(c1);
  const endColA1 = fSrvColToA1(c2);
  const startRowA1 = r1 + 1;
  const endRowA1 = r2 + 1;
  if (r1 === r2 && c1 === c2) {
    return `${startColA1}${startRowA1}`;
  } else {
    return `${startColA1}${startRowA1}:${endColA1}${endRowA1}`;
  }
} // End function fSrvConvertIndicesToA1

/** function fSrvColToA1
 * Purpose: Helper function to convert a 0-based column index to A1 notation letter(s).
 * Assumptions: Input is a non-negative integer.
 * @param {number} col - The 0-based column index.
 * @returns {string} The A1 notation label (e.g., A, Z, AA).
 */
function fSrvColToA1(col) {
  let label = "";
  let c = col;
  while (c >= 0) {
    label = String.fromCharCode((c % 26) + 65) + label;
    c = Math.floor(c / 26) - 1;
  }
  return label;
} // End function fSrvColToA1

/** function doGet
 * Purpose: Serves the web app and injects initial, essential parameters to the client.
 * Assumptions: URL parameters 'csID', 'userEmail', and 'gameVer' are provided.
 * Notes: This is the main server-side entry point for the web app. It populates server-side globals and injects data into the HTML template for client-side use in the `gIndex` object.
 * @param {object} e - The Apps Script event object containing URL parameters.
 * @returns {HtmlOutput} The fully rendered HTML page.
 */
function doGet(e) {
  let csId = null;
  let klId = null;
  let userEmail = null;
  let gameVer = null;

  try {
    csId = e?.parameter?.csID;
    userEmail = e?.parameter?.userEmail;
    gameVer = e?.parameter?.gameVer;

    if (!csId || typeof csId !== "string") {
      console.error(
        "doGet Error: csID parameter not provided or invalid in URL."
      );
      return HtmlService.createHtmlOutput(
        "❌ csID parameter missing or invalid in URL."
      );
    }
    if (!userEmail || typeof userEmail !== "string") {
      console.warn(
        "doGet Warning: userEmail parameter not provided or invalid in URL."
      );
      userEmail = "";
    }
    if (gameVer === null || gameVer === undefined) {
      console.warn("doGet Warning: gameVer parameter not provided in URL.");
      gameVer = "";
    } else if (typeof gameVer !== "string") {
      gameVer = String(gameVer);
    }
    Logger.log(
      `doGet: Received CS ID: ${csId}, User Email: ${userEmail}, Game Ver: ${gameVer}`
    );
    gSrv.ids.sheets.cs = csId;
    Logger.log(`doGet: Assigned gSrv.ids.sheets.cs = ${gSrv.ids.sheets.cs}`);
    klId = fSrvGetMyKlId(csId);
    gSrv.ids.sheets.kl = klId;
    Logger.log(`doGet: Assigned gSrv.ids.sheets.kl = ${gSrv.ids.sheets.kl}`);
    const template = HtmlService.createTemplateFromFile("index");
    template.csID = csId;
    template.userEmail = userEmail;
    template.gameVer = gameVer;

    return template
      .evaluate()
      .setTitle("MetaScape Turbo")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    const errorContext = `CS ID: ${csId ||
      "Unknown"}, Email: ${
      userEmail || "Unknown"
    }, GameVer: ${gameVer ||
      "Unknown"}, KL ID Fetch Attempted: ${
      klId !== null || "No (failed before KL fetch)"
    }`;
    console.error(
      `doGet Error (${errorContext}): ${error.message}${
        error.stack ? "\n" + error.stack : ""
      }`
    );
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
} // End function doGet

/** function fSrvGetMyKlId
 * Purpose: Retrieves the MyKL ID from the player's Character Sheet Data tab.
 * Assumptions: The CS has a tab named 'Data' with the KL ID in cell F8.
 * @param {string} myCsId - The Sheet ID of the player's Character Sheet.
 * @returns {string} The player's MyKL Sheet ID.
 * @throws {Error} If IDs cannot be retrieved or sheets/tabs are not found.
 */
function fSrvGetMyKlId(myCsId) {
  if (!myCsId || typeof myCsId !== "string") {
    console.error(
      "fSrvGetMyKlId Error: Character Sheet ID (myCsId) was not provided or invalid."
    );
    throw new Error("getMyKlId: Character Sheet ID (myCsId) was not provided.");
  }
  Logger.log(`fSrvGetMyKlId: Attempting to get MyKL ID from CS ID: ${myCsId}`);
  try {
    const csSpreadsheet = SpreadsheetApp.openById(myCsId);
    if (!csSpreadsheet) {
      throw new Error(
        `Could not open Character Sheet with ID: ${myCsId}. Check permissions and ID validity.`
      );
    }
    const dataSheet = csSpreadsheet.getSheetByName(gSrv.DATA_TAB_NAME);
    if (!dataSheet) {
      throw new Error(
        `Sheet named "${gSrv.DATA_TAB_NAME}" not found in Character Sheet ID: ${myCsId}.`
      );
    }
    const myKlId = dataSheet.getRange(gSrv.MYKL_ID_CELL_A1).getValue();
    if (!myKlId || typeof myKlId !== "string" || myKlId.trim() === "") {
      throw new Error(
        `Could not retrieve a valid MyKL ID from cell ${gSrv.MYKL_ID_CELL_A1} in sheet "${gSrv.DATA_TAB_NAME}". Value was: "${myKlId}"`
      );
    }
    Logger.log(`fSrvGetMyKlId: Successfully retrieved MyKL ID: ${myKlId}`);
    return myKlId.trim();
  } catch (e) {
    console.error(
      `Error in fSrvGetMyKlId for CS ID ${myCsId}: ${e.message}\nStack: ${e.stack}`
    );
    throw new Error(`Server error getting MyKL ID: ${e.message || e}`);
  }
} // End function fSrvGetMyKlId

/** function fSrvValidateDesignerPassword
 * Purpose: Validates a password attempt against the stored designer password.
 * Assumptions: The designer password is set in the `gSrv` global object.
 * @param {string} passwordAttempt - The password entered by the user.
 * @returns {boolean} True if the password matches, false otherwise.
 */
function fSrvValidateDesignerPassword(passwordAttempt) {
  const funcName = "fSrvValidateDesignerPassword";
  const storedPassword = gSrv.designerPassword;
  if (!storedPassword) {
    console.error(
      `${funcName}: Designer password is not set in server globals (gSrv.designerPassword).`
    );
    Logger.log(
      `${funcName}: ERROR - Designer password not configured on server.`
    );
    return false;
  }
  if (typeof passwordAttempt !== "string") {
    Logger.log(
      `${funcName}: Invalid password attempt type received: ${typeof passwordAttempt}`
    );
    return false;
  }
  const isValid = passwordAttempt === storedPassword;
  Logger.log(
    `${funcName}: Password validation attempt. Result: ${
      isValid ? "Success" : "Failure"
    }`
  );
  return isValid;
} // End function fSrvValidateDesignerPassword

/** function fSrvRefreshUITemplateCache
 * Purpose: Re-reads the Master CS 'Game' tab and saves it to the Firestore UI cache.
 * Assumptions: This is an administrator-only function. The gIndex object is passed from the client.
 * @param {object} gIndex - The client-side gIndex object, containing at least GameVer.
 * @returns {object} A success/failure object like { success: boolean, message?: string }.
 */
function fSrvRefreshUITemplateCache(gIndex) {
  const funcName = "fSrvRefreshUITemplateCache";
  Logger.log(
    `${funcName}: Starting UI Template Cache Refresh for v${gIndex.GameVer}...`
  );
  if (!gIndex || !gIndex.GameVer) {
    const msg = "Game Version is required to create a versioned UI template.";
    Logger.log(`${funcName} Error: ${msg}`);
    return {
      success: false,
      message: msg
    };
  }
  try {
    const firestore = fSrvGetFirestoreInstance();
    if (!firestore) {
      throw new Error("Failed to get Firestore instance.");
    }
    const masterCsId = gSrv.ids.sheets.mastercs;
    if (!masterCsId) {
      throw new Error("MasterCS ID is not defined in server configuration.");
    }
    Logger.log(`   -> Reading from MasterCS ID: ${masterCsId}`);
    const uiTemplateData = fSrvReadCSGameSheet({
      CSID: masterCsId
    });
    if (!uiTemplateData || !uiTemplateData.arr || !uiTemplateData.format) {
      throw new Error(
        "Failed to read or process data from the Master Character Sheet."
      );
    }
    Logger.log(`   -> Successfully read and packaged MasterCS 'Game' tab data.`);
    const gameVerMajor = String(gIndex.GameVer).trim().split(".")[0];
    const collectionName = `v${gameVerMajor} Game UI`;
    const baseDocumentId = `UITemplate`;
    fSrvSaveObjectAsChunkedDocs(
      firestore,
      uiTemplateData,
      collectionName,
      baseDocumentId
    );
    Logger.log(
      `✅ ${funcName}: Successfully triggered refresh and save for UI Template Cache.`
    );
    return {
      success: true,
      message: `UI Template Cache saved to collection '${collectionName}'.`,
    };
  } catch (e) {
    const errorMsg = `Error in ${funcName}: ${e.message}`;
    console.error(errorMsg, e.stack);
    Logger.log(`❌ ${funcName} Error: ${e.message}`);
    return {
      success: false,
      message: e.message
    };
  }
} // End function fSrvRefreshUITemplateCache

/** function fSrvBuildTagMaps
 * Purpose: Builds rowTag and colTag maps from a 2D array.
 * Assumptions: Row tags are in column 0, column tags are in row 0.
 * @param {any[][]} fullData - The 2D array of data.
 * @returns {object} An object { rowTag: {}, colTag: {} } containing the maps.
 */
function fSrvBuildTagMaps(fullData) {
  const rowTagMap = {};
  const colTagMap = {};
  if (!Array.isArray(fullData) || fullData.length === 0) {
    console.warn("fSrvBuildTagMaps: Input data array is empty or invalid.");
    return {
      rowTag: rowTagMap,
      colTag: colTagMap
    };
  }
  const numRows = fullData.length;
  const numCols = fullData[0]?.length || 0;
  for (let r = 0; r < numRows; r++) {
    const rowHeaderCell = fullData[r]?.[0];
    if (typeof rowHeaderCell === "string" && rowHeaderCell.trim()) {
      rowHeaderCell.split(",").forEach((tag) => {
        const trimmedTag = tag.trim();
        if (trimmedTag) {
          rowTagMap[trimmedTag] = r;
        }
      });
    }
  }
  const colHeaderRow = fullData[0];
  if (Array.isArray(colHeaderRow)) {
    for (let c = 0; c < numCols; c++) {
      const colHeaderCell = colHeaderRow[c];
      if (typeof colHeaderCell === "string" && colHeaderCell.trim()) {
        colHeaderCell.split(",").forEach((tag) => {
          const trimmedTag = tag.trim();
          if (trimmedTag) {
            colTagMap[trimmedTag] = c;
          }
        });
      }
    }
  } else {
    console.warn(
      "fSrvBuildTagMaps: Header row (row 0) is missing or invalid. Cannot build column tags."
    );
  }
  return {
    rowTag: rowTagMap,
    colTag: colTagMap
  };
} // End function fSrvBuildTagMaps

/** function fSrvResolveTag
 * Purpose: Resolves a single row or column tag/index using the provided tag maps.
 * Assumptions: The tagMap is a valid object.
 * @param {string | number} tagOrIndex - The tag string or 0-based index.
 * @param {object} tagMap - The corresponding tag map (rowTag or colTag).
 * @param {string} [type="unknown"] - 'row' or 'col' for logging purposes.
 * @returns {number} The resolved 0-based index, or NaN if resolution fails.
 */
function fSrvResolveTag(tagOrIndex, tagMap, type = "unknown") {
  if (typeof tagOrIndex === "number" && !isNaN(tagOrIndex) && tagOrIndex >= 0) {
    return tagOrIndex;
  } else if (typeof tagOrIndex === "string" && tagOrIndex.trim()) {
    const trimmedTag = tagOrIndex.trim();
    if (tagMap.hasOwnProperty(trimmedTag)) {
      return tagMap[trimmedTag];
    } else {
      console.warn(
        `fSrvResolveTag: Could not resolve ${type} tag "${trimmedTag}"`
      );
      return NaN;
    }
  } else {
    console.warn(
      `fSrvResolveTag: Invalid input provided for ${type} resolution:`,
      tagOrIndex
    );
    return NaN;
  }
} // End function fSrvResolveTag

/** function fSrvGetInitialGridData
 * Purpose: Acts as the main data loader, trying a fast path (Firestore cache) before a slow path (Google Sheet).
 * Assumptions: The "slow path" automatically creates the cache for subsequent loads (self-healing).
 * @param {object} gIndex - The client-side gIndex object, containing GameVer and CSID.
 * @returns {object} A wrapper object { data: { arr, format, notesArr }, source: string }.
 */
function fSrvGetInitialGridData(gIndex) {
  const funcName = "fSrvGetInitialGridData";
  try {
    Logger.log("--> Attempting Fast Path: Load UI from Firestore Cache...");
    const cachedData = fSrvGetUITemplateFromCache(gIndex);
    Logger.log("--> ✅ Fast Path SUCCESS: UI loaded from Firestore Cache.");
    return {
      data: cachedData,
      source: "Firestore"
    };
  } catch (e) {
    Logger.log(`--> ℹ️ Fast Path FAILED: ${e.message}. Falling back to Slow Path.`);
    Logger.log("--> Attempting Slow Path: Load UI from Google Sheet...");
    const sheetData = fSrvReadCSGameSheet(gIndex);
    Logger.log("--> ✅ Slow Path SUCCESS: UI loaded from Google Sheet.");
    try {
      Logger.log(
        "   -> Self-Healing: Attempting to save the loaded Sheet data to create the cache for the next load..."
      );
      const gameVerMajor = String(gIndex.GameVer).trim().split(".")[0];
      const collectionName = `v${gameVerMajor} Game UI`;
      const baseDocumentId = `UITemplate`;
      const firestore = fSrvGetFirestoreInstance();
      fSrvSaveObjectAsChunkedDocs(
        firestore,
        sheetData,
        collectionName,
        baseDocumentId
      );
      Logger.log("   -> ✅ Self-Healing: Cache created successfully.");
    } catch (saveError) {
      Logger.log(
        `   -> ⚠️ Self-Healing WARNING: Could not save UI Template to cache after slow load. Error: ${saveError.message}`
      );
    }
    return {
      data: sheetData,
      source: "Google Sheets"
    };
  }
} // End function fSrvGetInitialGridData

/** function fSrvReadCSGameSheet
 * Purpose: Loads full data, format, and notes from the 'Game' sheet of a given spreadsheet ID.
 * Assumptions: The gIndex object contains a valid CSID.
 * @param {object} gIndex - The client-side gIndex object.
 * @returns {object} A structured object { arr, format, notesArr } containing sheet data.
 * @throws {Error} If sheet ID is invalid or the sheet cannot be opened/processed.
 */
function fSrvReadCSGameSheet(gIndex) {
  const funcName = "fSrvReadCSGameSheet";
  try {
    if (!gIndex.CSID || typeof gIndex.CSID !== "string") {
      throw new Error("Invalid or missing Sheet ID provided.");
    }
    Logger.log(`${funcName}: Opening Spreadsheet ID: ${gIndex.CSID}`);
    const ss = SpreadsheetApp.openById(gIndex.CSID);
    if (!ss) {
      throw new Error(
        `Could not open Spreadsheet with ID: ${gIndex.CSID}. Check permissions and ID validity.`
      );
    }
    const sh = ss.getSheetByName("Game");
    if (!sh) {
      throw new Error(
        `Sheet named "Game" not found in Spreadsheet ID: ${gIndex.CSID}.`
      );
    }
    Logger.log(
      `${funcName}: Successfully opened sheet "Game" in ID: ${gIndex.CSID}`
    );
    Logger.log(`${funcName}: Extracting data from sheet "Game"...`);
    return fSrvExtractSheetData(sh);
  } catch (e) {
    const context =
      e.message.includes("openById") ||
      e.message.includes("Spreadsheet with ID") ?
      "opening spreadsheet" :
      e.message.includes("getSheetByName") || e.message.includes("not found") ?
      "getting 'Game' sheet" :
      e.stack?.includes("fSrvExtractSheetData") ?
      "processing sheet data" :
      "during operation";
    const msg = `${funcName}: Error ${context}: ${e.message}`;
    console.error(msg + "\nStack:\n" + e.stack);
    throw new Error(msg);
  }
} // End function fSrvReadCSGameSheet

/** function fSrvExtractSheetData
 * Purpose: Reads all data, formats, and notes from a given sheet object.
 * Assumptions: The sheet object `sh` is a valid Apps Script Sheet object.
 * @param {Sheet} sh - The Google Sheet object to extract data from.
 * @returns {object} A structured object { arr, format, notesArr }.
 */
function fSrvExtractSheetData(sh) {
  const rngData = sh.getDataRange();
  const arr = rngData.getValues();
  if (fSrvIsSheetTrulyEmpty(sh, arr)) {
    return fSrvBuildReturnObject([
      []
    ], {
      bg: [
        []
      ],
      fontColorHex: [
        []
      ],
      weight: [
        []
      ],
      fontSize: [
        []
      ],
      fontStyle: [
        []
      ],
      fontFamily: [
        []
      ],
      wrap: [
        []
      ],
      merges: [],
      colWidths: [],
      borders: [],
    }, [
      []
    ]);
  }
  const numCols = arr[0]?.length || 0;
  const format = fSrvBuildFormatObject(sh, rngData, numCols);
  const notesArr = rngData.getNotes();
  return fSrvBuildReturnObject(arr, format, notesArr);
} // End function fSrvExtractSheetData

/** function fSrvIsSheetTrulyEmpty
 * Purpose: Determines if a sheet has zero real content.
 * Assumptions: Checks for minimal cell content and last row/column data.
 * @param {Sheet} sh - The Google Sheet object.
 * @param {any[][]} arr - The 2D array of values from the sheet.
 * @returns {boolean} True if the sheet is considered empty, false otherwise.
 */
function fSrvIsSheetTrulyEmpty(sh, arr) {
  const numRows = arr.length;
  const numCols = arr[0]?.length || 0;
  const onlyEmpty = numRows === 1 && numCols === 1 && arr[0][0] === "";
  return (
    (numRows === 0 || numCols === 0 || onlyEmpty) &&
    sh.getLastRow() === 0 &&
    sh.getLastColumn() === 0
  );
} // End function fSrvIsSheetTrulyEmpty

/** function fSrvBuildFormatObject
 * Purpose: Constructs the complete formatting object from a sheet.
 * Assumptions: The sheet and data range objects are valid.
 * @param {Sheet} sh - The Google Sheet object.
 * @param {Range} rngData - The DataRange object from the sheet.
 * @param {number} numCols - The number of columns in the data range.
 * @returns {object} The complete format object.
 */
function fSrvBuildFormatObject(sh, rngData, numCols) {
  const fontColorObjects = rngData.getFontColorObjects();
  const fontColorHex = fontColorObjects.map((row) =>
    row.map(
      (obj) =>
      obj?.asRgbColor?.()?.asHexString?.()?.replace(/^#ff/, "#") ?? null
    )
  );
  const colWidths = Array.from({
    length: numCols
  }, (_, c) =>
    sh.getColumnWidth(c + 1)
  );
  const mergedRanges = rngData.getMergedRanges();
  const merges = mergedRanges
    .map((r) => ({
      row: r.getRow() - 1,
      col: r.getColumn() - 1,
      rowspan: r.getNumRows(),
      colspan: r.getNumColumns(),
    }))
    .filter(
      (m) =>
      Number.isInteger(m.row) &&
      Number.isInteger(m.col) &&
      (m.rowspan > 1 || m.colspan > 1)
    );
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
    borders: [],
  };
} // End function fSrvBuildFormatObject

/** function fSrvBuildReturnObject
 * Purpose: Combines arrays into the final return structure for sheet data.
 * Assumptions: The input arrays are correctly structured.
 * @param {any[][]} arr - The 2D array of cell values.
 * @param {object} format - The format object.
 * @param {string[][]} notesArr - The 2D array of cell notes.
 * @returns {object} The combined object { arr, format, notesArr }.
 */
function fSrvBuildReturnObject(arr, format, notesArr) {
  return {
    arr: arr,
    format: format,
    notesArr: notesArr,
  };
} // End function fSrvBuildReturnObject

/** function fSrvLoadFullGoogleSheetAndTags
 * Purpose: Loads all data and tags from a specified sheet in a specified workbook.
 * Assumptions: The workbook and sheet exist and are accessible.
 * @param {string} workbookAbr - Abbreviation ('db', 'mastercs', 'mycs', etc.).
 * @param {string} sheetName - The name of the sheet to read.
 * @param {string} csId - The Character Sheet ID, required for 'mycs' and 'mykl'.
 * @returns {object} An object { ColTags, RowTags, sheetText2D }.
 * @throws {Error} If inputs are invalid, sheets are not found, or duplicate tags exist.
 */
function fSrvLoadFullGoogleSheetAndTags(workbookAbr, sheetName, csId) {
  const funcName = "fSrvLoadFullGoogleSheetAndTags";
  Logger.log(
    `${funcName}: Loading Tags & Data for Workbook: "${workbookAbr}", Sheet: "${sheetName}", CSID: "${csId}".`
  );
  if (!workbookAbr || typeof workbookAbr !== "string") {
    throw new Error(`${funcName}: Invalid or missing workbookAbr provided.`);
  }
  if (!sheetName || typeof sheetName !== "string" || sheetName.trim() === "") {
    throw new Error(`${funcName}: Invalid or empty sheetName provided.`);
  }
  const trimmedSheetName = sheetName.trim();
  const lowerWorkbookAbr = workbookAbr.toLowerCase();
  let workbookID = null;
  try {
    switch (lowerWorkbookAbr) {
      case "db":
        workbookID = gSrv.ids.sheets.db;
        break;
      case "mastercs":
        workbookID = gSrv.ids.sheets.mastercs;
        break;
      case "masterkl":
        workbookID = gSrv.ids.sheets.masterkl;
        break;
      case "mycs":
        if (!csId) {
          throw new Error("CSID is required to identify 'MyCS' workbook.");
        }
        workbookID = csId;
        break;
      case "mykl":
        if (!csId) {
          throw new Error("CSID is required to look up 'MyKL' workbook ID.");
        }
        workbookID = fSrvGetMyKlId(csId);
        if (!workbookID) {
          throw new Error(
            `Could not find linked 'MyKL' workbook ID for CSID: ${csId}`
          );
        }
        break;
      default:
        throw new Error(`Unsupported workbook abbreviation: "${workbookAbr}"`);
    }
    if (!workbookID) {
      throw new Error(
        `Workbook ID for "${workbookAbr}" resolved to null or empty.`
      );
    }
    Logger.log(
      `   -> Resolved Workbook ID for "${workbookAbr}": ${workbookID}`
    );
  } catch (e) {
    console.error(`${funcName} Error resolving Workbook ID: ${e.message}`);
    throw new Error(
      `Failed to resolve Workbook ID for "${workbookAbr}": ${e.message}`
    );
  }
  let ss, sh;
  try {
    ss = SpreadsheetApp.openById(workbookID);
    sh = ss.getSheetByName(trimmedSheetName);
    if (!sh) {
      throw new Error(
        `Sheet named "${trimmedSheetName}" not found in Workbook ID: ${workbookID} (Abr: ${workbookAbr}).`
      );
    }
    Logger.log(
      `   -> Successfully opened sheet "${trimmedSheetName}" in Workbook "${workbookAbr}".`
    );
  } catch (e) {
    console.error(
      `Error accessing sheet "${trimmedSheetName}" in Workbook "${workbookAbr}" (ID: ${workbookID}): ${e.message}`
    );
    throw new Error(
      `Failed to access sheet "${trimmedSheetName}" in Workbook "${workbookAbr}": ${e.message}`
    );
  }
  const dataRange = sh.getDataRange();
  const sheetText2D = dataRange.getValues();
  const numRows = sheetText2D.length;
  const numCols = numRows > 0 ? sheetText2D[0]?.length || 0 : 0;
  if (
    numRows === 0 ||
    numCols === 0 ||
    (numRows === 1 && numCols === 1 && sheetText2D[0][0] === "")
  ) {
    throw new Error(
      `${funcName}: Sheet "${trimmedSheetName}" in Workbook "${workbookAbr}" exists but appears empty.`
    );
  }
  Logger.log(
    `   -> Read ${numRows}x${numCols} cells from "${trimmedSheetName}".`
  );
  const colTagsMap = {};
  const colTagsSeen = new Set();
  const headerRow = sheetText2D[0];
  for (let c = 0; c < numCols; c++) {
    const cellValue = headerRow[c];
    if (typeof cellValue === "string" && cellValue.trim() !== "") {
      const tags = cellValue
        .split(",")
        .map((tag) => tag.trim())
        .filter(Boolean);
      for (const tag of tags) {
        const lowerTag = tag.toLowerCase();
        if (colTagsSeen.has(lowerTag)) {
          throw new Error(
            `${funcName}: Duplicate Column Tag found (case-insensitive): "${tag}" in Row 0, Column ${
              c + 1
            }. Sheet: "${trimmedSheetName}", Workbook: "${workbookAbr}".`
          );
        }
        if (colTagsMap.hasOwnProperty(tag)) {
          Logger.log(
            `   -> WARNING: Column Tag "${tag}" reused in Row 0. Mapping to last occurrence (Col ${c}).`
          );
        }
        colTagsMap[tag] = c;
        colTagsSeen.add(lowerTag);
      }
    }
  }
  Logger.log(
    `   -> Processed ${Object.keys(colTagsMap).length} unique column tags.`
  );
  const rowTagsMap = {};
  const rowTagsSeen = new Set();
  for (let r = 0; r < numRows; r++) {
    const cellValue =
      sheetText2D[r] && sheetText2D[r].length > 0 ?
      sheetText2D[r][0] :
      undefined;
    if (typeof cellValue === "string" && cellValue.trim() !== "") {
      const tags = cellValue
        .split(",")
        .map((tag) => tag.trim())
        .filter(Boolean);
      for (const tag of tags) {
        const lowerTag = tag.toLowerCase();
        if (rowTagsSeen.has(lowerTag)) {
          throw new Error(
            `${funcName}: Duplicate Row Tag found (case-insensitive): "${tag}" in Column 0, Row ${
              r + 1
            }. Sheet: "${trimmedSheetName}", Workbook: "${workbookAbr}".`
          );
        }
        if (rowTagsMap.hasOwnProperty(tag)) {
          Logger.log(
            `   -> WARNING: Row Tag "${tag}" reused in Col 0. Mapping to last occurrence (Row ${r}).`
          );
        }
        rowTagsMap[tag] = r;
        rowTagsSeen.add(lowerTag);
      }
    }
  }
  Logger.log(
    `   -> Processed ${Object.keys(rowTagsMap).length} unique row tags.`
  );
  Logger.log(
    `${funcName}: Successfully loaded tags and data for Workbook "${workbookAbr}", Sheet "${trimmedSheetName}".`
  );
  return {
    ColTags: colTagsMap,
    RowTags: rowTagsMap,
    sheetText2D: sheetText2D,
  };
} // End function fSrvLoadFullGoogleSheetAndTags

/** function fSrvGetSheetRangeDataNTags
 * Purpose: Reads data and tags from a specified range within a sheet.
 * Assumptions: Can handle full sheet reads or specific range slices. 'Calc_LastRow' is a special tag.
 * @param {string} sheetKeyOrId - Key from gSrv.ids.sheets OR a direct Sheet fileId.
 * @param {string} sheetName - The name of the sheet to read.
 * @param {object} [rangeObject=null] - Optional. {r1, c1, r2, c2} using tags or indices.
 * @returns {object} An object with { data, colTags, rowTags }.
 * @throws {Error} If inputs are invalid or sheet/range access fails.
 */
function fSrvGetSheetRangeDataNTags(
  sheetKeyOrId,
  sheetName,
  rangeObject = null
) {
  const rangeLog = rangeObject ?
    JSON.stringify(rangeObject) :
    "Not Provided (Full Sheet)";
  Logger.log(
    `fSrvGetSheetRangeDataNTags: Request received. Key/ID: "${sheetKeyOrId}", Sheet: "${sheetName}", Range: ${rangeLog}`
  );
  let fileId = null;
  let identifiedBy = "";
  let absoluteRowTagMap = {};
  let absoluteColTagMap = {};
  let relativeRowTagMap = {};
  let relativeColTagMap = {};
  if (!sheetKeyOrId || typeof sheetKeyOrId !== "string") {
    const errorMsg = `Invalid or missing sheetKeyOrId parameter.`;
    console.error(`fSrvGetSheetRangeDataNTags Error: ${errorMsg}`);
    throw new Error(`getServerSheetData: ${errorMsg}`);
  }
  const isLikelySheetId =
    sheetKeyOrId.length > 30 && /^[a-z0-9_-]+$/i.test(sheetKeyOrId);
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
    Logger.log(
      `   -> Interpreted first argument as Key: "${sheetKeyOrId}", resolved to ID: ${fileId}`
    );
  }
  if (!fileId) {
    throw new Error(`getServerSheetData: Could not determine File ID.`);
  }
  if (!sheetName || typeof sheetName !== "string") {
    throw new Error(`getServerSheetData: Invalid or missing sheetName.`);
  }
  let isRangeProvidedAndValid = false;
  if (rangeObject && typeof rangeObject === "object") {
    if (
      rangeObject.r1 !== undefined &&
      rangeObject.c1 !== undefined &&
      rangeObject.r2 !== undefined &&
      rangeObject.c2 !== undefined
    ) {
      isRangeProvidedAndValid = true;
      Logger.log(`   -> Valid rangeObject structure provided.`);
    } else {
      Logger.log(
        `   -> rangeObject provided but incomplete: ${JSON.stringify(
          rangeObject
        )}. Defaulting to full sheet.`
      );
      rangeObject = null;
    }
  } else if (rangeObject) {
    Logger.log(
      `   -> rangeObject provided but not an object: ${typeof rangeObject}. Defaulting to full sheet.`
    );
    rangeObject = null;
  } else {
    Logger.log(`   -> rangeObject not provided. Defaulting to full sheet.`);
  }
  try {
    const ss = SpreadsheetApp.openById(fileId);
    const sh = ss.getSheetByName(sheetName);
    if (!sh) {
      throw new Error(
        `Sheet named "${sheetName}" not found in Sheet (${identifiedBy}).`
      );
    }
    const fullData = sh.getDataRange().getValues();
    const numRows = fullData.length;
    const numCols = fullData[0]?.length || 0;
    Logger.log(
      `fSrvGetSheetRangeDataNTags: Read ${numRows}x${numCols} cells from "${sheetName}".`
    );
    const absoluteTagMaps = fSrvBuildTagMaps(fullData);
    absoluteRowTagMap = absoluteTagMaps.rowTag;
    absoluteColTagMap = absoluteTagMaps.colTag;
    if (numRows === 0 || numCols === 0) {
      console.warn(
        `fSrvGetSheetRangeDataNTags: Sheet "${sheetName}" (${identifiedBy}) appears empty.`
      );
      return {
        data: [
          []
        ],
        colTags: {},
        rowTags: {}
      };
    }
    if (isRangeProvidedAndValid) {
      Logger.log(`   -> Processing provided range...`);
      const r1_abs = fSrvResolveTag(rangeObject.r1, absoluteRowTagMap, "row");
      const c1_abs = fSrvResolveTag(rangeObject.c1, absoluteColTagMap, "col");
      const c2_abs = fSrvResolveTag(rangeObject.c2, absoluteColTagMap, "col");
      let r2_abs;
      if (
        typeof rangeObject.r2 === "string" &&
        rangeObject.r2.toLowerCase() === "calc_lastrow"
      ) {
        if (isNaN(c1_abs)) {
          throw new Error(
            `Cannot calculate last row: Column 'c1' (${rangeObject.c1}) could not be resolved.`
          );
        }
        Logger.log(`   -> Calculating last row for column index ${c1_abs}...`);
        const lastSheetRow = sh.getLastRow();
        r2_abs = -1;
        for (let r = lastSheetRow - 1; r >= 0; r--) {
          const cellValue = fullData[r]?.[c1_abs];
          if (
            cellValue !== undefined &&
            cellValue !== null &&
            String(cellValue).trim() !== ""
          ) {
            r2_abs = r;
            Logger.log(
              `   -> Found last non-empty cell at row index ${r2_abs}.`
            );
            break;
          }
        }
        if (r2_abs === -1) {
          if (!isNaN(r1_abs)) {
            r2_abs = r1_abs;
            Logger.log(
              `   -> Warning: Could not find last non-empty row in column ${c1_abs}. Using r1 index ${r1_abs} as fallback.`
            );
          } else {
            throw new Error(
              `Cannot calculate last row: Column ${c1_abs} appears empty and r1 ('${rangeObject.r1}') is also invalid.`
            );
          }
        }
      } else {
        r2_abs = fSrvResolveTag(rangeObject.r2, absoluteRowTagMap, "row");
      }
      if ([r1_abs, c1_abs, r2_abs, c2_abs].some(isNaN)) {
        let failedTags = [];
        if (isNaN(r1_abs)) failedTags.push(`r1: ${rangeObject.r1}`);
        if (isNaN(c1_abs)) failedTags.push(`c1: ${rangeObject.c1}`);
        if (isNaN(r2_abs))
          failedTags.push(`r2: ${rangeObject.r2} (resolved: ${r2_abs})`);
        if (isNaN(c2_abs)) failedTags.push(`c2: ${rangeObject.c2}`);
        throw new Error(
          `Could not resolve absolute tags/indices in rangeObject: ${JSON.stringify(
            rangeObject
          )}. Failed Tags: ${failedTags.join(", ")}`
        );
      }
      Logger.log(
        `fSrvGetSheetRangeDataNTags: Resolved range to absolute indices: r1=${r1_abs}, c1=${c1_abs}, r2=${r2_abs}, c2=${c2_abs}`
      );
      const rStart = Math.min(r1_abs, r2_abs);
      const rEnd = Math.max(r1_abs, r2_abs);
      const cStart = Math.min(c1_abs, c2_abs);
      const cEnd = Math.max(c1_abs, c2_abs);
      if (rStart >= numRows || cStart >= numCols) {
        console.warn(
          `fSrvGetSheetRangeDataNTags: Resolved range start [${rStart}, ${cStart}] is outside the bounds of the sheet data [${numRows}, ${numCols}]. Returning empty data.`
        );
        return {
          data: [
            []
          ],
          colTags: {},
          rowTags: {}
        };
      }
      const extractedData = fullData
        .slice(rStart, rEnd + 1)
        .map((row) => row.slice(cStart, cEnd + 1));
      const extractedRows = extractedData.length;
      const extractedCols = extractedData[0]?.length || 0;
      for (const tag in absoluteColTagMap) {
        const absoluteIndex = absoluteColTagMap[tag];
        if (absoluteIndex >= cStart && absoluteIndex <= cEnd) {
          const relativeIndex = absoluteIndex - cStart;
          relativeColTagMap[tag] = relativeIndex;
        }
      }
      for (const tag in absoluteRowTagMap) {
        const absoluteIndex = absoluteRowTagMap[tag];
        if (absoluteIndex >= rStart && absoluteIndex <= rEnd) {
          const relativeIndex = absoluteIndex - rStart;
          relativeRowTagMap[tag] = relativeIndex;
        }
      }
      Logger.log(
        `fSrvGetSheetRangeDataNTags: Built relative tag maps for slice. Rel Rows: ${
          Object.keys(relativeRowTagMap).length
        }, Rel Cols: ${Object.keys(relativeColTagMap).length}`
      );
      let returnData;
      if (extractedRows === 1 && extractedCols === 1) {
        Logger.log(
          `fSrvGetSheetRangeDataNTags: Formatting sliced data as single value.`
        );
        returnData = extractedData[0][0];
      } else if (extractedRows === 1) {
        Logger.log(
          `fSrvGetSheetRangeDataNTags: Formatting sliced data as 1D array (single row).`
        );
        returnData = extractedData[0];
      } else if (
        extractedCols === 1 &&
        extractedData.every((row) => row.length === 1)
      ) {
        Logger.log(
          `fSrvGetSheetRangeDataNTags: Formatting sliced data as 1D array (single column).`
        );
        returnData = extractedData.map((row) => row[0]);
      } else {
        Logger.log(
          `fSrvGetSheetRangeDataNTags: Formatting sliced data as 2D array (${extractedRows}x${extractedCols}).`
        );
        returnData = extractedData;
      }
      return {
        data: returnData,
        colTags: relativeColTagMap,
        rowTags: relativeRowTagMap,
      };
    } else {
      Logger.log(
        `   -> No range provided or range invalid. Returning full sheet data...`
      );
      let returnData;
      if (numRows === 1 && numCols === 1) {
        Logger.log(
          `fSrvGetSheetRangeDataNTags: Formatting full sheet data as single value.`
        );
        returnData = fullData[0][0];
      } else if (numRows === 1) {
        Logger.log(
          `fSrvGetSheetRangeDataNTags: Formatting full sheet data as 1D array (single row).`
        );
        returnData = fullData[0];
      } else if (numCols === 1 && fullData.every((row) => row.length === 1)) {
        Logger.log(
          `fSrvGetSheetRangeDataNTags: Formatting full sheet data as 1D array (single column).`
        );
        returnData = fullData.map((row) => row[0]);
      } else {
        Logger.log(
          `fSrvGetSheetRangeDataNTags: Formatting full sheet data as 2D array (${numRows}x${numCols}).`
        );
        returnData = fullData;
      }
      return {
        data: returnData,
        colTags: absoluteColTagMap,
        rowTags: absoluteRowTagMap,
      };
    }
  } catch (e) {
    console.error(
      `Error in fSrvGetSheetRangeDataNTags for Input "${sheetKeyOrId}", Sheet "${sheetName}", Range ${rangeLog}: ${e.message}\nStack: ${e.stack}`
    );
    throw new Error(`Server error processing sheet data: ${e.message || e}`);
  }
} // End function fSrvGetSheetRangeDataNTags

/** function fSrvSaveURLtoNamesAndLogToDBandPS
 * Purpose: Writes bundled log and header data to the GMScreen and PartyLog sheets.
 * Assumptions: The dataBundle contains all necessary keys.
 * @param {object} dataBundle - An object containing { log, vit, nish, url, raceClass, level, playerChar, slotNum }.
 * @returns {boolean} True if both writes succeeded, false otherwise.
 * @throws {Error} If critical errors occur.
 */
function fSrvSaveURLtoNamesAndLogToDBandPS(dataBundle) {
  const funcName = "fSrvSaveURLtoNamesAndLogToDBandPS";
  Logger.log(
    `${funcName}: Received data bundle. Preparing to write to DB and PS.`
  );
  const requiredKeys = [
    "log", "vit", "nish", "url", "raceClass", "level", "playerChar", "slotNum",
  ];
  if (
    !dataBundle ||
    typeof dataBundle !== "object" ||
    !requiredKeys.every((key) => dataBundle.hasOwnProperty(key))
  ) {
    const missing = requiredKeys.filter(
      (key) => !dataBundle?.hasOwnProperty(key)
    );
    const errorMsg = `Invalid or incomplete dataBundle received. Missing keys: ${missing.join(
      ", "
    )}`;
    console.error(`${funcName} Error: ${errorMsg}`);
    throw new Error(`${funcName}: ${errorMsg}`);
  }
  const slotNumTag = dataBundle.slotNum;
  if (
    !slotNumTag ||
    typeof slotNumTag !== "string" ||
    !slotNumTag.startsWith("Slot")
  ) {
    const errorMsg = `Invalid slotNum ("${slotNumTag}") received in dataBundle. Must be a valid Slot tag (e.g., 'Slot3').`;
    console.error(`${funcName} Error: ${errorMsg}`);
    return false;
  }
  Logger.log(`${funcName}: Using Slot Tag: ${slotNumTag}`);
  const targets = [{
    key: "db",
    sheetName: "GMScreen"
  }, {
    key: "ps",
    sheetName: "PartyLog"
  }, ];
  const baseCellTagR = "Log";
  const baseCellTagC = slotNumTag;
  const numHeaderRows = 6;
  const dataToWrite = [
    [dataBundle.url],
    [dataBundle.raceClass],
    [dataBundle.level],
    [dataBundle.vit],
    [dataBundle.nish],
    [dataBundle.playerChar],
    [dataBundle.log],
  ];
  let overallSuccess = true;
  for (const target of targets) {
    Logger.log(
      `${funcName}: Processing target: Key='${target.key}', Sheet='${target.sheetName}'`
    );
    let success = false;
    try {
      const fileId = gSrv.ids.sheets[target.key.toLowerCase()];
      if (!fileId) {
        throw new Error(`Could not find File ID for key '${target.key}'`);
      }
      const ss = SpreadsheetApp.openById(fileId);
      const sh = ss.getSheetByName(target.sheetName);
      if (!sh) {
        throw new Error(
          `Sheet named "${target.sheetName}" not found in Sheet ID: ${fileId} (Key: ${target.key}).`
        );
      }
      const fullData = sh.getDataRange().getValues();
      if (fullData.length === 0 || fullData[0]?.length === 0) {
        console.warn(
          `${funcName}: Target sheet "${target.sheetName}" appears empty. Cannot resolve base cell.`
        );
        throw new Error(`Target sheet "${target.sheetName}" is empty.`);
      }
      const {
        rowTag,
        colTag
      } = fSrvBuildTagMaps(fullData);
      const baseRowIndex = fSrvResolveTag(baseCellTagR, rowTag, "row");
      const baseColIndex = fSrvResolveTag(baseCellTagC, colTag, "col");
      if (isNaN(baseRowIndex) || isNaN(baseColIndex)) {
        throw new Error(
          `Could not resolve base cell tags ('${baseCellTagR}', '${baseCellTagC}') in sheet "${target.sheetName}".`
        );
      }
      Logger.log(
        `   -> Resolved Base Cell ('${baseCellTagR}', '${baseCellTagC}') to [${baseRowIndex}, ${baseColIndex}] in "${target.sheetName}".`
      );
      const startRowIndex = baseRowIndex - numHeaderRows;
      const startColIndex = baseColIndex;
      const numRowsToWrite = dataToWrite.length;
      const numColsToWrite = 1;
      if (startRowIndex < 0) {
        throw new Error(
          `Calculated start row index (${startRowIndex}) is invalid (must be >= 0). Base cell ('${baseCellTagR}') might be too high.`
        );
      }
      const targetRange = sh.getRange(
        startRowIndex + 1,
        startColIndex + 1,
        numRowsToWrite,
        numColsToWrite
      );
      const targetA1 = targetRange.getA1Notation();
      Logger.log(
        `   -> Target range calculated: ${targetA1} (${numRowsToWrite}x${numColsToWrite})`
      );
      targetRange.setValues(dataToWrite);
      Logger.log(
        `   -> Successfully wrote data to ${targetA1} in sheet "${target.sheetName}".`
      );
      success = true;
    } catch (e) {
      console.error(
        `Error writing to target ${target.key}/${target.sheetName}: ${e.message}\nStack: ${e.stack}`
      );
      overallSuccess = false;
    }
  }
  Logger.log(
    `${funcName}: Finished processing all targets. Overall Success: ${overallSuccess}`
  );
  if (!overallSuccess) {
    Logger.log(`${funcName}: Write failed for at least one target.`);
  }
  return overallSuccess;
} // End function fSrvSaveURLtoNamesAndLogToDBandPS

/** function fSrvGetUITemplateFromCache
 * Purpose: Reads and reassembles a chunked UI Template from the Firestore cache.
 * Assumptions: This is the "fast path" for loading the initial UI.
 * @param {object} gIndex - The client-side gIndex object, containing at least GameVer.
 * @returns {object} The reassembled UI data object parsed from JSON.
 * @throws {Error} If the cache is not found, is malformed, or cannot be parsed.
 */
function fSrvGetUITemplateFromCache(gIndex) {
  const funcName = "fSrvGetUITemplateFromCache";
  const firestore = fSrvGetFirestoreInstance();
  if (!firestore) {
    throw new Error("Could not get Firestore instance.");
  }
  const gameVerMajor = String(gIndex.GameVer).trim().split(".")[0];
  const collectionName = `v${gameVerMajor} Game UI`;
  const baseDocumentId = `UITemplate`;
  const metadataPath = `${collectionName}/${baseDocumentId}_metadata`;
  Logger.log(`   -> ${funcName}: Attempting to load from ${metadataPath}`);
  const metadataDoc = firestore.getDocument(metadataPath);
  if (metadataDoc && metadataDoc.fields) {
    const metadataFields = metadataDoc.fields;
    const metadata = {};
    for (const key in metadataFields) {
      metadata[key] = fSrvConvertFirestoreTypesToJS(metadataFields[key]);
    }
    if (typeof metadata.totalChunks === "number") {
      const totalChunks = metadata.totalChunks;
      Logger.log(
        `   -> ${funcName}: Metadata found. Total chunks to load: ${totalChunks}.`
      );
      if (totalChunks === 0) {
        return {
          arr: [
            []
          ],
          format: {},
          notesArr: [
            []
          ]
        };
      }
      let jsonStringChunks = new Array(totalChunks);
      for (let i = 1; i <= totalChunks; i++) {
        const chunkDocId = `${baseDocumentId}_chunk_${i}of${totalChunks}`;
        const chunkPath = `${collectionName}/${chunkDocId}`;
        const chunkDoc = firestore.getDocument(chunkPath);
        const chunkFields = chunkDoc.fields;
        const chunkData = {};
        for (const key in chunkFields) {
          chunkData[key] = fSrvConvertFirestoreTypesToJS(chunkFields[key]);
        }
        if (typeof chunkData.chunkData !== "string") {
          throw new Error(`Data in chunk ${i} is not a string.`);
        }
        jsonStringChunks[chunkData._chunkIndex] = chunkData.chunkData;
      }
      const fullJsonString = jsonStringChunks.join("");
      Logger.log(
        `   -> ${funcName}: All ${totalChunks} chunks loaded and reassembled. Parsing...`
      );
      return JSON.parse(fullJsonString);
    }
    throw new Error(
      "Metadata document is malformed (missing or invalid totalChunks)."
    );
  }
  throw new Error("Metadata document not found or is empty.");
} // End function fSrvGetUITemplateFromCache

/** function fSrvSaveObjectAsChunkedDocs
 * Purpose: Saves a large JavaScript object to Firestore as a series of chunked documents.
 * Assumptions: The Firestore instance is valid.
 * @param {object} firestore - The authenticated Firestore instance.
 * @param {object} objectToSave - The large JavaScript object to be saved.
 * @param {string} collectionName - The name of the Firestore collection.
 * @param {string} baseDocumentId - The base name for the documents.
 * @returns {void}
 * @throws {Error} If saving fails.
 */
function fSrvSaveObjectAsChunkedDocs(
  firestore,
  objectToSave,
  collectionName,
  baseDocumentId
) {
  const funcName = "fSrvSaveObjectAsChunkedDocs";
  const MAX_CHUNK_SIZE = 800000;
  const jsonString = JSON.stringify(objectToSave);
  const totalSize = jsonString.length;
  Logger.log(
    `   -> ${funcName}: Serialized object to JSON string of size ${totalSize} chars.`
  );
  const chunks = [];
  for (let i = 0; i < totalSize; i += MAX_CHUNK_SIZE) {
    chunks.push(jsonString.substring(i, i + MAX_CHUNK_SIZE));
  }
  const totalChunks = chunks.length;
  Logger.log(`   -> ${funcName}: Split JSON string into ${totalChunks} chunk(s).`);
  const metadataDocId = `${baseDocumentId}_metadata`;
  const metadataPath = `${collectionName}/${metadataDocId}`;
  const metadataObject = {
    totalChunks: totalChunks,
    totalSize: totalSize,
    _lastUpdated: new Date(),
  };
  Logger.log(`   -> ${funcName}: Saving metadata to ${metadataPath}`);
  firestore.updateDocument(metadataPath, metadataObject, false);
  for (let i = 0; i < totalChunks; i++) {
    const chunkIndex = i + 1;
    const chunkDocId = `${baseDocumentId}_chunk_${chunkIndex}of${totalChunks}`;
    const chunkPath = `${collectionName}/${chunkDocId}`;
    const chunkData = {
      chunkData: chunks[i],
      _chunkIndex: i,
    };
    Logger.log(
      `   -> ${funcName}: Saving data chunk ${chunkIndex}/${totalChunks} to ${chunkPath}`
    );
    firestore.updateDocument(chunkPath, chunkData, false);
  }
  Logger.log(`   -> ${funcName}: All chunks saved successfully.`);
} // End function fSrvSaveObjectAsChunkedDocs

/** function fSrvGetFirestoreInstance
 * Purpose: Initializes and returns an authenticated Firestore instance.
 * Assumptions: Credentials are stored in PropertiesService.
 * @param {}
 * @returns {object | null} Authenticated Firestore instance or null on error.
 */
function fSrvGetFirestoreInstance() {
  const funcName = "fSrvGetFirestoreInstance";
  Logger.log(`${funcName}: Attempting to initialize Firestore...`);
  let clientEmail, privateKeyRaw, projectId, processedKey;
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    clientEmail = scriptProperties.getProperty("FIRESTORE_CLIENT_EMAIL");
    privateKeyRaw = scriptProperties.getProperty("FIRESTORE_PRIVATE_KEY");
    projectId = scriptProperties.getProperty("FIRESTORE_PROJECT_ID");
    let missingCred = false;
    if (!clientEmail) {
      Logger.log(
        `   -> ${funcName} Error: FIRESTORE_CLIENT_EMAIL not found or empty.`
      );
      missingCred = true;
    }
    if (!privateKeyRaw) {
      Logger.log(
        `   -> ${funcName} Error: FIRESTORE_PRIVATE_KEY not found or empty.`
      );
      missingCred = true;
    }
    if (!projectId) {
      Logger.log(
        `   -> ${funcName} Error: FIRESTORE_PROJECT_ID not found or empty.`
      );
      missingCred = true;
    }
    if (missingCred) {
      console.error(
        `${funcName} Error: Missing Firestore credentials in Script Properties.`
      );
      return null;
    }
    processedKey = privateKeyRaw;
    if (processedKey.startsWith('"') && processedKey.endsWith('"')) {
      processedKey = processedKey.substring(1, processedKey.length - 1);
    }
    processedKey = processedKey.replaceAll("\\n", "\n");
    Logger.log(
      `   -> Calling FirestoreApp.getFirestore for project ${projectId}...`
    );
    const firestore = FirestoreApp.getFirestore(
      clientEmail,
      processedKey,
      projectId
    );
    if (!firestore) {
      Logger.log(
        `   -> ${funcName} Error: FirestoreApp.getFirestore returned null/undefined.`
      );
      console.error(
        `${funcName} Error: FirestoreApp.getFirestore failed to return an instance.`
      );
      return null;
    }
    Logger.log(
      `${funcName}: Firestore instance initialized successfully for project ${projectId}.`
    );
    return firestore;
  } catch (e) {
    console.error(
      `Error caught in ${funcName}: ${e.message}\nStack: ${e.stack}`
    );
    Logger.log(
      `   -> ❌ Exception during Firestore initialization: ${e.message}`
    );
    Logger.log(
      `   -> Details at time of error: ProjectID=${projectId || "N/A"}, Email=${
        clientEmail || "N/A"
      }`
    );
    return null;
  }
} // End function fSrvGetFirestoreInstance

/** function fSrvSaveTurboDataToFirestore
 * Purpose: Saves grid and custom notes data to a user-specific Firestore document.
 * Assumptions: Data is processed into an Array of Row Objects for Firestore compatibility.
 * @param {object} gIndex - Object from the client containing Email, CSID, and GameVer.
 * @param {any[][]} fullArrData - The complete gUI.arr from the client.
 * @param {string[][]} customNotesData - The complete gUI.customNotes array.
 * @param {object} charInfo - DEPRECATED/UNUSED.
 * @returns {object} An object { success: boolean, message?: string }.
 */
function fSrvSaveTurboDataToFirestore(
  gIndex,
  fullArrData,
  customNotesData,
  charInfo
) {
  const funcName = "fSrvSaveTurboDataToFirestore";
  Logger.log(
    `${funcName}: Saving Grid & Notes for User: ${gIndex.Email}, CS ID: ${gIndex.CSID}...`
  );
  if (!gIndex.GameVer || typeof gIndex.GameVer !== "string" || gIndex.GameVer.trim() === "") {
    return {
      success: false,
      message: "Invalid or missing Game Version provided."
    };
  }
  if (!gIndex.Email || typeof gIndex.Email !== "string" || gIndex.Email.indexOf("@") === -1) {
    return {
      success: false,
      message: "Invalid User Email provided."
    };
  }
  if (!gIndex.CSID || typeof gIndex.CSID !== "string") {
    return {
      success: false,
      message: "Invalid Character Sheet ID provided."
    };
  }
  if (!Array.isArray(fullArrData) || (fullArrData.length > 0 && !Array.isArray(fullArrData[0]))) {
    return {
      success: false,
      message: "Invalid fullArrData provided (must be 2D array).",
    };
  }
  if (!Array.isArray(customNotesData)) {
    return {
      success: false,
      message: "Invalid customNotesData provided."
    };
  }
  const firestore = fSrvGetFirestoreInstance();
  if (!firestore) {
    const msg = "Failed to initialize Firestore instance.";
    Logger.log(`${funcName} Error: ${msg}`);
    return {
      success: false,
      message: "Server configuration error (Firestore).",
    };
  }
  const gameVerMajor = String(gIndex.GameVer).trim().split(".")[0];
  const collectionPath = `v${gameVerMajor} ${gIndex.Email}`;
  const gameDocId = `Turbo_Game_${gIndex.CSID}`;
  const gameDocPath = `${collectionPath}/${gameDocId}`;
  Logger.log(`   -> Target Firestore Path: ${gameDocPath}`);
  const processArray = (arr) => {
    const arrayOfObjects = [];
    const numRows = arr.length;
    for (let r = 0; r < numRows; r++) {
      const rowData = arr[r] || [];
      const rowKey = `row${r}`;
      const rowObject = {};
      rowObject[rowKey] = rowData;
      arrayOfObjects.push(rowObject);
    }
    return arrayOfObjects;
  };
  const gUIarrForFirestore = processArray(fullArrData);
  const customNotesForFirestore = processArray(customNotesData);
  const dataToSave = {
    gUIarr: gUIarrForFirestore,
    customNotes: customNotesForFirestore,
    _lastUpdated: new Date(),
  };
  try {
    Logger.log(
      `   -> Calling firestore.updateDocument for path: ${gameDocPath}...`
    );
    firestore.updateDocument(gameDocPath, dataToSave, false);
    Logger.log(`      -> ✅ Successfully saved data.`);
    return {
      success: true
    };
  } catch (e) {
    console.error(
      `Exception saving data to ${gameDocPath}: ${e.message}\nStack: ${e.stack}`
    );
    Logger.log(`   -> ❌ Exception during data save: ${e.message}`);
    return {
      success: false,
      message: `Firestore save failed: ${e.message}`
    };
  }
} // End function fSrvSaveTurboDataToFirestore

/** function fSrvConvertFirestoreTypesToJS
 * Purpose: Recursively converts Firestore's typed value objects into standard JavaScript types.
 * Assumptions: Input is a value object from a Firestore document response.
 * @param {object} firestoreValue - A value object from Firestore.
 * @returns {any} The corresponding standard JavaScript value or type.
 */
function fSrvConvertFirestoreTypesToJS(firestoreValue) {
  if (!firestoreValue) return firestoreValue;
  if (firestoreValue.stringValue !== undefined)
    return firestoreValue.stringValue;
  if (firestoreValue.integerValue !== undefined)
    return parseInt(firestoreValue.integerValue, 10);
  if (firestoreValue.doubleValue !== undefined)
    return parseFloat(firestoreValue.doubleValue);
  if (firestoreValue.booleanValue !== undefined)
    return firestoreValue.booleanValue;
  if (firestoreValue.nullValue !== undefined) return null;
  if (firestoreValue.timestampValue !== undefined)
    return new Date(firestoreValue.timestampValue);
  if (firestoreValue.arrayValue && firestoreValue.arrayValue.values) {
    return firestoreValue.arrayValue.values.map((element) =>
      fSrvConvertFirestoreTypesToJS(element)
    );
  }
  if (firestoreValue.mapValue && firestoreValue.mapValue.fields) {
    const jsObject = {};
    for (const key in firestoreValue.mapValue.fields) {
      jsObject[key] = fSrvConvertFirestoreTypesToJS(
        firestoreValue.mapValue.fields[key]
      );
    }
    return jsObject;
  }
  Logger.log(
    `fSrvConvertFirestoreTypesToJS: Encountered unexpected value structure: ${JSON.stringify(
      firestoreValue
    ).substring(0, 100)}... Returning as is.`
  );
  return firestoreValue;
} // End function fSrvConvertFirestoreTypesToJS

/** function fSrvUnpackFirestoreArrayTo2D
 * Purpose: Converts the Firestore array-of-row-objects format into a standard 2D JavaScript array.
 * Assumptions: Input is an array of objects like `[{ "row0": [...] }, { "row1": [...] }]`.
 * @param {object[]} firestoreArr - Array from Firestore after type conversion.
 * @returns {any[][]} A standard 2D JavaScript array.
 */
function fSrvUnpackFirestoreArrayTo2D(firestoreArr) {
  const funcName = "fSrvUnpackFirestoreArrayTo2D";
  if (!Array.isArray(firestoreArr)) {
    Logger.log(`${funcName}: Input is not an array. Returning empty array.`);
    return [];
  }
  const new2DArray = [];
  let maxRow = -1;
  let maxCols = 0;
  for (const rowObject of firestoreArr) {
    if (typeof rowObject !== "object" || rowObject === null) continue;
    const key = Object.keys(rowObject)[0];
    if (!key || !key.startsWith("row")) continue;
    const rowNum = parseInt(key.substring(3), 10);
    if (isNaN(rowNum)) continue;
    const rowData = rowObject[key];
    if (!Array.isArray(rowData)) {
      Logger.log(
        `${funcName}: Value for key ${key} is not an array. Skipping.`
      );
      continue;
    }
    new2DArray[rowNum] = rowData;
    if (rowNum > maxRow) maxRow = rowNum;
    if (rowData.length > maxCols) maxCols = rowData.length;
  }
  for (let r = 0; r <= maxRow; r++) {
    if (typeof new2DArray[r] === "undefined") {
      new2DArray[r] = Array(maxCols).fill("");
    } else {
      while (new2DArray[r].length < maxCols) {
        new2DArray[r].push("");
      }
    }
  }
  while (new2DArray.length <= maxRow) {
    new2DArray.push(Array(maxCols).fill(""));
  }
  if (maxRow === -1) {
    return [
      []
    ];
  }
  return new2DArray;
} // End function fSrvUnpackFirestoreArrayTo2D

/** function fSrvCheckAndLoadFirestoreGUIarrAs2D
 * Purpose: Checks Firestore for saved grid and notes data and returns it as 2D arrays.
 * Assumptions: The data is stored in a versioned, user-specific collection.
 * @param {object} gIndex - Object from the client with Email, CSID, and GameVer.
 * @returns {object} An object { success, firestoreArr?, customNotes?, message? }.
 */
function fSrvCheckAndLoadFirestoreGUIarrAs2D(gIndex) {
  const funcName = "fSrvCheckAndLoadFirestoreGUIarrAs2D";
  Logger.log(
    `${funcName}: Checking Firestore for data for User: ${gIndex.Email}, CS ID: ${gIndex.CSID}...`
  );
  if (!gIndex.GameVer || typeof gIndex.GameVer !== "string" || gIndex.GameVer.trim() === "") {
    const msg = "Invalid or missing Game Version provided.";
    Logger.log(`${funcName} Error: ${msg}`);
    return {
      success: false,
      message: msg
    };
  }
  if (!gIndex.Email || typeof gIndex.Email !== "string" || gIndex.Email.indexOf("@") === -1) {
    const msg = "Invalid User Email provided.";
    Logger.log(`${funcName} Error: ${msg}`);
    return {
      success: false,
      message: msg
    };
  }
  if (!gIndex.CSID || typeof gIndex.CSID !== "string") {
    const msg = "Invalid Character Sheet ID provided.";
    Logger.log(`${funcName} Error: ${msg}`);
    return {
      success: false,
      message: msg
    };
  }
  const firestore = fSrvGetFirestoreInstance();
  if (!firestore) {
    const msg = "Failed to initialize Firestore instance.";
    Logger.log(`${funcName} Error: ${msg}`);
    return {
      success: false,
      message: "Server configuration error (Firestore).",
    };
  }
  const gameVerMajor = String(gIndex.GameVer).trim().split(".")[0];
  const collectionPath = `v${gameVerMajor} ${gIndex.Email}`;
  const documentId = `Turbo_Game_${gIndex.CSID}`;
  const documentPath = `${collectionPath}/${documentId}`;
  Logger.log(`   -> Target Firestore Path: ${documentPath}`);
  try {
    const doc = firestore.getDocument(documentPath);
    if (!doc || !doc.fields || !doc.fields.gUIarr) {
      const msg = `Document not found or missing 'gUIarr' field at path: ${documentPath}.`;
      Logger.log(`   -> ${funcName}: ${msg}`);
      return {
        success: false,
        message: "No saved grid data found in Firestore for this character.",
      };
    }
    Logger.log(`   -> Document found. Processing fields...`);
    const arrDataRaw = doc.fields.gUIarr;
    const arrDataConverted = fSrvConvertFirestoreTypesToJS(arrDataRaw);
    if (!Array.isArray(arrDataConverted)) {
      throw new Error(
        "Invalid data type for gUIarr after conversion. Expected array."
      );
    }
    const unpackedArr = fSrvUnpackFirestoreArrayTo2D(arrDataConverted);
    Logger.log(`   -> ✅ Successfully fetched and unpacked gUI.arr data.`);
    let unpackedNotes = null;
    if (doc.fields.customNotes) {
      const notesDataRaw = doc.fields.customNotes;
      const notesDataConverted = fSrvConvertFirestoreTypesToJS(notesDataRaw);
      if (Array.isArray(notesDataConverted)) {
        unpackedNotes = fSrvUnpackFirestoreArrayTo2D(notesDataConverted);
        Logger.log(`   -> ✅ Successfully fetched and unpacked customNotes data.`);
      } else {
        Logger.log(
          `   -> ⚠️ Warning: 'customNotes' field found but is not a valid array. Ignoring.`
        );
      }
    } else {
      Logger.log(
        `   -> ℹ️ No 'customNotes' field found in document. This is normal for older saves.`
      );
    }
    return {
      success: true,
      firestoreArr: unpackedArr,
      customNotes: unpackedNotes,
    };
  } catch (e) {
    const safeErrorMessage = e.message?.includes("permission") ?
      "Permission denied accessing Firestore." :
      e.message?.includes("NOT_FOUND") ?
      "No saved data found in Firestore for this character." :
      `Server error during Firestore read: ${e.message || e}`;
    console.error(
      `Exception in ${funcName} for path ${documentPath}: ${e.message}\nStack: ${e.stack}`
    );
    Logger.log(
      `   -> ❌ Exception during Firestore read for ${gIndex.CSID}: ${safeErrorMessage}`
    );
    return {
      success: false,
      message: safeErrorMessage
    };
  }
} // End function fSrvCheckAndLoadFirestoreGUIarrAs2D

/** function fSrvSaveFullSheetTextAndTagsToFirestore
 * Purpose: Saves loaded sheet data and tags to Firestore, chunking if necessary.
 * Assumptions: The data object from fSrvLoadFullGoogleSheetAndTags is valid.
 * @param {object} gIndex - Contains Email, CSID, GameVer.
 * @param {string} workbookAbr - Abbreviation ('db', 'mycs', etc.).
 * @param {string} sheetName - The name of the sheet.
 * @param {object} data - Object containing { ColTags, RowTags, sheetText2D }.
 * @returns {object} An object { success: boolean, message?: string }.
 */
function fSrvSaveFullSheetTextAndTagsToFirestore(
  gIndex,
  workbookAbr,
  sheetName,
  data
) {
  const funcName = "fSrvSaveFullSheetTextAndTagsToFirestore";
  const MAX_CHUNK_SIZE_ESTIMATE = 500000;
  Logger.log(
    `${funcName}: Saving data for Workbook: "${workbookAbr}", Sheet: "${sheetName}", Version: ${gIndex?.GameVer}, Email: ${gIndex?.Email}, CSID: ${gIndex?.CSID}...`
  );
  const lowerWorkbookAbr = workbookAbr?.toLowerCase() || "";
  const trimmedSheetName = sheetName?.trim() || "";
  if (
    !data ||
    typeof data !== "object" ||
    !data.ColTags ||
    !data.RowTags ||
    !data.sheetText2D ||
    !Array.isArray(data.sheetText2D)
  ) {
    return {
      success: false,
      message: "Invalid data object provided (missing ColTags, RowTags, or sheetText2D array).",
    };
  }
  if (!trimmedSheetName) {
    return {
      success: false,
      message: "Invalid or empty sheetName provided."
    };
  }
  const firestore = fSrvGetFirestoreInstance();
  if (!firestore) {
    const msg = "Failed to initialize Firestore instance.";
    Logger.log(`${funcName} Error: ${msg}`);
    return {
      success: false,
      message: "Server configuration error (Firestore).",
    };
  }
  let baseCollectionName;
  let baseDocumentId;
  let documentPathBase;
  try {
    const pathInfo = fSrvCalcFirestorePath(
      workbookAbr,
      trimmedSheetName,
      gIndex
    );
    baseCollectionName = pathInfo.collectionName;
    baseDocumentId = pathInfo.documentId;
    documentPathBase = `${baseCollectionName}/${baseDocumentId}`;
    Logger.log(`   -> Base Firestore Path Calculated: ${documentPathBase}`);
  } catch (pathError) {
    Logger.log(
      `   -> ❌ Error determining Firestore path: ${pathError.message}`
    );
    return {
      success: false,
      message: pathError.message
    };
  }
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
  const chunks = [];
  let currentChunk = [];
  let currentChunkSizeEstimate = 0;
  Logger.log(
    `   -> Slicing data based on estimated size (Target: ${MAX_CHUNK_SIZE_ESTIMATE} bytes)...`
  );
  for (let i = 0; i < arrayOfRowObjects.length; i++) {
    const rowObject = arrayOfRowObjects[i];
    let rowObjectSizeEstimate = 0;
    try {
      rowObjectSizeEstimate = JSON.stringify(rowObject).length;
    } catch (e) {
      Logger.log(
        `   -> Warning: Could not estimate size for row object at index ${i}. Assuming small size (0). Error: ${e.message}`
      );
    }
    if (
      currentChunk.length > 0 &&
      currentChunkSizeEstimate + rowObjectSizeEstimate > MAX_CHUNK_SIZE_ESTIMATE
    ) {
      chunks.push(currentChunk);
      Logger.log(
        `      -> Chunk ${chunks.length} finalized with ${currentChunk.length} rows (Estimated size: ${currentChunkSizeEstimate} bytes).`
      );
      currentChunk = [rowObject];
      currentChunkSizeEstimate = rowObjectSizeEstimate;
    } else {
      currentChunk.push(rowObject);
      currentChunkSizeEstimate += rowObjectSizeEstimate;
    }
  }
  if (currentChunk.length > 0) {
    chunks.push(currentChunk);
    Logger.log(
      `      -> Chunk ${chunks.length} finalized with ${currentChunk.length} rows (Estimated size: ${currentChunkSizeEstimate} bytes).`
    );
  }
  const totalChunks =
    arrayOfRowObjects.length > 0 ? Math.max(1, chunks.length) : 0;
  Logger.log(`   -> Total data chunks determined: ${totalChunks}`);
  const metadataDocId = `${baseDocumentId}_metadata`;
  const metadataPath = `${baseCollectionName}/${metadataDocId}`;
  const metadataObject = {
    ColTags: data.ColTags,
    RowTags: data.RowTags,
    totalChunks: totalChunks,
    _lastUpdated: new Date(),
  };
  let metadataSaveSuccess = false;
  try {
    Logger.log(`   -> Saving Metadata Document to: ${metadataPath}`);
    firestore.updateDocument(metadataPath, metadataObject, false);
    metadataSaveSuccess = true;
    Logger.log(`      -> ✅ Successfully saved Metadata Document.`);
  } catch (e) {
    const errorMsg = `Failed to save Metadata Document (${metadataPath}): ${
      e.message || e
    }`;
    console.error(`${funcName} Error: ${errorMsg}\nStack: ${e.stack}`);
    Logger.log(`   -> ❌ ${errorMsg}`);
    return {
      success: false,
      message: errorMsg
    };
  }
  let allChunksSaved = true;
  if (totalChunks > 0) {
    for (let i = 0; i < totalChunks; i++) {
      const chunkIndex = i + 1;
      const chunkDocId = `${baseDocumentId}_${chunkIndex}of${totalChunks}`;
      const chunkPath = `${baseCollectionName}/${chunkDocId}`;
      const chunkData = {
        rowDataChunk: chunks[i],
        _lastUpdated: new Date(),
      };
      try {
        Logger.log(
          `   -> Saving Data Chunk ${chunkIndex}/${totalChunks} to: ${chunkPath}`
        );
        firestore.updateDocument(chunkPath, chunkData, false);
        Logger.log(
          `      -> ✅ Successfully saved Data Chunk ${chunkIndex}/${totalChunks}.`
        );
      } catch (e) {
        const errorMsg = `Failed to save Data Chunk ${chunkIndex}/${totalChunks} (${chunkPath}): ${
          e.message || e
        }`;
        console.error(`${funcName} Error: ${errorMsg}\nStack: ${e.stack}`);
        Logger.log(`   -> ❌ ${errorMsg}`);
        allChunksSaved = false;
      }
    }
  } else {
    Logger.log(`   -> No data chunks to save (source data likely empty).`);
  }
  if (metadataSaveSuccess && allChunksSaved) {
    Logger.log(
      `   -> ✅ Successfully saved Metadata and all ${totalChunks} Data Chunk(s).`
    );
    return {
      success: true
    };
  } else {
    const finalMessage = `Firestore save partially failed. Metadata saved: ${metadataSaveSuccess}. All data chunks saved: ${allChunksSaved}. Check logs for details.`;
    Logger.log(`   -> ❌ ${finalMessage}`);
    return {
      success: false,
      message: finalMessage
    };
  }
} // End function fSrvSaveFullSheetTextAndTagsToFirestore

/** function fSrvCalcFirestorePath
 * Purpose: Determines the Firestore collection and document ID for a given resource.
 * Assumptions: Handles versioning for DB and user-specific paths.
 * @param {string} workbookAbr - Abbreviation ('db', 'mastercs', 'mycs', etc.).
 * @param {string} sheetName - The name of the sheet.
 * @param {object} gIndex - Object with GameVer, Email, and CSID.
 * @returns {object} An object { collectionName: string, documentId: string }.
 * @throws {Error} If inputs are invalid or workbook abbreviation is unsupported.
 */
function fSrvCalcFirestorePath(workbookAbr, sheetName, gIndex) {
  const funcName = "fSrvCalcFirestorePath";
  if (!workbookAbr || typeof workbookAbr !== "string") {
    throw new Error(`${funcName}: Invalid or missing workbookAbr provided.`);
  }
  if (!sheetName || typeof sheetName !== "string" || sheetName.trim() === "") {
    throw new Error(`${funcName}: Invalid or empty sheetName provided.`);
  }
  const lowerWorkbookAbr = workbookAbr.toLowerCase();
  const trimmedSheetName = sheetName.trim();
  const validWorkbooks = ["db", "mastercs", "masterkl", "mycs", "mykl"];
  if (!validWorkbooks.includes(lowerWorkbookAbr)) {
    throw new Error(
      `${funcName}: Unsupported workbook abbreviation: "${workbookAbr}"`
    );
  }
  let collectionName = "";
  let documentId = "";
  switch (lowerWorkbookAbr) {
    case "db":
    case "mastercs":
    case "masterkl":
      if (
        !gIndex.GameVer ||
        typeof gIndex.GameVer !== "string" ||
        gIndex.GameVer.trim() === ""
      ) {
        throw new Error(
          `${funcName}: Game Version is required for workbook type '${workbookAbr}'.`
        );
      }
      const gameVerMajorForDB = String(gIndex.GameVer).trim().split(".")[0];
      if (lowerWorkbookAbr === "db") {
        collectionName = `v${gameVerMajorForDB} DB`;
        documentId = trimmedSheetName;
      } else if (lowerWorkbookAbr === "mastercs") {
        collectionName = "MasterCS";
        documentId = `v${gIndex.GameVer.trim()} ${trimmedSheetName}`;
      } else {
        collectionName = "MasterKL";
        documentId = `v${gIndex.GameVer.trim()} ${trimmedSheetName}`;
      }
      break;
    case "mycs":
    case "mykl":
      if (
        !gIndex.GameVer ||
        typeof gIndex.GameVer !== "string" ||
        gIndex.GameVer.trim() === ""
      ) {
        throw new Error(
          `${funcName}: Game Version is required for workbook type '${workbookAbr}'.`
        );
      }
      if (
        !gIndex.Email ||
        typeof gIndex.Email !== "string" ||
        gIndex.Email.indexOf("@") === -1
      ) {
        throw new Error(
          `${funcName}: User Email is required and must be valid for workbook type '${workbookAbr}'.`
        );
      }
      if (!gIndex.CSID || typeof gIndex.CSID !== "string") {
        throw new Error(
          `${funcName}: Character Sheet ID is required for workbook type '${workbookAbr}'.`
        );
      }
      const gameVerMajor = String(gIndex.GameVer).trim().split(".")[0];
      collectionName = `v${gameVerMajor} ${gIndex.Email}`;
      if (lowerWorkbookAbr === "mycs") {
        documentId = `MyCS_${trimmedSheetName}_${gIndex.CSID}`;
      } else {
        documentId = `MyKL_${trimmedSheetName}_OfMyCS_${gIndex.CSID}`;
      }
      break;
  }
  if (!collectionName || !documentId) {
    throw new Error(
      `${funcName}: Failed to determine collectionName or documentId for workbook '${workbookAbr}'.`
    );
  }
  return {
    collectionName,
    documentId
  };
} // End function fSrvCalcFirestorePath

/** function fSrvGetRequiredCaches
 * Purpose: Loads a list of required caches from Firestore in a single bulk operation.
 * Assumptions: Self-heals by reading from Sheets if a cache is missing from Firestore.
 * @param {object[]} requiredCaches - An array of cache definitions, e.g., [{ key: 'dbAbilitiesFSData', label: 'DB/Abilities' }].
 * @param {object} gIndex - The standard gIndex object from the client.
 * @returns {object} A result object { success, caches, message? }.
 */
function fSrvGetRequiredCaches(requiredCaches, gIndex) {
  const funcName = "fSrvGetRequiredCaches";
  Logger.log(
    `${funcName}: Received request to bulk-load ${requiredCaches.length} caches.`
  );
  if (!Array.isArray(requiredCaches) || requiredCaches.length === 0) {
    return {
      success: false,
      caches: {},
      message: "Invalid or empty cache list provided.",
    };
  }
  const loadedCaches = {};
  let overallSuccess = true;
  for (const cacheInfo of requiredCaches) {
    const {
      key,
      label
    } = cacheInfo;
    const [workbookAbr, sheetName] = label.split("/");
    try {
      Logger.log(`   -> ${funcName}: Processing cache '${key}' (${label})...`);
      const cacheExists = fSrvVerifyFirestorePathExists(
        workbookAbr,
        sheetName,
        gIndex
      );
      if (cacheExists) {
        Logger.log(`      -> Cache exists. Reading from Firestore.`);
        const response = fSrvGetFirestoreFSData(workbookAbr, sheetName, gIndex);
        if (response.success) {
          loadedCaches[key] = response.FSData;
        } else {
          throw new Error(response.message || "Failed to read existing cache.");
        }
      } else {
        Logger.log(
          `      -> Cache NOT found. Self-healing: Reading from Sheet...`
        );
        const sheetData = fSrvLoadFullGoogleSheetAndTags(
          workbookAbr,
          sheetName,
          gIndex.CSID
        );
        Logger.log(`      -> Self-healing: Saving '${key}' to Firestore...`);
        fSrvSaveFullSheetTextAndTagsToFirestore(
          gIndex,
          workbookAbr,
          sheetName,
          sheetData
        );
        const response = fSrvGetFirestoreFSData(workbookAbr, sheetName, gIndex);
        if (response.success) {
          loadedCaches[key] = response.FSData;
        } else {
          throw new Error(
            response.message || "Failed to read cache after self-healing."
          );
        }
      }
    } catch (e) {
      Logger.log(
        `   -> ❌ ${funcName}: CRITICAL FAILURE processing cache '${key}'. Error: ${e.message}`
      );
      console.error(`Error in ${funcName} for ${key}: ${e.stack}`);
      overallSuccess = false;
    }
  }
  Logger.log(
    `${funcName}: Finished bulk load. Returning ${
      Object.keys(loadedCaches).length
    } of ${requiredCaches.length} requested caches.`
  );
  return {
    success: overallSuccess,
    caches: loadedCaches
  };
} // End function fSrvGetRequiredCaches

/** function fSrvGetFirestoreFSData
 * Purpose: Reads and reassembles data from potentially chunked Firestore documents.
 * Assumptions: Data was saved using fSrvSaveFullSheetTextAndTagsToFirestore.
 * @param {string} workbookAbr - Workbook abbreviation ('db', 'mycs', etc.).
 * @param {string} sheetName - The sheet name associated with the data.
 * @param {object} gIndex - Object containing CSID, GameVer, Email.
 * @returns {object} On success: { success, FSData }, on failure: { success, message }.
 */
function fSrvGetFirestoreFSData(workbookAbr, sheetName, gIndex) {
  const funcName = "fSrvGetFirestoreFSData";
  Logger.log(
    `${funcName}: Reading document(s) for Workbook: "${workbookAbr}", Sheet: "${sheetName}", Ver: ${gIndex?.GameVer}, Email: ${gIndex?.Email}, CSID: ${gIndex?.CSID}...`
  );
  let firestore;
  let baseCollectionName;
  let baseDocumentId;
  let metadataPath;
  let absoluteColTagMap = {};
  let absoluteRowTagMap = {};
  try {
    firestore = fSrvGetFirestoreInstance();
    if (!firestore) {
      return {
        success: false,
        message: "Server configuration error (Firestore).",
      };
    }
    try {
      const pathInfo = fSrvCalcFirestorePath(workbookAbr, sheetName, gIndex);
      baseCollectionName = pathInfo.collectionName;
      baseDocumentId = pathInfo.documentId;
    } catch (pathError) {
      Logger.log(
        `   -> ❌ Error determining Firestore path: ${pathError.message}`
      );
      return {
        success: false,
        message: pathError.message
      };
    }
    metadataPath = `${baseCollectionName}/${baseDocumentId}_metadata`;
    Logger.log(`   -> Target Metadata Path: ${metadataPath}`);
    let metadataDoc;
    try {
      metadataDoc = firestore.getDocument(metadataPath);
    } catch (e) {
      const isNotFoundError = e.message?.toUpperCase().includes("NOT_FOUND");
      const errorMsg = isNotFoundError ?
        `Metadata document not found at path: ${metadataPath}. Data may be missing or not yet saved.` :
        `Error fetching metadata document (${metadataPath}): ${
          e.message || e
        }`;
      Logger.log(`   -> ${funcName}: ${errorMsg}`);
      return {
        success: false,
        message: errorMsg
      };
    }
    if (!metadataDoc || !metadataDoc.fields) {
      const msg = `Metadata document not found or empty at path: ${metadataPath}.`;
      Logger.log(`   -> ${funcName}: ${msg}`);
      return {
        success: false,
        message: msg
      };
    }
    const colTagsRaw = metadataDoc.fields.ColTags;
    const rowTagsRaw = metadataDoc.fields.RowTags;
    const totalChunksRaw = metadataDoc.fields.totalChunks;
    if (
      !colTagsRaw ||
      typeof colTagsRaw.mapValue === "undefined" ||
      !rowTagsRaw ||
      typeof rowTagsRaw.mapValue === "undefined" ||
      !totalChunksRaw ||
      typeof totalChunksRaw.integerValue === "undefined"
    ) {
      const msg =
        "Invalid metadata document structure found (missing/invalid ColTags, RowTags, or totalChunks).";
      Logger.log(`   -> ${funcName} Error: ${msg}`);
      return {
        success: false,
        message: msg
      };
    }
    absoluteColTagMap = fSrvConvertFirestoreTypesToJS(colTagsRaw);
    absoluteRowTagMap = fSrvConvertFirestoreTypesToJS(rowTagsRaw);
    const totalChunks = parseInt(totalChunksRaw.integerValue, 10);
    if (
      typeof absoluteColTagMap !== "object" ||
      absoluteColTagMap === null ||
      Array.isArray(absoluteColTagMap) ||
      typeof absoluteRowTagMap !== "object" ||
      absoluteRowTagMap === null ||
      Array.isArray(absoluteRowTagMap) ||
      isNaN(totalChunks) ||
      totalChunks < 0
    ) {
      const msg =
        "Invalid data types found in metadata after conversion (ColTags/RowTags not objects, or totalChunks not integer >= 0).";
      Logger.log(`   -> ${funcName} Error: ${msg}`);
      return {
        success: false,
        message: msg
      };
    }
    Logger.log(
      `   -> Metadata validated. Total Chunks: ${totalChunks}. ColTags: ${
        Object.keys(absoluteColTagMap).length
      }, RowTags: ${Object.keys(absoluteRowTagMap).length}`
    );
    if (totalChunks === 0) {
      Logger.log(`   -> Total chunks is 0. Returning empty data structure.`);
      return {
        success: true,
        FSData: {
          colTagsMap: absoluteColTagMap,
          rowTagsMap: absoluteRowTagMap,
          text: [
            []
          ],
        },
      };
    }
    const fetchedChunkDocs = [];
    const missingChunks = [];
    Logger.log(`   -> Attempting to fetch ${totalChunks} data chunk(s)...`);
    for (let i = 1; i <= totalChunks; i++) {
      const chunkDocId = `${baseDocumentId}_${i}of${totalChunks}`;
      const chunkPath = `${baseCollectionName}/${chunkDocId}`;
      try {
        const chunkDoc = firestore.getDocument(chunkPath);
        if (chunkDoc && chunkDoc.fields && chunkDoc.fields.rowDataChunk) {
          fetchedChunkDocs.push(chunkDoc);
        } else {
          missingChunks.push(i);
          Logger.log(
            `      -> ❌ Failed to fetch or find valid 'rowDataChunk' in chunk ${i}/${totalChunks} at ${chunkPath}.`
          );
        }
      } catch (e) {
        const isNotFoundError = e.message?.toUpperCase().includes("NOT_FOUND");
        Logger.log(
          `      -> ❌ Exception fetching chunk ${i}/${totalChunks} at ${chunkPath}: ${
            e.message || e
          }${isNotFoundError ? " (NOT_FOUND)" : ""}`
        );
        missingChunks.push(i);
      }
    }
    if (missingChunks.length > 0) {
      const errorMsg = `Failed to load all required data chunks. Missing chunk(s): ${missingChunks.join(
        ", "
      )} of ${totalChunks}. Data is incomplete.`;
      Logger.log(`   -> ${funcName} Error: ${errorMsg}`);
      return {
        success: false,
        message: errorMsg
      };
    }
    Logger.log(`   -> Successfully fetched all ${totalChunks} data chunk(s).`);
    const combinedRowObjects = [];
    Logger.log(`   -> Reassembling data from chunks...`);
    for (let i = 0; i < fetchedChunkDocs.length; i++) {
      const chunkDoc = fetchedChunkDocs[i];
      const chunkIndex = i + 1;
      const rowDataChunkRaw = chunkDoc.fields.rowDataChunk;
      const rowDataChunkConverted =
        fSrvConvertFirestoreTypesToJS(rowDataChunkRaw);
      if (!Array.isArray(rowDataChunkConverted)) {
        const errorMsg = `Invalid rowDataChunk format found in chunk ${chunkIndex} after conversion (expected array).`;
        Logger.log(`   -> ${funcName} Error: ${errorMsg}`);
        return {
          success: false,
          message: errorMsg
        };
      }
      combinedRowObjects.push(...rowDataChunkConverted);
    }
    Logger.log(
      `   -> Reassembled ${combinedRowObjects.length} total row objects.`
    );
    const fullData2D = fSrvUnpackFirestoreArrayTo2D(combinedRowObjects);
    const numRowsFinal = fullData2D.length;
    const numColsFinal = fullData2D[0]?.length || 0;
    Logger.log(
      `   -> Unpacked reassembled data into final 2D array (${numRowsFinal}x${numColsFinal}).`
    );
    const assembledFSDataObject = {
      colTagsMap: absoluteColTagMap,
      rowTagsMap: absoluteRowTagMap,
      text: fullData2D,
    };
    Logger.log(
      `   -> ✅ Successfully read and formatted sliced Firestore data.`
    );
    return {
      success: true,
      FSData: assembledFSDataObject,
    };
  } catch (e) {
    const safeErrorMessage = e.message?.includes("permission") ?
      "Permission denied accessing Firestore." :
      `Server error during Firestore read/process: ${e.message || e}`;
    console.error(
      `Exception caught in ${funcName} accessing path ${
        metadataPath || "Unknown"
      }: ${e.message}\nStack: ${e.stack}`
    );
    Logger.log(
      `   -> ❌ Exception during Firestore read/process for ${
        metadataPath || "Unknown"
      }: ${safeErrorMessage}`
    );
    return {
      success: false,
      message: safeErrorMessage
    };
  }
} // End function fSrvGetFirestoreFSData

/** function fSrvVerifyFirestorePathExists
 * Purpose: Checks if the metadata Firestore document exists for a given resource.
 * Assumptions: A document existing means its `updateTime` property is present.
 * @param {string} workbookAbr - Workbook abbreviation ('db', 'mycs', etc.).
 * @param {string} sheetName - The sheet name associated with the data.
 * @param {object} gIndex - Object with GameVer, Email, and CSID.
 * @returns {boolean} True if the metadata document exists, false otherwise.
 */
function fSrvVerifyFirestorePathExists(workbookAbr, sheetName, gIndex) {
  const funcName = "fSrvVerifyFirestorePathExists";
  let firestore;
  let metadataPath;
  try {
    firestore = fSrvGetFirestoreInstance();
    if (!firestore) {
      Logger.log(
        `   -> ${funcName}: Firestore initialization failed. Cannot verify path.`
      );
      return false;
    }
    let baseCollectionName;
    let baseDocumentId;
    try {
      const pathInfo = fSrvCalcFirestorePath(workbookAbr, sheetName, gIndex);
      baseCollectionName = pathInfo.collectionName;
      baseDocumentId = pathInfo.documentId;
    } catch (pathError) {
      Logger.log(
        `   -> ${funcName}: Error calculating base path: ${pathError.message}. Assuming path does not exist.`
      );
      return false;
    }
    metadataPath = `${baseCollectionName}/${baseDocumentId}_metadata`;
    Logger.log(
      `   -> Calculated Firestore Metadata Path to check: ${metadataPath}`
    );
    const doc = firestore.getDocument(metadataPath);
    if (doc && doc.updateTime) {
      Logger.log(
        `   -> Metadata document found at path: ${metadataPath}. Exists: true.`
      );
      return true;
    } else {
      Logger.log(
        `   -> Metadata document NOT found at path: ${metadataPath}. Exists: false.`
      );
      return false;
    }
  } catch (e) {
    const isNotFoundError =
      e.message && e.message.toUpperCase().includes("NOT_FOUND");
    if (isNotFoundError) {
      Logger.log(
        `   -> ${funcName}: Explicit NOT_FOUND error for metadata path ${metadataPath}. Exists: false.`
      );
      return false;
    } else {
      console.error(
        `Exception caught in ${funcName} accessing metadata path ${
          metadataPath || "Unknown"
        }: ${e.message}\nStack: ${e.stack}`
      );
      Logger.log(
        `   -> ❌ Exception during Firestore metadata check for ${
          metadataPath || "Unknown"
        }: ${e.message}. Assuming path does not exist.`
      );
      return false;
    }
  }
} // End function fSrvVerifyFirestorePathExists
