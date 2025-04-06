// DB

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  Menu  (end initialize)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// fDBCreateMenu //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Create MetaScape menu
function fDBCreateMenu() {
  SpreadsheetApp.getUi()
    .createMenu("*** GM Screen")
    .addItem("Monster Refresh", "fDBMenuMonRefresh")
    .addItem("Clear Roll Logs", "fDBMenuClearRollLogs")
    .addToUi();
  SpreadsheetApp.getUi()
    .createMenu("Kits")
    .addItem("Load Kit", "fDBMenuLoadKit")
    .addItem("Save Kit", "fDBMenuSaveKit")
    .addToUi();
  SpreadsheetApp.getUi()
    .createMenu("Abilities")
    .addItem("Remove old <Element> IDs from Rules", "fDBMenuCullBookIDs")
    .addItem("Save v1:, v2:, etc To Rules", "fDBMenuSaveVersions")
    .addItem(`Update <Element>'s Notes From Rules`, "fDBMenuUpdateElementNotes")
    .addToUi();
  SpreadsheetApp.getUi()
    .createMenu("System ***")
    .addItem("1 - Authorize Script", "fDBMenuAuthorize")
    .addSeparator()
    .addItem("Refresh Menus", "fDBMenuRefreshAll")
    .addToUi();
  if (gGetVal("db", "Data", "designer", "Val") === true) {
    SpreadsheetApp.getUi()
      .createMenu("DESIGNER")
      .addItem("Name_ID Sync", "fDBMenuName_IDSync")
      .addSeparator()
      .addItem("Hide All", "fDBMenuHideAll")
      .addItem("Un-Hide All", "fDBMenuUn_HideAll")
      .addToUi();
  }
} // end fDBCreateMenu

// Menu Functions //////////////////////////////////////////////////////////////////////////////////////////////////
// GMScreen Menu
function fDBMenuMonRefresh() {
  fDBRunMenuOrButton("MonRefresh");
}
function fDBMenuClearRollLogs() {
  fDBRunMenuOrButton("ClearRollLogs");
}
// Kit Menu
function fDBMenuLoadKit() {
  fDBRunMenuOrButton("LoadKit");
}
function fDBMenuSaveKit() {
  fDBRunMenuOrButton("SaveKit");
}
// Abilities Menu
function fDBMenuSaveVersions() {
  fDBRunMenuOrButton("SaveVersions");
}
function fDBMenuCullBookIDs() {
  fDBRunMenuOrButton("CullBookIDs");
}
function fDBMenuUpdateElementNotes() {
  fDBRunMenuOrButton("UpdateNotes");
}
function fDBMenuDelBogusKitBases() {
  fDBRunMenuOrButton("DelBogusKitBases");
}
function fDBMenuPushKitToAbil() {
  fDBRunMenuOrButton("PushKitToAbil");
}
function fDBMenuPopulateKitFromAbil() {
  fDBRunMenuOrButton("PopulateKitFromAbil");
}
// System Menu
function fDBMenuAuthorize() {
  fDBRunMenuOrButton("Authorize");
}
function fDBMenuRefreshAll() {
  fDBRunMenuOrButton("RefreshMenu");
}
// Designer Menu
function fDBMenuName_IDSync() {
  fDBRunMenuOrButton("MenuName_IDSync");
}
function fDBMenuHideAll() {
  fDBRunMenuOrButton("HideAll");
}
function fDBMenuUn_HideAll() {
  fDBRunMenuOrButton("Un_HideAll");
}
// End Menu Functions

// Image Button Functions //////////////////////////////////////////////////////////////////////////////////////////////////
function fDBButtonMon() {
  fDBRunMenuOrButton("MonRefresh");
}
function fDBButtonOpen() {
  fDBRunMenuOrButton("LoadKit");
}
function fDBButtonSave() {
  fDBRunMenuOrButton("SaveKit");
}
// End Button Functions

// fDBRunMenuOrButton //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> To run all menu & button choices inside a try-catch-error
function fDBRunMenuOrButton(menuChoice) {
  try {
    switch (menuChoice) {
      // GM Screen
      case "MonRefresh":
        fDBMonRefresh();
        break;
      case "ClearRollLogs":
        fDBClearRollLogs();
        break;
      // Kits
      case "LoadKit":
        fDBLoadKit();
        break;
      case "SaveKit":
        fDBSaveKit();
        break;
      // Abilities
      case "SaveVersions":
        fDBSaveVersionsToBooks();
        break;
      case "CullBookIDs":
        fDBCullOldIDsFromBooks();
        break;
      case "UpdateNotes":
        fDBUpdateNotes();
        break;
      case "DelBogusKitBases":
        fFixBogusAbilBases();
        break;
      case "PushKitToAbil":
        fPushKitToAbil();
        break;
      case "PopulateKitFromAbil":
        fPopulateKitFromAbil();
        break;
      // System Menu
      case "Authorize":
        SpreadsheetApp.getUi().alert(
          `AUTHORIZED`,
          `Script Authorized!`,
          SpreadsheetApp.getUi().ButtonSet.OK
        );
        break;
      case "RefreshMenu":
        fDBCreateMenu();
        break;
      // Designer Menu
      case "MenuName_IDSync":
        fDBName_IDSync();
        break;
      case "HideAll":
        gHideAll("db");
        break;
      case "Un_HideAll":
        gUn_HideAll("db");
        break;
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert(error); // NOTE: an error of End or end will simply end the program.
  }
} // end fDBRunMenuOrButton

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  GM Screen  (end g. data)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// fDBMonRefresh //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Refreshes all Monsters on <GM Screen>
function fDBMonRefresh() {
  // Get <GMScreen> info
  const sheetRef = gSheetRef("db", "gmscreen");
  const monArr = gArr("db", "gmscreen");
  const monMultAll_R = gKeyR("db", "gmscreen", "MonMultAll");
  const monColMult_R = gKeyR("db", "gmscreen", "MonColMult");
  const monMultAll_C = gHeaderC("db", "gmscreen", "MonMultAll");
  const monTblFirst_R = gKeyR("db", "gmscreen", "MonTblFirst_R");
  const monTblLast_R = gKeyR("db", "gmscreen", "MonTblLast_R");
  const monTblPremadeFirst_R = monTblFirst_R;
  const monTblPremadeLast_R = gKeyR("db", "gmscreen", "PremadeMonLast_C");
  const monTblCustFirst_R = gKeyR("db", "gmscreen", "CustMonFirst_R");
  const monTblCustLast_R = monTblLast_R;
  const monID_C = gHeaderC("db", "gmscreen", "MonID");
  const monIndMM_C = gHeaderC("db", "gmscreen", "MM");
  const monActive_C = gHeaderC("db", "gmscreen", "ActiveMonTF");
  const monNameID_C = gHeaderC("db", "gmscreen", "MonName");
  const monCustName_C = gHeaderC("db", "gmscreen", "MonCustName");
  const monNish_C = gHeaderC("db", "gmscreen", "Nish");
  const monMR_C = gHeaderC("db", "gmscreen", "MR");
  const monVit_C = gHeaderC("db", "gmscreen", "Vit");
  const monAtk_C = gHeaderC("db", "gmscreen", "Atk");
  const monDmg_C = gHeaderC("db", "gmscreen", "Dmg");
  const monDef_C = gHeaderC("db", "gmscreen", "Def");
  const monAr_C = gHeaderC("db", "gmscreen", "Ar");
  const monPic_C = gHeaderC("db", "gmscreen", "Pic");
  const monNotes_C = gHeaderC("db", "gmscreen", "Notes");
  const monMultAll = monArr[monMultAll_R][monMultAll_C];
  const nishColMult =
    isNaN(monArr[monColMult_R][monNish_C]) ||
    monArr[monColMult_R][monNish_C] <= 0
      ? 1
      : monArr[monColMult_R][monNish_C];
  const mrColMult =
    isNaN(monArr[monColMult_R][monMR_C]) || monArr[monColMult_R][monMR_C] <= 0
      ? 1
      : monArr[monColMult_R][monMR_C];
  const vitColMult =
    isNaN(monArr[monColMult_R][monVit_C]) || monArr[monColMult_R][monVit_C] <= 0
      ? 1
      : monArr[monColMult_R][monVit_C];
  const atkColMult =
    isNaN(monArr[monColMult_R][monAtk_C]) || monArr[monColMult_R][monAtk_C] <= 0
      ? 1
      : monArr[monColMult_R][monAtk_C];
  const dmgColMult =
    isNaN(monArr[monColMult_R][monDmg_C]) || monArr[monColMult_R][monDmg_C] <= 0
      ? 1
      : monArr[monColMult_R][monDmg_C];
  const defColMult =
    isNaN(monArr[monColMult_R][monDef_C]) || monArr[monColMult_R][monDef_C] <= 0
      ? 1
      : monArr[monColMult_R][monDef_C];
  const arColMult =
    isNaN(monArr[monColMult_R][monAr_C]) || monArr[monColMult_R][monAr_C] <= 0
      ? 1
      : monArr[monColMult_R][monAr_C];

  if (isNaN(monMultAll) || monMultAll <= 0) {
    throw new Error(`MonMult must be a number and it must be greater than 0`);
  }

  // Get <Monsters> data
  const dataNish_C = gHeaderC("db", "Monsters", "Nish");
  const dataMR_C = gHeaderC("db", "Monsters", "MR");
  const dataVit_C = gHeaderC("db", "Monsters", "Vit");
  const dataAtk_C = gHeaderC("db", "Monsters", "Atk");
  const dataDmg_C = gHeaderC("db", "Monsters", "Dmg");
  const dataDef_C = gHeaderC("db", "Monsters", "Def");
  const dataAr_C = gHeaderC("db", "Monsters", "Ar");
  const dataPic_C = gHeaderC("db", "Monsters", "Pic");
  const dataNotes_C = gHeaderC("db", "Monsters", "Notes");

  // Loop r from monTblPremadeFirst_R to monTblPremadeLast_R and do this...
  for (let r = monTblPremadeFirst_R; r <= monTblPremadeLast_R; r++) {
    if (monArr[r][monNameID_C] === "") {
      monArr[r][monIndMM_C] = "";
      monArr[r][monActive_C] = false;
      gFillArraySection(monArr, r, r, monCustName_C, monNotes_C, "");

      // Set the font color of all cells in the specified range to black
      sheetRef
        .getRange(r + 1, monNish_C + 1, 1, monAr_C - monNish_C + 1)
        .setFontWeight("normal");
    } else {
      const monID = monArr[r][monID_C];
      const monIndMM = monArr[r][monIndMM_C];
      if (fDBTextisNonBold(sheetRef, monArr, r, monNish_C))
        monArr[r][monNish_C] = Math.round(
          gGetVal("db", "Monsters", monID, dataNish_C) *
            nishColMult *
            (monIndMM === "" ? monMultAll : monIndMM)
        );
      if (fDBTextisNonBold(sheetRef, monArr, r, monMR_C))
        monArr[r][monMR_C] = Math.round(
          +gGetVal("db", "Monsters", monID, dataMR_C) * mrColMult
        );
      if (fDBTextisNonBold(sheetRef, monArr, r, monVit_C))
        monArr[r][monVit_C] = Math.round(
          gGetVal("db", "Monsters", monID, dataVit_C) *
            vitColMult *
            (monIndMM === "" ? monMultAll : monIndMM)
        );
      if (fDBTextisNonBold(sheetRef, monArr, r, monAtk_C))
        monArr[r][monAtk_C] = Math.round(
          gGetVal("db", "Monsters", monID, dataAtk_C) *
            atkColMult *
            (monIndMM === "" ? monMultAll : monIndMM)
        );
      if (fDBTextisNonBold(sheetRef, monArr, r, monDmg_C))
        monArr[r][monDmg_C] = Math.round(
          gGetVal("db", "Monsters", monID, dataDmg_C) *
            dmgColMult *
            (monIndMM === "" ? monMultAll : monIndMM)
        );
      if (fDBTextisNonBold(sheetRef, monArr, r, monDef_C))
        monArr[r][monDef_C] = Math.round(
          gGetVal("db", "Monsters", monID, dataDef_C) *
            defColMult *
            (monIndMM === "" ? monMultAll : monIndMM)
        );
      if (fDBTextisNonBold(sheetRef, monArr, r, monAr_C))
        monArr[r][monAr_C] = Math.round(
          gGetVal("db", "Monsters", monID, dataAr_C) *
            arColMult *
            (monIndMM === "" ? monMultAll : monIndMM)
        );
      monArr[r][monPic_C] = gGetVal("db", "Monsters", monID, dataPic_C);
      monArr[r][monNotes_C] = gGetVal("db", "Monsters", monID, dataNotes_C);
    }
  }

  gSaveArraySectionToSheet(
    sheetRef,
    monArr,
    monTblPremadeFirst_R,
    monTblPremadeLast_R,
    monIndMM_C,
    monNotes_C
  );

  // Loop r from monTblCustFirst_R to monTblCustLast_R and do this...
  for (let r = monTblCustFirst_R; r <= monTblCustLast_R; r++) {
    if (monArr[r][monCustName_C] === "") {
      monArr[r][monActive_C] = false;
      gFillArraySection(monArr, r, r, monCustName_C, monAr_C, "");
    }
  }

  gSaveArraySectionToSheet(
    sheetRef,
    monArr,
    monTblCustFirst_R,
    monTblCustLast_R,
    monActive_C,
    monAr_C
  );
} // end fDBMonRefresh

// fDBTextisNonBold //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: returns true if array cell's sheet location is nonBold or empty, if empty, makes sure to set the cell to nonBold
function fDBTextisNonBold(ref, arr, r, c) {
  // Store the cell range in a variable to minimize calls to the sheet
  const cellRng = ref.getRange(r + 1, c + 1);

  // Check if cell is nonBold, if so return true (also captures empty AND nonBold as it should)
  if (cellRng.getFontWeight() === "normal") return true;

  // If the cell is empty and bold (which it must be due to above) then unbold it and return true
  if (cellRng.getValue() === "") {
    cellRng.setFontWeight("normal");
    return true;
  }

  // At this point, the cell has to be non-empty and bold
  return false;
} // end fDBTextisNonBold

// fDBClearRollLogs //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Clear's all Roll Logs
function fDBClearRollLogs(ss) {
  // Set <GMScreen> data
  const gmScreenRef = gSheetRef("db", "GMScreen");
  const gmScreenArr = gArr("db", "GMScreen");
  const gmScreenFirst_R = gKeyR("db", "GMScreen", "URL");
  const gmScreenLast_R = gKeyR("db", "GMScreen", "Log");
  const gmScreenFirst_C = gHeaderC("db", "GMScreen", "Slot1");
  const gmScreenLast_C = gHeaderC("db", "GMScreen", "Slot9");

  gFillArraySection(
    gmScreenArr,
    gmScreenFirst_R,
    gmScreenLast_R,
    gmScreenFirst_C,
    gmScreenLast_C,
    ""
  );
  gSaveArraySectionToSheet(
    gmScreenRef,
    gmScreenArr,
    gmScreenFirst_R,
    gmScreenLast_R,
    gmScreenFirst_C,
    gmScreenLast_C
  );
} // End fDBClearRollLogs

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  Kits  (end g. GM Screen)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// fDBLoadKit //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Loads the selected kit into the <KitEditor>
function fDBLoadKit() {
  const kitEditorRef = gSheetRef("db", "KitEditor");
  const kitEditorArr = gArr("db", "KitEditor");
  const kitEditorDataFirst_R = gDataFirst_R("db", "KitEditor");
  const kitEditorDataLast_R = gDataLast_R("db", "KitEditor");
  const kitEditorKitNameID_R = gKeyR("db", "KitEditor", "KitNameID");
  const kitEditorMorphStr_R = gKeyR("db", "KitEditor", "MorphStr");
  const kitEditorAbilityNameID_C = gHeaderC("db", "KitEditor", "AbilNameID");
  const kitEditorKitMorphStr_C = gHeaderC("db", "KitEditor", "MorphString");
  const kitEditorNameID_C = gHeaderC("db", "KitEditor", "KitNameID");
  const kitEditorOtherMorph_C = gHeaderC("db", "KitEditor", "OtherMorph");
  const kitEditorMorph1_C = gHeaderC("db", "KitEditor", "Morph1");
  const kitEditorMorph2_C = gHeaderC("db", "KitEditor", "Morph2");

  const kitNameID = kitEditorArr[kitEditorKitNameID_R][kitEditorNameID_C];
  const kitID = gTestID("db", "Kits", kitNameID);
  if (!kitID)
    throw new Error(
      `In fDBLoadKit, the selected kit "${kitNameID}" does not exist in <Kits>`
    );

  gFillArraySection(
    kitEditorArr,
    kitEditorDataFirst_R,
    kitEditorDataLast_R,
    kitEditorAbilityNameID_C,
    kitEditorMorph2_C,
    ""
  );

  const kitMorphStr = gGetVal("db", "Kits", kitID, "MorphString");
  gSaveVal(
    "db",
    "KitEditor",
    kitEditorMorphStr_R,
    kitEditorKitMorphStr_C,
    kitMorphStr
  );

  if (kitMorphStr) {
    let morphAbilArr = kitMorphStr.split("\\"); // The use of // produces a single /
    morphAbilArr.shift(); // Removes the first element of the array, which will be an empty element
    let r = kitEditorDataFirst_R;
    for (const element of morphAbilArr) {
      if (r > kitEditorDataLast_R)
        throw new Error(
          `The kit "${kitNameID}" has more abilities than will fit on the <KitEditor> table. You need to add rows to this table.`
        );
      const [abilID, otherMorphStr, sk1MorphStr, sk2MorphStr] =
        element.split("|");
      kitEditorArr[r][kitEditorAbilityNameID_C] = gGetVal(
        "db",
        "Abilities",
        abilID,
        "Name_ID"
      );
      kitEditorArr[r][kitEditorOtherMorph_C] = otherMorphStr;
      kitEditorArr[r][kitEditorMorph1_C] = sk1MorphStr;
      kitEditorArr[r][kitEditorMorph2_C] = sk2MorphStr;
      r++;
    }
  }

  gSaveArraySectionToSheet(
    kitEditorRef,
    kitEditorArr,
    kitEditorDataFirst_R,
    kitEditorDataLast_R,
    kitEditorAbilityNameID_C,
    kitEditorMorph2_C
  );

  // Sync <KitEditor> and <Abilities> bases and save Kit to capture any fSyncBase changes.
  fDBSyncBase();
  const blankAfterSave = false;
  fDBSaveKit(blankAfterSave);
} // End fDBLoadKit

// fDBSaveKit //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Saves Kit Morph to <Kits> table
function fDBSaveKit(blankAfterSave = true) {
  // force Load to capture all changes
  let kitEditor = getObjDBKitEditor(true);
  let kits = getObjDBKits(true);

  // Load <Kits> Info
  const kitMorphString_C = gHeaderC("db", "Kits", "MorphString");

  // Load <Kits> Info
  const kitNameID = kitEditor.arr[kitEditor.kitNameID_R][kitEditor.kitNameID_C];
  const kitID = gTestID("db", "Kits", kitNameID);
  if (!kitID)
    throw new Error(
      `In fDBSaveKit, the selected kit "${kitNameID}" does not exist in <Kits>`
    );

  // Sync <KitEditor> and <Abilities> bases
  fDBSyncBase();
  // Refresh kitEditor
  kitEditor = getObjDBKitEditor(true);

  // Ability MorphStr = \abilID|other morphs|sk1 morphs|sk2 morphs and each morph set can be comma delimited
  let kitMorphStr = "";
  let abilityMorphStr;
  for (let r = kitEditor.dataFirst_R; r <= kitEditor.dataLast_R; r++) {
    abilityMorphStr = "";
    const abilNameID = kitEditor.arr[r][kitEditor.nameID_C];
    const abilID = gTestID("db", "Abilities", abilNameID);
    if (abilID) {
      const otherMorphStr = kitEditor.arr[r][kitEditor.otherMorph_C];
      const morph1Str = kitEditor.arr[r][kitEditor.morph1_C];
      const morph2Str = kitEditor.arr[r][kitEditor.morph2_C];
      abilityMorphStr = fDBMakeAbilityMorph(
        otherMorphStr,
        morph1Str,
        morph2Str
      );
      kitMorphStr = abilityMorphStr
        ? `${kitMorphStr}\\${abilID}|${abilityMorphStr}`
        : kitMorphStr; // the two \\ will produce just one \
    }
  }

  // Do NOT save an empty kitMorphStr, but otherwise...
  if (kitMorphStr)
    gSaveVal(
      "db",
      "KitEditor",
      kitEditor.kitMorphString_R,
      kitEditor.kitMorphString_C,
      kitMorphStr
    );
  if (kitMorphStr) gSaveVal("db", "Kits", kitID, kitMorphString_C, kitMorphStr);

  // Clear table to prevent accidentally choosing a new ability and overridding it with another
  if (blankAfterSave) {
    gFillArraySection(
      kitEditor.arr,
      kitEditor.dataFirst_R,
      kitEditor.dataLast_R,
      kitEditor.nameID_C,
      kitEditor.morph2_C,
      ""
    );
    gSaveArraySectionToSheet(
      kitEditor.ref,
      kitEditor.arr,
      kitEditor.dataFirst_R,
      kitEditor.dataLast_R,
      kitEditor.nameID_C,
      kitEditor.morph2_C
    );
  }

  // force Load <KitEditor> and <kits> to capture all changes
  kitEditor = getObjDBKitEditor(true);
  kits = getObjDBKits(true);
} // End fDBSaveKit

// fDBMakeAbilityMorph //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Returns a cleaned and tested morph string for the ability
// Format: \abilID|other morphs|sk1 morphs|sk2 morphs and each morph set can be comma delimited
function fDBMakeAbilityMorph(otherMorphStr, morph1Str, morph2Str) {
  const finalOtherMorph = fDBCleanOtherMorph_Str(otherMorphStr);
  const finalMorph1 = fDBCleanMorph1or2_Str(morph1Str);
  const finalMorph2 = fDBCleanMorph1or2_Str(morph2Str);

  let abilMorphStr = `${finalOtherMorph}|${finalMorph1}|${finalMorph2}`;
  if (abilMorphStr === "||") abilMorphStr = "";

  return abilMorphStr;
} // End fDBMakeAbilityMorph

// fDBCleanOtherMorph_Str //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> returns a cleaned and tested string of Other Morph, that is comma delimited with no leading or trailing commas
function fDBCleanOtherMorph_Str(morphString) {
  morphString = String(morphString);
  if (morphString === "") return "";

  morphString = String(morphString)
    .replace(/\s+/g, "") // Remove all whitespace
    .replace(/[|\\]/g, "") // Remove any | or \ as these are used as delimiters for klMorph below
    .replace(/,+/g, ",") // Reduce multiple commas to one
    .replace(/^,|,$/g, ""); // Remove leading and trailing commas

  morphString = morphString.toUpperCase();

  // Split morphString into morphArr for testing
  const morphArr = morphString.split(",");

  let testStr;
  // For each element of morphArr verify that it matches one of the cases
  for (const element of morphArr) {
    switch (true) {
      case element === "":
      case element === "~":
      case element === "KNOWN":
      case element === "LEARN":

      // Starts and ends with ( and ), e.g. (Catongi)
      case /^\(.*\)$/.test(element):
        break;

      default:
        throw new Error(
          `In fDBCleanOtherMorph_Str MorphOther has an invalid argument: "${element}"\n`
        );
    }
  }

  // Return the cleaned and verified morphString
  return morphString;
} // End fDBCleanOtherMorph_Str

// fDBCleanMorph1or2_Str //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> returns a cleaned and tested string of Morph1 or Morph2, that is comma delimited with no leading or trailing commas
function fDBCleanMorph1or2_Str(morphString) {
  morphString = String(morphString);
  if (morphString === "") return "";

  morphString = String(morphString)
    .replace(/\s+/g, "") // Remove all whitespace
    .replace(/[|\\]/g, "") // Remove any | or \ as these are used as delimiters for klMorph below
    .replace(/,+/g, ",") // Reduce multiple commas to one
    .replace(/^,|,$/g, ""); // Remove leading and trailing commas

  morphString = morphString.toUpperCase();

  const morphArr = morphString.split(",");

  for (const element of morphArr) {
    switch (true) {
      case element === "":
        break;

      case /[~TPLAROYGBIVSUE]/.test(element):
        break;

      default:
        throw new Error(
          `In fDBCleanMorph1or2_Str morph has an invalid argument: "${element}"\n`
        );
    }
  }

  return morphString;
} // End fDBCleanMorph1or2_Str

// fDBSyncBase //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Synchronizes bases between <KitEditor> and <Abilities> based on certain conditions
function fDBSyncBase() {
  // force Load data
  let kits = getObjDBKits(true);
  let kitEd = getObjDBKitEditor(true);
  let abil = getObjDBAbilities(true);

  const kitID = gTestID(
    "db",
    "Kits",
    kitEd.arr[kitEd.kitNameID_R][kitEd.kitNameID_C]
  );
  const kitHeader_C = gTestHeaderC("db", "Abilities", kitID);

  let primaryBaseTab = [""];

  for (let r = kitEd.dataFirst_R; r <= kitEd.dataLast_R; r++) {
    const abilID = kitEd.arr[r][kitEd.key_C];

    // If abilID exists (not '')
    if (abilID) {
      const abil_R = gKeyR("db", "Abilities", abilID);
      const abilName = abil.arr[abil_R][abil.nameID_C].split("   ")[0].trim();

      const kitOtherMorph = kitEd.arr[r][kitEd.otherMorph_C];
      const kitCantSkill = /(^~|,~)/.test(kitOtherMorph);

      const abilBase1 = abil.arr[abil_R][abil.base1_C];
      const abilBase2 = abil.arr[abil_R][abil.base2_C];

      const abilBase1isNum = kitCantSkill
        ? false
        : abilBase1 !== "" && !isNaN(abilBase1);
      const abilBase2isNum = kitCantSkill
        ? false
        : abilBase2 !== "" && !isNaN(abilBase2);

      const kitMorph1 = kitEd.arr[r][kitEd.morph1_C];
      const kitMorph2 = kitEd.arr[r][kitEd.morph2_C];

      let abilSk1Morph = "";
      let abilSk2Morph = "";

      // If Kit is an RC Kit then...
      if (kitHeader_C) {
        abilSk1Morph = kitCantSkill ? "" : abil.arr[abil_R][kitHeader_C];
        abilSk2Morph = kitCantSkill ? "" : abil.arr[abil_R][kitHeader_C + 1];

        // Else use Default Sk1 and Sk2 PLAGHE from <Abilities>
      } else {
        abilSk1Morph = kitCantSkill
          ? ""
          : abil.arr[abil_R][abil.defaultPLAGHESk1_C];
        abilSk2Morph = kitCantSkill
          ? ""
          : abil.arr[abil_R][abil.defaultPLAGHESk2_C];
      }

      const finalMorph1 = fDBSetFinalMorphString(
        kitMorph1,
        abilSk1Morph,
        abilBase1isNum,
        primaryBaseTab,
        abilName
      );
      const finalMorph2 = fDBSetFinalMorphString(
        kitMorph2,
        abilSk2Morph,
        abilBase2isNum,
        primaryBaseTab,
        abilName
      );

      kitEd.arr[r][kitEd.morph1_C] = finalMorph1;
      kitEd.arr[r][kitEd.morph2_C] = finalMorph2;

      if (kitHeader_C) {
        abil.arr[abil_R][kitHeader_C] = finalMorph1;
        abil.arr[abil_R][kitHeader_C + 1] = finalMorph2;
      }
    }
  }

  gSaveArraySectionToSheet(
    kitEd.ref,
    kitEd.arr,
    kitEd.dataFirst_R,
    kitEd.dataLast_R,
    kitEd.morph1_C,
    kitEd.morph2_C
  );
  if (kitHeader_C)
    gSaveArraySectionToSheet(
      abil.ref,
      abil.arr,
      abil.dataFirst_R,
      abil.dataLast_R,
      kitHeader_C,
      kitHeader_C + 1
    );

  // Force load
  kits = getObjDBKits(true);
  kitEd = getObjDBKitEditor(true);
  abil = getObjDBAbilities(true);
} // End fDBSyncBase

// fDBSetFinalMorphString //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Determines if a morph should be used based on <Abilities> abilBaseisNum
// Purpose -> If both kitMorph and abilMorph are present, prompts the user to decide which to use
// Purpose -> If only one of kitMorph or abilMorph is present, that morph is used
// Purpose -> Returns finalMorphString
function fDBSetFinalMorphString(
  kitMorph,
  abilMorph,
  abilBaseIsNum,
  primaryTab,
  abilName
) {
  // If baseIsNum is false, set finalMorphString to an empty string
  if (!abilBaseIsNum) {
    finalMorphString = "";
  } else {
    // If both kitMorph and abilMorph are present
    if (kitMorph && abilMorph) {
      if (String(kitMorph) !== String(abilMorph)) {
        primaryTab[0] =
          primaryTab[0] ||
          fDBPromptKitEdOrAbilAsPrimary(abilName, kitMorph, abilMorph); // fPromptKitEdOrAbilAsPrimary returns 'KitEditor' or 'Abilities'
        finalMorphString = primaryTab[0] === "KitEditor" ? kitMorph : abilMorph;
      } else {
        finalMorphString = kitMorph;
      }

      // If only one of kitMorph or abilMorph is present, use that value
    } else if (kitMorph || abilMorph) {
      finalMorphString = kitMorph ? kitMorph : abilMorph;

      // Else set finalMorphString to A
    } else finalMorphString = "A";
  }

  return finalMorphString;
} // End fDBSetFinalMorphString

// fDBPromptKitEdOrAbilAsPrimary //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Prompts the user to choose between <KitEditor> or <Abilities> and returns the chosen value
function fDBPromptKitEdOrAbilAsPrimary(abilName, kitMorph, abilMorph) {
  const ui = SpreadsheetApp.getUi();

  // Prompt the user to choose between 'KitEditor' or 'Abilities'
  const response = ui.alert(
    "Choose Primary Source",
    `Morph conflict: ${abilName} <KitEditor>: ${kitMorph}, <Ability> ${abilMorph}\n\nPlease choose whether the primary source of ALL Morphs should be\n<KitEditor> or <Abilities>.\n\nClick "NO" for <KitEditor>\nClick "YES" for <Abilities>`,
    ui.ButtonSet.YES_NO_CANCEL
  );

  // Return the chosen value based on the user's response
  if (response == ui.Button.NO) {
    return "KitEditor";
  } else if (response == ui.Button.YES) {
    return "Abilities";
  } else throw new Error(`Operation Canceled.`);
} // End fDBPromptKitEdOrAbilAsPrimary

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  Abilities Menu  (end g. Kits Menu)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// fDBSaveVersionsToBooks //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Saves Ability Version blocks (v1:  ..., v2:  ..., etc.) to Rule Books CM, EM, RB, SG
function fDBSaveVersionsToBooks() {
  let versions = getObjDBVersions();

  // *** Sort <Versions> on VerID so all versions of an ability are contiguous (But NOT in version order due to 10 coming before 2, etc.) ***
  const numRowsToSort = versions.dataLast_R - versions.dataFirst_R + 1;
  const numColToSort = versions.last_C - versions.first_C + 1;
  var sortRange = versions.ref.getRange(
    versions.dataFirst_R + 1,
    versions.first_C + 1,
    numRowsToSort,
    numColToSort
  );
  sortRange.sort(1);

  // Force reload of versions
  versions = getObjDBVersions(true);

  const cmRef = DocumentApp.openById(g.id.cm);
  const emRef = DocumentApp.openById(g.id.em);

  fDBFindIDsWithVersionsAndReplace(cmRef, versions);
  fDBFindIDsWithVersionsAndReplace(emRef, versions);

  // Left over from the first run, this deletes versions blocks that don't start with "v1:" where each Ver was its own paragraph, rather than the entire block being a paragraph
  fDBDeleteVer2Up(cmRef);
  fDBDeleteVer2Up(emRef);

  cmRef.saveAndClose();
  emRef.saveAndClose();
} // End fDBSaveVersionsToBooks

// fDBFindIDsWithVersionsAndReplace //////////////////////////////////////////////////////////////////////////////////////////////////
// Changes all "v#: " paragraphs to @!@!@!
function fDBFindIDsWithVersionsAndReplace(docRef, versions) {
  let docBody = docRef.getBody();
  const paragraphArr = docBody.getParagraphs();

  for (let i = 0; i < paragraphArr.length; i++) {
    let text = paragraphArr[i].getText();
    if (text.includes("||")) {
      const abilIDArr = text.split("||");
      let newVerBlock = "";

      for (let j = 1; j < abilIDArr.length; j++) {
        const abilID = abilIDArr[j].trim(); // Assuming abilID is immediately after '||'
        newVerBlock = fDBNewVerBlock(versions, abilID);
        if (newVerBlock) {
          fDBReplaceOldVerBlock(paragraphArr, i, newVerBlock);
          break;
        }
      }
    }
  }
} // End fDBFindIDsWithVersionsAndReplace

// fDBNewVerBlock //////////////////////////////////////////////////////////////////////////////////////////////////
// Get newVersion Block of Text for provided abilID. If abilID not in db <Versions> returns false
function fDBNewVerBlock(versions, abilID) {
  const idVer1 = `${abilID}.v1`;
  const firstVer_R = gKeyR("db", "Versions", idVer1);

  if (firstVer_R) {
    let lastVer_R = firstVer_R;
    while (
      lastVer_R < versions.dataLast_R &&
      versions.arr[lastVer_R][versions.id_C] === abilID
    ) {
      lastVer_R++;
    }
    // Decrement to stay within the last valid version
    if (
      lastVer_R > firstVer_R &&
      versions.arr[lastVer_R][versions.id_C] !== abilID
    ) {
      lastVer_R--;
    }

    let newVersionBlock = "";
    for (let r = firstVer_R; r <= lastVer_R; r++) {
      newVersionBlock += `\n${versions.arr[r][versions.verText_C]}`;
    }

    // trim removes leading \n
    return newVersionBlock.trim();
  }

  return false;
} // End fDBNewVerBlock

// fDBReplaceOldVerBlock //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Finds old version block starting from start_i and replaces it with newVerBlock
function fDBReplaceOldVerBlock(paragraphArr, start_i, newVerBlock) {
  // Loop through paragraphs starting at start_i looking for the old version block
  let keepSearching = true;
  let i = start_i + 1; // Start looking after the identified starting paragraph, declared outside the for loop so it can be returned by this function
  for (; i < paragraphArr.length && keepSearching; i++) {
    let text = paragraphArr[i].getText();
    let match = text.match(/^v(\d+):/); // Use regex to find 'v#: ' at the start of the paragraph
    if (match) {
      // Replace the entire content of the version paragraph with the new version block
      paragraphArr[i].replaceText(".*", newVerBlock);
      keepSearching = false; // Stop searching after replacing
    } else if (
      paragraphArr[i].getHeading() !== DocumentApp.ParagraphHeading.NORMAL
    ) {
      // If the paragraph is a header, stop the search
      keepSearching = false;
    }
  }
} // End fDBReplaceOldVerBlock

// fDBDeleteVer2Up //////////////////////////////////////////////////////////////////////////////////////////////////
// Deletes all versions other than 1 from Rule Books CM, EM, RB, SG
function fDBDeleteVer2Up(docRef) {
  let docBody = docRef.getBody();
  const paragraphs = docBody.getParagraphs();

  // Loop through paragraphs in reverse to avoid indexing issues after deletion
  for (let i = paragraphs.length - 1; i >= 0; i--) {
    let text = paragraphs[i].getText();
    let match = text.match(/^v(\d+):/); // Use regex to find 'v#: ' at the start of the paragraph
    if (match && match[1] !== "1") {
      // Check if version number is not '1'
      docBody.removeChild(paragraphs[i]); // Remove the paragraph
    }
  }
} // End fDBDeleteVer2Up

// fDBCullOldIDsFromBooks //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Removes old ||oldIDs that are no longer in db <Elements> from Rule Books CM, EM, SG
function fDBCullOldIDsFromBooks() {
  const cmRef = DocumentApp.openById(g.id.cm);
  const emRef = DocumentApp.openById(g.id.em);
  const sgRef = DocumentApp.openById(g.id.sg);

  fDBCullOneBooksIDs(cmRef);
  fDBCullOneBooksIDs(emRef);
  fDBCullOneBooksIDs(sgRef);

  cmRef.saveAndClose();
  emRef.saveAndClose();
  sgRef.saveAndClose();
} // End fDBCullOldIDsFromBooks

// fDBCullOneBooksIDs //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Removes old ||oldIDs that are no longer in db <Elements> from a Rule Book
function fDBCullOneBooksIDs(bookRef) {
  let docBody = bookRef.getBody();
  const paragraphArr = docBody.getParagraphs();

  // Since Paragraphs may be removed, run the search from last paragraphArr element to the first
  for (var i = paragraphArr.length - 1; i >= 0; i--) {
    var oldText = paragraphArr[i].getText();
    if (oldText.startsWith("||")) {
      var barAbilArr = oldText.match(/\|\|(\w{6})/g) || [];
      var newText = oldText;

      barAbilArr.forEach(function (barAbil) {
        if (!fDBAbilIDIsFound(barAbil)) {
          newText = newText.replace(barAbil, "");
        }
      });

      if (newText.trim() && newText !== oldText) {
        paragraphArr[i].setText(newText.trim());

        // If newText is empty then delete the paragraph
      } else if (!newText.trim()) {
        // Use direct reference to remove paragraph to avoid stale reference issues
        docBody.removeChild(docBody.getParagraphs()[i]);
      }
    }
  }
} // End fDBCullOneBooksIDs

// fDBAbilIDIsFound //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Checks if an abilID is found in db <Elements> returns ID (without ||) if found and false if not found
function fDBAbilIDIsFound(barAbil) {
  const abilID = barAbil.slice(2);

  return gTestID("db", "Elements", abilID);
} // End fDBAbilIDIsFound

// fDBUpdateNotes //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Updates note column of <Elements> from CM,EM or SG
function fDBUpdateNotes() {
  let elem = getObjDBElements();

  // Blank all <Elements> notes
  gFillArraySection(
    elem.arr,
    elem.dataFirst_R,
    elem.dataLast_R,
    elem.notes_C,
    elem.notes_C,
    ""
  );

  const cmRef = DocumentApp.openById(g.id.cm);
  const emRef = DocumentApp.openById(g.id.em);
  const sgRef = DocumentApp.openById(g.id.sg);

  fDBFindAllNotes(cmRef, elem);
  fDBFindAllNotes(emRef, elem);
  fDBFindAllNotes(sgRef, elem);

  gSaveArraySectionToSheet(
    elem.ref,
    elem.arr,
    elem.dataFirst_R,
    elem.dataLast_R,
    elem.notes_C,
    elem.notes_C
  );
} // END fDBUpdateNotes

// fDBFindAllNotes //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Finds all ||abilID in book and adds its notes to <Elements>
function fDBFindAllNotes(bookRef, elem) {
  let docBody = bookRef.getBody();
  const paragraphArr = docBody.getParagraphs();

  for (var i = 0; i < paragraphArr.length; i++) {
    var text = paragraphArr[i].getText();
    if (text.startsWith("||")) {
      var barAbilArr = text.match(/\|\|(\w{6})/g) || []; // Gets EXACTLY '||' and next 6 characters
      let abilNotes = fDBFindFirstHeaderAbove(paragraphArr, i);

      // Accumulate notes down the book until a header is found
      for (
        var j = i + 1;
        j < paragraphArr.length &&
        paragraphArr[j].getHeading() === DocumentApp.ParagraphHeading.NORMAL;
        j++
      ) {
        abilNotes += paragraphArr[j].getText() + "\n"; // Collecting notes text, adding newline for readability
      }

      // Update notes for each abilID found in the paragraph
      barAbilArr.forEach(function (barAbil) {
        const abilID = barAbil.slice(2); // Assuming barAbil includes '||', slice it off to get the ID
        if (gTestID("db", "Elements", abilID)) {
          // Check if the ID is valid and found
          const elem_R = gKeyR("db", "Elements", abilID); // Get the row in the Elements database
          if (elem_R) {
            elem.arr[elem_R][elem.notes_C] = abilNotes;
          }
        }
      });
    }
  }
} // End fDBFindAllNotes

// fDBFindFirstHeaderAbove //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Finds the first header text above a given paragraph index in a Google Doc
// Returns -> The text of the first header found above the specified index or an empty string if no header is found
function fDBFindFirstHeaderAbove(paragraphArr, current_i) {
  let headerAbove = ""; // Initialize as empty string in case no headers are found above
  for (let k = current_i - 1; k >= 0; k--) {
    // Check if the paragraph is a header based on its style
    if (paragraphArr[k].getHeading() !== DocumentApp.ParagraphHeading.NORMAL) {
      headerAbove = paragraphArr[k].getText();
      break; // Stop searching after finding the first header
    }
  }
  return `${headerAbove}\n`;
} // End fDBFindFirstHeaderAbove

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  Designer Menu  (end g. Abilities Menu)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// fDBName_IDSync //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Scan's all sheets, parses text in cell A1's notes for "Fix Name_ID: C,D,etc.".
// It then scans all cells from HeaderR+1 to end and updates the contents to match <Elements>
function fDBName_IDSync() {
  const ssRef = gSSRef("db");
  const sheets = ssRef.getSheets();

  sheets.forEach((sheet) => {
    const note = sheet.getRange("A1").getNote();
    const fixNameIDPattern = /Fix Name_ID:\s*([A-Z,]+)/; // Removes Fix Name_ID and white space
    const match = note.match(fixNameIDPattern);

    if (match) {
      const sheetName = sheet.getName();
      let sheetArr = gArr("db", sheetName);
      const dataFirst_R = gDataFirst_R("db", sheetName);
      const dataLast_R = gDataLast_R("db", sheetName);

      let colList = match[1].replace(/\s/g, "");
      let colArr = colList.split(",");

      let arr_C;
      colArr.forEach((letter) => {
        arr_C = letterToColumn(letter) - 1; // Convert column letter to array index
        for (let r = dataFirst_R; r <= dataLast_R; r++) {
          if (sheetArr[r][arr_C] === "") continue;
          let elementID = gTestID("db", "elements", sheetArr[r][arr_C]);
          if (elementID === false)
            throw new Error(
              `In fDBName_IDSync sheet ${sheetName}'s ID not found: "${sheetArr[r][arr_C]}"`
            );
          sheetArr[r][arr_C] = gGetVal("db", "elements", elementID, "Name_ID");
        }
        // Save sheetArr to sheet from rows data_R to end of column arr_C
        gSaveArraySectionToSheet(
          sheet,
          sheetArr,
          dataFirst_R,
          dataLast_R,
          arr_C,
          arr_C
        );
      });
    }
  });

  // Nested Function: to convert column letter to column number
  function letterToColumn(letter) {
    let column = 0,
      length = letter.length;
    for (let i = 0; i < length; i++) {
      column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
    }
    return column;
  }
} // END fDBName_IDSync
