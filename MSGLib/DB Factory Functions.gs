////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  DB Factory Functions  ()
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// getObjDBGMScreen //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> performs a gLoadTable if necessary (or forceLoad), then if necessary (or forceLoad) reloads the g.obj[ss][sheetname]
function getObjDBGMScreen(forceLoad = false) {
  const ss = "db";
  const sheetName = "gmscreen";

  // Load Table (note: if g.[ss][sheetName] already exists it will not be reloaded unless forceLoad = true, to save run time)
  gLoadTable(ss, sheetName, forceLoad);

  // Don't reload g.obj[ss][sheetName] unless forceLoad is true or g.obj[ss][sheetName] doesn't exist.
  if (!forceLoad && g?.obj?.[ss]?.[sheetName]) return g.obj[ss][sheetName];

  // Create or update the sheet object - specifically use the existing g.obj[ss][sheetName] if it exists, else g.obj[ss][sheetName] doesn't yet exist and will be cretaed below
  // IMPORTANT - this guarantees that g.obj[ss][sheetName] will always be the same object and that calling this Factory Function will never create more than one object, (if new -> obj is made, if exists -> obj is updated)
  const newObj = g?.obj?.[ss]?.[sheetName] || {};

  // Assign new values to the object properties
  Object.assign(newObj, {
    ref: gSheetRef(ss, sheetName),
    arr: gArr(ss, sheetName),

    dataFirst_R: gDataFirst_R(ss, sheetName),
    dataLast_R: gDataLast_R(ss, sheetName),

    url_R: gKeyR(ss, sheetName, "URL"),
    log_R: gKeyR(ss, sheetName, "Log"),
    monFirst_R: gKeyR(ss, sheetName, "MonTblFirst_R"),
    monLast_R: gKeyR(ss, sheetName, "MonTblLast_R"),

    monActiveTF_C: gHeaderC(ss, sheetName, "ActiveMonTF"),
    monName_C: gHeaderC(ss, sheetName, "MonName"),
    monCustName_C: gHeaderC(ss, sheetName, "MonCustName"),
    monVit_C: gHeaderC(ss, sheetName, "Vit"),
    monNotes_C: gHeaderC(ss, sheetName, "Notes"),

    slot1_C: gHeaderC(ss, sheetName, "Slot1"),
    slot2_C: gHeaderC(ss, sheetName, "Slot2"),
    slot3_C: gHeaderC(ss, sheetName, "Slot3"),
    slot4_C: gHeaderC(ss, sheetName, "Slot4"),
    slot5_C: gHeaderC(ss, sheetName, "Slot5"),
    slot6_C: gHeaderC(ss, sheetName, "Slot6"),
    slot7_C: gHeaderC(ss, sheetName, "Slot7"),
    slot8_C: gHeaderC(ss, sheetName, "Slot8"),
    slot9_C: gHeaderC(ss, sheetName, "Slot9"),
  });

  // Build any necessary child keys of g
  g.obj = g.obj || {};
  g.obj[ss] = g.obj[ss] || {};

  g.obj[ss][sheetName] = newObj;

  return newObj;
} // End getObjDBGMScreen

// getObjDBAbilities //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> performs a gLoadTable if necessary (or forceLoad), then if necessary (or forceLoad) reloads the g.obj[ss][sheetname]
function getObjDBAbilities(forceLoad = false) {
  const ss = "db";
  const sheetName = "abilities";

  // Load Table (note: if g.[ss][sheetName] already exists it will not be reloaded unless forceLoad = true, to save run time)
  gLoadTable(ss, sheetName, forceLoad);

  // Don't reload g.obj[ss][sheetName] unless forceLoad is true or g.obj[ss][sheetName] doesn't exist.
  if (!forceLoad && g?.obj?.[ss]?.[sheetName]) return g.obj[ss][sheetName];

  // Create or update the sheet object - specifically use the existing g.obj[ss][sheetName] if it exists, else g.obj[ss][sheetName] doesn't yet exist and will be cretaed below
  // IMPORTANT - this guarantees that g.obj[ss][sheetName] will always be the same object and that calling this Factory Function will never create more than one object, (if new -> obj is made, if exists -> obj is updated)
  const newObj = g?.obj?.[ss]?.[sheetName] || {};

  // Assign new values to the object properties
  Object.assign(newObj, {
    ref: gSheetRef(ss, sheetName),
    arr: gArr(ss, sheetName),

    id_C: gHeaderC(ss, sheetName, "ID"),
    nameID_C: gHeaderC(ss, sheetName, "Name_ID"),
    base1_C: gHeaderC(ss, sheetName, "Base1"),
    base2_C: gHeaderC(ss, sheetName, "Base2"),
    defaultPLAGHESk1_C: gHeaderC(ss, sheetName, "DefaultPLAGHESk1"),
    defaultPLAGHESk2_C: gHeaderC(ss, sheetName, "DefaultPLAGHESk2"),
    parentKit_C: gHeaderC(ss, sheetName, "ParentKit"),
    parentMinVer_C: gHeaderC(ss, sheetName, "ParentMinVer"),
    kitFirstSk1_C: gHeaderC(ss, sheetName, "KitFirstSk2") - 1,
    kitLastSk1_C: gHeaderC(ss, sheetName, "KitLastSk2") - 1,

    dataFirst_R: gDataFirst_R(ss, sheetName),
    dataLast_R: gDataLast_R(ss, sheetName),

    sk1Typ_C: gHeaderC(ss, sheetName, "SkTyp1"),
    sk2Typ_C: gHeaderC(ss, sheetName, "SkTyp2"),
    isItem_C: gHeaderC(ss, sheetName, "IsItem"),
    isNonRCPossession_C: gHeaderC(ss, sheetName, "isNonRCPossession"),
    notes_C: gHeaderC(ss, sheetName, "Notes"),
    pic_C: gHeaderC(ss, sheetName, "Pic"),
  });

  // Build any necessary child keys of g
  g.obj = g.obj || {};
  g.obj[ss] = g.obj[ss] || {};

  g.obj[ss][sheetName] = newObj;

  return newObj;
} // End getObjDBAbilities

// getObjDBElements //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> performs a gLoadTable if necessary (or forceLoad), then if necessary (or forceLoad) reloads the g.obj[ss][sheetname]
function getObjDBElements(forceLoad = false) {
  const ss = "db";
  const sheetName = "elements";

  // Load Table (note: if g.[ss][sheetName] already exists it will not be reloaded unless forceLoad = true, to save run time)
  gLoadTable(ss, sheetName, forceLoad);

  // Don't reload g.obj[ss][sheetName] unless forceLoad is true or g.obj[ss][sheetName] doesn't exist.
  if (!forceLoad && g?.obj?.[ss]?.[sheetName]) return g.obj[ss][sheetName];

  // Create or update the sheet object - specifically use the existing g.obj[ss][sheetName] if it exists, else g.obj[ss][sheetName] doesn't yet exist and will be cretaed below
  // IMPORTANT - this guarantees that g.obj[ss][sheetName] will always be the same object and that calling this Factory Function will never create more than one object, (if new -> obj is made, if exists -> obj is updated)
  const newObj = g?.obj?.[ss]?.[sheetName] || {};

  // Assign new values to the object properties
  Object.assign(newObj, {
    ref: gSheetRef(ss, sheetName),
    arr: gArr(ss, sheetName),
    dataFirst_R: gDataFirst_R(ss, sheetName),
    dataLast_R: gDataLast_R(ss, sheetName),

    id_C: gHeaderC(ss, sheetName, "ID"),
    isAbil_C: gHeaderC(ss, sheetName, "isAbil"),
    isGear_C: gHeaderC(ss, sheetName, "isGear"),
    book_C: gHeaderC(ss, sheetName, "Book"),
    notes_C: gHeaderC(ss, sheetName, "Notes"),
  });

  // Build any necessary child keys of g
  g.obj = g.obj || {};
  g.obj[ss] = g.obj[ss] || {};

  g.obj[ss][sheetName] = newObj;

  return newObj;
} // End getObjDBElements

// getObjDBKits //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> performs a gLoadTable if necessary (or forceLoad), then if necessary (or forceLoad) reloads the g.obj[ss][sheetname]
function getObjDBKits(forceLoad = false) {
  const ss = "db";
  const sheetName = "kits";

  // Load Table (note: if g.[ss][sheetName] already exists it will not be reloaded unless forceLoad = true, to save run time)
  gLoadTable(ss, sheetName, forceLoad);

  // Don't reload g.obj[ss][sheetName] unless forceLoad is true or g.obj[ss][sheetName] doesn't exist.
  if (!forceLoad && g?.obj?.[ss]?.[sheetName]) return g.obj[ss][sheetName];

  // Create or update the sheet object - specifically use the existing g.obj[ss][sheetName] if it exists, else g.obj[ss][sheetName] doesn't yet exist and will be cretaed below
  // IMPORTANT - this guarantees that g.obj[ss][sheetName] will always be the same object and that calling this Factory Function will never create more than one object, (if new -> obj is made, if exists -> obj is updated)
  const newObj = g?.obj?.[ss]?.[sheetName] || {};

  // Assign new values to the object properties
  Object.assign(newObj, {
    ref: gSheetRef(ss, sheetName),
    arr: gArr(ss, sheetName),
    dataFirst_R: gDataFirst_R(ss, sheetName),
    dataLast_R: gDataLast_R(ss, sheetName),

    id_C: gHeaderC(ss, sheetName, "ID"),
    nameOnly_C: gHeaderC(ss, sheetName, "Name"),
    nameID_C: gHeaderC(ss, sheetName, "Name_ID"),
  });

  // Build any necessary child keys of g
  g.obj = g.obj || {};
  g.obj[ss] = g.obj[ss] || {};

  g.obj[ss][sheetName] = newObj;

  return newObj;
} // End getObjDBKits

// getObjDBKitEditor //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> performs a gLoadTable if necessary (or forceLoad), then if necessary (or forceLoad) reloads the g.obj[ss][sheetname]
function getObjDBKitEditor(forceLoad = false) {
  const ss = "db";
  const sheetName = "kiteditor";

  // Load Table (note: if g.[ss][sheetName] already exists it will not be reloaded unless forceLoad = true, to save run time)
  gLoadTable(ss, sheetName, forceLoad);

  // Don't reload g.obj[ss][sheetName] unless forceLoad is true or g.obj[ss][sheetName] doesn't exist.
  if (!forceLoad && g?.obj?.[ss]?.[sheetName]) return g.obj[ss][sheetName];

  // Create or update the sheet object - specifically use the existing g.obj[ss][sheetName] if it exists, else g.obj[ss][sheetName] doesn't yet exist and will be cretaed below
  // IMPORTANT - this guarantees that g.obj[ss][sheetName] will always be the same object and that calling this Factory Function will never create more than one object, (if new -> obj is made, if exists -> obj is updated)
  const newObj = g?.obj?.[ss]?.[sheetName] || {};

  // Assign new values to the object properties
  Object.assign(newObj, {
    ref: gSheetRef(ss, sheetName),
    arr: gArr(ss, sheetName),
    dataFirst_R: gDataFirst_R(ss, sheetName),
    dataLast_R: gDataLast_R(ss, sheetName),

    kitNameID_R: gKeyR(ss, sheetName, "KitNameID"),
    kitMorphString_R: gKeyR(ss, sheetName, "MorphStr"),

    key_C: gHeaderC(ss, sheetName, "Key"),
    nameID_C: gHeaderC(ss, sheetName, "AbilNameID"),
    kitNameID_C: gHeaderC(ss, sheetName, "KitNameID"),
    kitMorphString_C: gHeaderC(ss, sheetName, "MorphString"),
    otherMorph_C: gHeaderC(ss, sheetName, "OtherMorph"),
    morph1_C: gHeaderC(ss, sheetName, "Morph1"),
    morph2_C: gHeaderC(ss, sheetName, "Morph2"),

    first_C: gHeaderC(ss, sheetName, "Key"),
    last_C: gHeaderC(ss, sheetName, "Morph2"),
  });

  // Build any necessary child keys of g
  g.obj = g.obj || {};
  g.obj[ss] = g.obj[ss] || {};

  g.obj[ss][sheetName] = newObj;

  return newObj;
} // End getObjDBKitEditor

// getObjDBVersions //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> performs a gLoadTable if necessary (or forceLoad), then if necessary (or forceLoad) reloads the g.obj[ss][sheetname]
function getObjDBVersions(forceLoad = false) {
  const ss = "db";
  const sheetName = "versions";

  // Load Table (note: if g.[ss][sheetName] already exists it will not be reloaded unless forceLoad = true, to save run time)
  gLoadTable(ss, sheetName, forceLoad);

  // Don't reload g.obj[ss][sheetName] unless forceLoad is true or g.obj[ss][sheetName] doesn't exist.
  if (!forceLoad && g?.obj?.[ss]?.[sheetName]) return g.obj[ss][sheetName];

  // Create or update the sheet object - specifically use the existing g.obj[ss][sheetName] if it exists, else g.obj[ss][sheetName] doesn't yet exist and will be cretaed below
  // IMPORTANT - this guarantees that g.obj[ss][sheetName] will always be the same object and that calling this Factory Function will never create more than one object, (if new -> obj is made, if exists -> obj is updated)
  const newObj = g?.obj?.[ss]?.[sheetName] || {};

  // Assign new values to the object properties
  Object.assign(newObj, {
    ref: gSheetRef(ss, sheetName),
    arr: gArr(ss, sheetName),
    dataFirst_R: gDataFirst_R(ss, sheetName),
    dataLast_R: gDataLast_R(ss, sheetName),
    verID_C: gHeaderC(ss, sheetName, "VerID"),
    id_C: gHeaderC(ss, sheetName, "ID"),
    ver_C: gHeaderC(ss, sheetName, "Ver"),
    act_C: gHeaderC(ss, sheetName, "Act"),
    dur_C: gHeaderC(ss, sheetName, "Dur"),
    rng_C: gHeaderC(ss, sheetName, "Rng"),
    meta_C: gHeaderC(ss, sheetName, "Meta"),
    uses_C: gHeaderC(ss, sheetName, "Uses"),
    regain_C: gHeaderC(ss, sheetName, "Regain"),
    verText_C: gHeaderC(ss, sheetName, "VerText"),
    first_C: gHeaderC(ss, sheetName, "VerID"),
    last_C: gHeaderC(ss, sheetName, "NotesText"),
  });

  // Build any necessary child keys of g
  g.obj = g.obj || {};
  g.obj[ss] = g.obj[ss] || {};

  g.obj[ss][sheetName] = newObj;

  return newObj;
} // End getObjDBVersions

// getObjDBGear //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> performs a gLoadTable if necessary (or forceLoad), then if necessary (or forceLoad) reloads the g.obj[ss][sheetname]
function getObjDBGear(forceLoad = false) {
  const ss = "db";
  const sheetName = "gear";

  // Load Table (note: if g.[ss][sheetName] already exists it will not be reloaded unless forceLoad = true, to save run time)
  gLoadTable(ss, sheetName, forceLoad);

  // Don't reload g.obj[ss][sheetName] unless forceLoad is true or g.obj[ss][sheetName] doesn't exist.
  if (!forceLoad && g?.obj?.[ss]?.[sheetName]) return g.obj[ss][sheetName];

  // Create or update the sheet object - specifically use the existing g.obj[ss][sheetName] if it exists, else g.obj[ss][sheetName] doesn't yet exist and will be cretaed below
  // IMPORTANT - this guarantees that g.obj[ss][sheetName] will always be the same object and that calling this Factory Function will never create more than one object, (if new -> obj is made, if exists -> obj is updated)
  const newObj = g?.obj?.[ss]?.[sheetName] || {};

  // Assign new values to the object properties
  Object.assign(newObj, {
    ref: gSheetRef(ss, sheetName),
    arr: gArr(ss, sheetName),

    dataFirst_R: gDataFirst_R(ss, sheetName),
    dataLast_R: gDataLast_R(ss, sheetName),

    id_C: gHeaderC(ss, sheetName, "ID"),
    nameID_C: gHeaderC(ss, sheetName, "Name_ID"),
    isArtifact_C: gHeaderC(ss, sheetName, "IsArtifact"),
    apCost_C: gHeaderC(ss, sheetName, "AP"),
    itemCR_C: gHeaderC(ss, sheetName, "ItemCR"),
    itemEnc_C: gHeaderC(ss, sheetName, "ItemEnc"),
    isWorn_C: gHeaderC(ss, sheetName, "IsWorn"),
    wornEnc_C: gHeaderC(ss, sheetName, "WornEnc"),
    notes_C: gHeaderC(ss, sheetName, "Notes"),
    pic_C: gHeaderC(ss, sheetName, "Pic"),
  });

  // Build any necessary child keys of g
  g.obj = g.obj || {};
  g.obj[ss] = g.obj[ss] || {};

  g.obj[ss][sheetName] = newObj;

  return newObj;
} // End getObjDBGear

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                   (end DB Factory Functions)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
