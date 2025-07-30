
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  KL Factory Functions ()
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


// getObjKLAllAbilities //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> performs a gLoadTable if necessary (or forceLoad), then if necessary (or forceLoad) reloads the g.obj[ss][sheetname]
function getObjKLAllAbilities(forceLoad = false) {
  const ss = 'mykl';
  const sheetName = 'allabilities';

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

    id_C: gHeaderC(ss, sheetName, 'ID'),
    nameID_C: gHeaderC(ss, sheetName, 'Name_ID'),
    source_C: gHeaderC(ss, sheetName, 'Source'),
    cantLearn_C: gHeaderC(ss, sheetName, 'CantLearn'),
    known_C: gHeaderC(ss, sheetName, 'Known'),
    learn_C: gHeaderC(ss, sheetName, 'Learn'),
    base1_C: gHeaderC(ss, sheetName, 'Base1'),
    base2_C: gHeaderC(ss, sheetName, 'Base2'),
    morphOther_C: gHeaderC(ss, sheetName, 'MorphOther'),
    sk1PLAGHE_C: gHeaderC(ss, sheetName, 'sk1PLAGHE'),
    sk2PLAGHE_C: gHeaderC(ss, sheetName, 'sk2PLAGHE'),
    finalSk1_C: gHeaderC(ss, sheetName, 'FinalSk1'),
    finalSk2_C: gHeaderC(ss, sheetName, 'FinalSk2'),
    first_C: gHeaderC(ss, sheetName, 'ID'),
    last_C: gHeaderC(ss, sheetName, 'FinalSk2'),
  });

  // Build any necessary child keys of g
  g.obj = g.obj || {};
  g.obj[ss] = g.obj[ss] || {};

  g.obj[ss][sheetName] = newObj;

  return newObj;
} // End getObjKLAllAbilities



// getObjKL_KLTab //////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * Performs a gLoadTable if necessary (or if forceLoad is true), then creates or updates the 
 * g.obj[ss][sheetName] object with pre-calculated references to its data table.
 * This acts as a factory/singleton, ensuring only one object is created per sheetName and is updated in place.
 * * @param {string} tabName The name of the sheet/tab to process.
 * @param {boolean} [forceLoad=false] If true, forces a reload of the data table and the object.
 * @returns {object} The object containing references for the specified sheet.
 */
function getObjKL_RCTab(tabName, forceLoad = false) {
  const ss = 'mykl';

  // Validate that tabName is a non-empty string.
  if (typeof tabName !== 'string' || tabName.length === 0) {
    throw new Error(`getObjKL_KLTab was passed an invalid tab name of "${tabName}"`);
  }
  
  const sheetName = tabName;
  // Load Table (note: if g[ss][sheetName] already exists it will not be reloaded unless forceLoad = true)
  gLoadTable(ss, sheetName, forceLoad);

  // If the object already exists and a reload is not forced, return the existing object to save processing time.
  if (!forceLoad && g?.obj?.[ss]?.[sheetName]) {
    return g.obj[ss][sheetName];
  }

  // Create or get a reference to the sheet object.
  // This ensures that g.obj[ss][sheetName] is always the same object.
  // If it's new, the object is created; if it exists, it's updated.
  const newObj = g?.obj?.[ss]?.[sheetName] || {};

  // Assign common properties to the object.
  Object.assign(newObj, {
    ref: gSheetRef(ss, sheetName),
    arr: gArr(ss, sheetName),
    dataFirst_R: gDataFirst_R(ss, sheetName),
    dataLast_R: gDataLast_R(ss, sheetName),

    myLvl_C: gHeaderC(ss, sheetName, 'MyLvl'),
    tier_C: gHeaderC(ss, sheetName, 'Tier'),
    apVal_C: gHeaderC(ss, sheetName, 'APVal'),

    myLvl_R: gKeyR(ss, sheetName, 'MyLvl'),
    tier_R: gKeyR(ss, sheetName, 'Tier'),
    bnsAP_R: gKeyR(ss, sheetName, 'BnsAP'),
    totalAP_R: gKeyR(ss, sheetName, 'TotalAP'),
    spentAP_R: gKeyR(ss, sheetName, 'SpentAP'),
    remainingAP_R: gKeyR(ss, sheetName, 'RemainingAP'),
  });

  // Conditionally add properties only if the sheetName is 'All'.
  if (sheetName === 'All') {
    newObj.rc_R = gKeyR(ss, sheetName, 'RC');
    newObj.rc_C = gHeaderC(ss, sheetName, 'RC');
  }

  // Ensure the global object structure exists before assignment.
  g.obj = g.obj || {};
  g.obj[ss] = g.obj[ss] || {};

  // Assign the new or updated object to the global namespace.
  g.obj[ss][sheetName] = newObj;

  return newObj;
} // End getObjKL_KLTab





// getObjKLMyAbilities //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> performs a gLoadTable if necessary (or forceLoad), then if necessary (or forceLoad) reloads the g.obj[ss][sheetname]
function getObjKLMyAbilities(forceLoad = false) {
  const ss = 'mykl';
  const sheetName = 'myabilities';

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

    reminderCB_C: gHeaderC(ss, sheetName, 'ReminderCB'),  
    id_C: gHeaderC(ss, sheetName, 'ID'),  
    known_C: gHeaderC(ss, sheetName, 'Known'),  
    learn_C: gHeaderC(ss, sheetName, 'Learn'),  
    charLevel_C: gHeaderC(ss, sheetName, 'Level'),  
    minLearnLvl_C: gHeaderC(ss, sheetName, 'MinLearnLvl'),  
    train_C: gHeaderC(ss, sheetName, 'Training'),  
    kitTrain_C: gHeaderC(ss, sheetName, 'KitTrain'),  
    abilAPCost_C: gHeaderC(ss, sheetName, 'AbilAPCost'),  
    minVerLvl_C: gHeaderC(ss, sheetName, 'MinVerLvl'),  
    ver_C: gHeaderC(ss, sheetName, 'Ver'),  
    maxVer_C: gHeaderC(ss, sheetName, 'vMax'),  
    nameID_C: gHeaderC(ss, sheetName, 'NameID'),  
    notes_C: gHeaderC(ss, sheetName, 'Notes'),  
    source_C: gHeaderC(ss, sheetName, 'Source'),  
    parentKit_C: gHeaderC(ss, sheetName, 'ParentKit'),  
    parentMinVer_C: gHeaderC(ss, sheetName, 'ParentMinVer'),  
    base1_C: gHeaderC(ss, sheetName, 'Base1'),  
    base2_C: gHeaderC(ss, sheetName, 'Base2'),  
    trainedSk1_C: gHeaderC(ss, sheetName, 'TrainedSk1'),  
    trainedSk2_C: gHeaderC(ss, sheetName, 'TrainedSk2'),  
    act_C: gHeaderC(ss, sheetName, 'Act'),  
    dur_C: gHeaderC(ss, sheetName, 'Dur'),  
    rng_C: gHeaderC(ss, sheetName, 'Rng'),  
    meta_C: gHeaderC(ss, sheetName, 'Meta'),  
    uses_C: gHeaderC(ss, sheetName, 'Uses'),  
    regain_C: gHeaderC(ss, sheetName, 'Regain'),  
    first_C: gHeaderC(ss, sheetName, 'ID'),  
    last_C: gHeaderC(ss, sheetName, 'Regain'),

    // For Header AP info
    reminderCB_R: gKeyR(ss, sheetName, 'ReminderCB'),
    gearAP_R: gKeyR(ss, sheetName, 'GearAPTot'),
    apData1_R: gKeyR(ss, sheetName, 'apData1'),
    apData2_R: gKeyR(ss, sheetName, 'apData2'),
    apData3_R: gKeyR(ss, sheetName, 'apData3'),
    charLevel_R: gKeyR(ss, sheetName, 'Level'),
    apCount_C: gHeaderC(ss, sheetName, 'APCount'),
    apSpent_C: gHeaderC(ss, sheetName, 'APSpent'),
    apTotal_C: gHeaderC(ss, sheetName, 'APTotal'),
  });

  // Build any necessary child keys of g
  g.obj = g.obj || {};
  g.obj[ss] = g.obj[ss] || {};

  g.obj[ss][sheetName] = newObj;

  return newObj;
} // End getObjKLMyAbilities



// getObjKLmyKLs //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> performs a gLoadTable if necessary (or forceLoad), then if necessary (or forceLoad) reloads the g.obj[ss][sheetname]
function getObjKLMyKLs(forceLoad = false) {
  const ss = 'mykl';
  const sheetName = 'MyKLs';

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
    nameID_C: gHeaderC(ss, sheetName, 'NameID'),
    notes_C: gHeaderC(ss, sheetName, 'Notes'),
  });

  // Build any necessary child keys of g
  g.obj = g.obj || {};
  g.obj[ss] = g.obj[ss] || {};

  g.obj[ss][sheetName] = newObj;

  return newObj;
} // End getObjKLMyKL



// getObjKnownAbilities //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> performs a gLoadTable if necessary (or forceLoad), then if necessary (or forceLoad) reloads the g.obj[ss][sheetname]
function getObjKnownAbilities(forceLoad = false) {
  const ss = 'mykl';
  const sheetName = 'KnownAbilities';

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

    id_C: gHeaderC(ss, sheetName, 'ID'), 
    nameID_C: gHeaderC(ss, sheetName, 'Name_ID'), 
    finalSk1_C: gHeaderC(ss, sheetName, 'FinalSk1'),
    finalSk2_C: gHeaderC(ss, sheetName, 'FinalSk2'),  
    ver_C: gHeaderC(ss, sheetName, 'Ver'),
    notes_C: gHeaderC(ss, sheetName, 'Notes'),  
    act_C: gHeaderC(ss, sheetName,'Act'),
    dur_C: gHeaderC(ss, sheetName,'Dur'),
    rng_C: gHeaderC(ss, sheetName,'Rng'),
    meta_C: gHeaderC(ss, sheetName,'Meta'),
    uses_C: gHeaderC(ss, sheetName,'Uses'),
    regain_C: gHeaderC(ss, sheetName,'Regain'),
    buff_C: gHeaderC(ss, sheetName, 'Buff'),  
    kitID_C: gHeaderC(ss, sheetName, 'KitID'),  
    kitBuff_C: gHeaderC(ss, sheetName, 'KitBuff'),  
    base1_C: gHeaderC(ss, sheetName, 'Base1'),  
    base2_C: gHeaderC(ss, sheetName, 'Base2'),
    sk1PLAGHE_C: gHeaderC(ss, sheetName, 'sk1PLAGHE'),  
    sk2PLAGHE_C: gHeaderC(ss, sheetName, 'sk2PLAGHE'),  
  });

  // Build any necessary child keys of g
  g.obj = g.obj || {};
  g.obj[ss] = g.obj[ss] || {};

  g.obj[ss][sheetName] = newObj;

  return newObj;
} // End getObjKnownAbilities









// getObjKLRogueAbilities //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> performs a gLoadTable if necessary (or forceLoad), then if necessary (or forceLoad) reloads the g.obj[ss][sheetname]
function getObjKLRogueAbilities(forceLoad = false) {
  const ss = 'mykl';
  const sheetName = 'rogueabilities';

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
    id_C: gHeaderC(ss, sheetName, 'ID'),
    known_C: gHeaderC(ss, sheetName, 'Known'),
    nameID_C: gHeaderC(ss, sheetName, 'NameID'),
    notes_C: gHeaderC(ss, sheetName, 'Notes'),
    first_C: gHeaderC(ss, sheetName, 'ID'),
    last_C: gHeaderC(ss, sheetName, 'Why'),
  });

  // Build any necessary child keys of g
  g.obj = g.obj || {};
  g.obj[ss] = g.obj[ss] || {};

  g.obj[ss][sheetName] = newObj;

  return newObj;
} // End getObjKLRogueAbilities






// getObjKLList //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> performs a gLoadTable if necessary (or forceLoad), then if necessary (or forceLoad) reloads the g.obj[ss][sheetname]
function getObjKLList(forceLoad = false) {
  const ss = 'mykl';
  const sheetName = 'list';

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

    kits_C: gHeaderC(ss, sheetName, 'Kits'),
  });

  // Build any necessary child keys of g
  g.obj = g.obj || {};
  g.obj[ss] = g.obj[ss] || {};

  g.obj[ss][sheetName] = newObj;

  return newObj;
} // End getObjKLList










////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                   (end KL Factory Functions)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////











