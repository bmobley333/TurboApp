////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  PS Factory Functions  ()
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// getObjPSPartyLog //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> performs a gLoadTable if necessary (or forceLoad), then if necessary (or forceLoad) reloads the g.obj[ss][sheetname]
function getObjPSPartyLog(forceLoad = false) {
  const ss = "ps";
  const sheetName = "partylog";

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
} // End getObjPSPartyLog

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                   (end PS Factory Functions)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
