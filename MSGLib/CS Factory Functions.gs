////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  CS Factory Functions  ()
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// getObjCSGame //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> performs a gLoadTable if necessary (or forceLoad), then if necessary (or forceLoad) reloads the g.obj[ss][sheetname]
function getObjCSGame(forceLoad = false) {
  const ss = "mycs";
  const sheetName = "game";

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
    abilTblStart_R: gKeyR(ss, sheetName, "AbilTblStart_R"),
    abilTblEnd_R: gKeyR(ss, sheetName, "AbilTblEnd_R"),
    gearTblStart_R: gKeyR(ss, sheetName, "GearTblStart_R"),
    gearTblEnd_R: gKeyR(ss, sheetName, "GearTblEnd_R"),

    rtf_R: gKeyR(ss, sheetName, "RTF"),
    rollLog_R: gKeyR(ss, sheetName, "RollLog"),
    custMon_R: gKeyR(ss, sheetName, "CustMon"),
    nish_R: gKeyR(ss, sheetName, "Nish"),
    meta_R: gKeyR(ss, sheetName, "meta"),
    metaChnl_R: gKeyR(ss, sheetName, "chnl"),
    luckBox_R: gKeyR(ss, sheetName, "LuckBox"),
    luckBoxPlus_R: gKeyR(ss, sheetName, "LuckPlus"),
    actTotal_R: gKeyR(ss, sheetName, "Act"),
    actPlus_R: gKeyR(ss, sheetName, "ActPlus"),
    vitTbl_R: gKeyR(ss, sheetName, "VitTbl"),
    natAR_R: gKeyR(ss, sheetName, "NatAR"),
    monsterFirst_R: gKeyR(ss, sheetName, "FirstMon_R"),
    monsterLast_R: gKeyR(ss, sheetName, "LastMon_R"),
    chaos1_R: gKeyR(ss, sheetName, "Chaos1"),
    chaosWrist_R: gKeyR(ss, sheetName, "ChaosWrist"),

    agi_R: gKeyR(ss, sheetName, "Agi"),
    str_R: gKeyR(ss, sheetName, "Str"),
    for_R: gKeyR(ss, sheetName, "For"),
    arc_R: gKeyR(ss, sheetName, "Arc"),
    nishAtr_R: gKeyR(ss, sheetName, "NishAtr_R"),

    pcDmg_R: gKeyR(ss, sheetName, "PCDmg"),
    pcAR_R: gKeyR(ss, sheetName, "PCAR"),

    rtf_C: gHeaderC(ss, sheetName, "RTF"),
    rollLog_C: gHeaderC(ss, sheetName, "RollLog"),
    permMorph1_C: gHeaderC(ss, sheetName, "PermMorph1"),
    morph1_C: gHeaderC(ss, sheetName, "Morph1"),
    sk1Typ_C: gHeaderC(ss, sheetName, "Sk1Typ"),
    sk1_C: gHeaderC(ss, sheetName, "Sk1"),
    sk1ChkBox_C: gHeaderC(ss, sheetName, "Sk1ChkBox"),
    abilNameID_C: gHeaderC(ss, sheetName, "Ability"),
    condition_C: gHeaderC(ss, sheetName, "Condition"),
    sk2ChkBox_C: gHeaderC(ss, sheetName, "Sk2ChkBox"),
    sk2_C: gHeaderC(ss, sheetName, "Sk2"),
    sk2Typ_C: gHeaderC(ss, sheetName, "Sk2Typ"),
    morph2_C: gHeaderC(ss, sheetName, "Morph2"),
    permMorph2_C: gHeaderC(ss, sheetName, "PermMorph2"),
    ver_C: gHeaderC(ss, sheetName, "Ver"),
    notes_C: gHeaderC(ss, sheetName, "Notes"),
    pic_C: gHeaderC(ss, sheetName, "Pic"),
    act_C: gHeaderC(ss, sheetName, "Act"),
    dur_C: gHeaderC(ss, sheetName, "Dur"),
    rng_C: gHeaderC(ss, sheetName, "Rng"),
    meta_C: gHeaderC(ss, sheetName, "Meta"),
    uses_C: gHeaderC(ss, sheetName, "Uses"),
    regain_C: gHeaderC(ss, sheetName, "Regain"),
    on_C: gHeaderC(ss, sheetName, "On"),

    nish_C: gHeaderC(ss, sheetName, "Nish"),
    luckBox_C: gHeaderC(ss, sheetName, "LuckBox"),
    metaR_C: gHeaderC(ss, sheetName, "R"),
    metaB_C: gHeaderC(ss, sheetName, "B"),
    actTotal_C: gHeaderC(ss, sheetName, "ActTot"),
    numOfMonsters_C: gHeaderC(ss, sheetName, "NumMon"),
    monsterName_C: gHeaderC(ss, sheetName, "MonName"),
    monsterVit_C: gHeaderC(ss, sheetName, "MonVit"),
    monsterDef_C: gHeaderC(ss, sheetName, "MonDef"),
    monsterAR_C: gHeaderC(ss, sheetName, "MonAR"),
    monsterAtk_C: gHeaderC(ss, sheetName, "MonAtk"),
    monsterDmg_C: gHeaderC(ss, sheetName, "MonDmg"),
    monsterPic_C: gHeaderC(ss, sheetName, "MonPic"),
    monsterSize_C: gHeaderC(ss, sheetName, "MonSize"),
    chaosUses_C: gHeaderC(ss, sheetName, "ChaosUses"),
    chaosAbility_C: gHeaderC(ss, sheetName, "ChaosAbility"),
    vitPlus_C: gHeaderC(ss, sheetName, "vitPlus"),
    vitMax_C: gHeaderC(ss, sheetName, "vitMax"),
    vit1st_C: gHeaderC(ss, sheetName, "vit1st"),
    vitNow_C: gHeaderC(ss, sheetName, "vitNow"),
    vitNowPercent_C: gHeaderC(ss, sheetName, "VitNowPercent"),

    abilTableFirst_C: gHeaderC(ss, sheetName, "PermMorph1"),
    abilTableLast_C: gHeaderC(ss, sheetName, "Regain"),
    last_C: gHeaderC(ss, sheetName, "LastC"),

    // Socketed Gear Table
    socketTblStart_R: gKeyR(ss, sheetName, "SocketTblStart"),
    socketTblEnd_R: gKeyR(ss, sheetName, "SocketTblEnd"),

    socketedGear_C: gHeaderC(ss, sheetName, "SocketedGear"),
    socketColors_C: gHeaderC(ss, sheetName, "SocketColors"),
    socketedGems_C: gHeaderC(ss, sheetName, "SocketedGems"),

    // MR Table
    gearMRTbl_R: gKeyR(ss, sheetName, "GearMRTbl"),
    gearMRCheckBox_R: gKeyR(ss, sheetName, "GearMRCheckBox"),
    gearCarryTbl_R: gKeyR(ss, sheetName, "GearCarryTbl"),

    gearMRCol1_C: gHeaderC(ss, sheetName, "GearMRCol1"),
    gearMRCol6_C: gHeaderC(ss, sheetName, "GearMRCol6"),

    // Gear Table and related
    gearSpe_R: gKeyR(ss, sheetName, "GearSpe"),
    gearPlusMR_R: gKeyR(ss, sheetName, "GearPlusMR"),
    gearStr_R: gKeyR(ss, sheetName, "GearStr"),
    gearPlusCarry_R: gKeyR(ss, sheetName, "GearPlusCarry"),
    possEncTot_R: gKeyR(ss, sheetName, "PossEncTot"),
    possAPTot_R: gKeyR(ss, sheetName, "PossAPTot"),

    gearSpeAndStr_C: gHeaderC(ss, sheetName, "GearSpeAndStr"),
    gearPlusMR_C: gHeaderC(ss, sheetName, "GearPlusMR"),
    gearPlusCarry_C: gHeaderC(ss, sheetName, "GearPlusCarry"),

    possEncTot_C: gHeaderC(ss, sheetName, "PossEncTot"),
    possNum_C: gHeaderC(ss, sheetName, "PossNum"),
    possName_C: gHeaderC(ss, sheetName, "PossName"),
    possWorn_C: gHeaderC(ss, sheetName, "PossWorn"),
    possEnc_C: gHeaderC(ss, sheetName, "PossEnc"),
    possCrEa_C: gHeaderC(ss, sheetName, "PossCrEa"),
    possPerOff_C: gHeaderC(ss, sheetName, "PossPerOff"),
    possCrTot_C: gHeaderC(ss, sheetName, "PossCrTot"),
    possIsArtifact_C: gHeaderC(ss, sheetName, "PossIsArtifact"),
    possAPEa_C: gHeaderC(ss, sheetName, "PossAPEa"),
    possAPTot_C: gHeaderC(ss, sheetName, "PossAPTot"),
    possGrandAPTot_C: gHeaderC(ss, sheetName, "PossGrandAPTot"),
  });

  // Build any necessary child keys of g
  g.obj = g.obj || {};
  g.obj[ss] = g.obj[ss] || {};

  g.obj[ss][sheetName] = newObj;

  return newObj;
} // End getObjCSGame

// getObjCSRaceClass //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> performs a gLoadTable if necessary (or forceLoad), then if necessary (or forceLoad) reloads the g.obj[ss][sheetname]
function getObjCSRaceClass(forceLoad = false) {
  const ss = "mycs";
  const sheetName = "raceclass";

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

    playerName_R: gKeyR(ss, sheetName, "playerName"),
    slot_R: gKeyR(ss, sheetName, "slot"),
    charName_R: gKeyR(ss, sheetName, "charName"),
    charLevel_R: gKeyR(ss, sheetName, "level"),
    rc_R: gKeyR(ss, sheetName, "RC"),
    race_R: gKeyR(ss, sheetName, "race"),

    value_C: gHeaderC(ss, sheetName, "Val"),
  });

  // Build any necessary child keys of g
  g.obj = g.obj || {};
  g.obj[ss] = g.obj[ss] || {};

  g.obj[ss][sheetName] = newObj;

  return newObj;
} // End getObjCSRaceClass

// getObjCSList //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> performs a gLoadTable if necessary (or forceLoad), then if necessary (or forceLoad) reloads the g.obj[ss][sheetname]
function getObjCSList(forceLoad = false) {
  const ss = "mycs";
  const sheetName = "list";

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
    abilityNameID_C: gHeaderC(ss, sheetName, "AbilitiesNameID"),
    first_C: gHeaderC(ss, sheetName, "On"),
    last_C: gHeaderC(ss, sheetName, "AbilitiesNameID"),
  });

  // Build any necessary child keys of g
  g.obj = g.obj || {};
  g.obj[ss] = g.obj[ss] || {};

  g.obj[ss][sheetName] = newObj;

  return newObj;
} // End getObjCSList

// getObjCSPermMorph //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> performs a gLoadTable if necessary (or forceLoad), then if necessary (or forceLoad) reloads the g.obj[ss][sheetname]
function getObjCSPermMorph(forceLoad = false) {
  const ss = "mycs";
  const sheetName = "permmorph";

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
    abilNameID_C: gHeaderC(ss, sheetName, "AbilNameID"),
    condition_C: gHeaderC(ss, sheetName, "Condition"),
    morph1_C: gHeaderC(ss, sheetName, "Morph1"),
    oldSk1_C: gHeaderC(ss, sheetName, "OldSk1"),
    newSk1_C: gHeaderC(ss, sheetName, "PermSk1"),
    morphRoll1_C: gHeaderC(ss, sheetName, "PermMorphfRoll1"),
    morph2_C: gHeaderC(ss, sheetName, "Morph2"),
    oldSk2_C: gHeaderC(ss, sheetName, "OldSk2"),
    newSk2_C: gHeaderC(ss, sheetName, "PermSk2"),
    morphRoll2_C: gHeaderC(ss, sheetName, "PermMorphfRoll2"),
    why_C: gHeaderC(ss, sheetName, "Why"),
    first_C: gHeaderC(ss, sheetName, "ID"),
    last_C: gHeaderC(ss, sheetName, "Why"),
  });

  // Build any necessary child keys of g
  g.obj = g.obj || {};
  g.obj[ss] = g.obj[ss] || {};

  g.obj[ss][sheetName] = newObj;

  return newObj;
} // End getObjCSPermMorph

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                   (end CS Factory Functions)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
