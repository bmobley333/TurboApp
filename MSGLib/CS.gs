// CS

// 2025.03.29

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  Menu (end initialize)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// fCSCreateMenu //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Create MetaScape menu
function fCSCreateMenu() {
  SpreadsheetApp.getUi()
    .createMenu("*** Game")
    .addItem("Turbo", "fCSMenuTurboUI") // Note this is ONLY to create the "Turbo" menu choice. All turbo code including the rest of the normal menu functions are ran natively from the Character Sheet and the Turbo Web App
    .addItem("Roll", "fCSMenuRoll")
    .addItem("Roll - Lucked", "fCSMenuRollLucked")
    .addItem("Roll - Free", "fCSMenuRollFree")
    .addSeparator()
    .addItem("Show Sidebar", "fCSMenuShowSidebar")
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu("Jump to")
        .addItem("Top (Nish)", "fCSMenuScrollToNish")
        .addItem("Bottom (Gear)", "fCSMenuScrollToGear")
    )
    .addSeparator()
    .addItem("Meta - Nish Start or End", "fCSMenuNishStartOrEnd")
    .addItem("Meta Flood", "fCSMenuMetaFlood")
    .addSeparator()
    .addItem("GM Award", "fCSMenuGMAward")
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu("Healing")
        .addItem("Natural Healing", "fCSMenuNaturalHealing")
        .addItem("All Wounds", "fCSMenuHealAll")
    )
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu("Rest or Time")
        .addItem("Rest", "fCSMenuRest")
        .addItem("Sleep", "fCSMenuSleep")
        .addItem("New Game Session", "fCSMenuNewGameSession")
    )
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu("Clear")
        .addItem("Channeled Meta", "fCSMenuClearChanneledMeta")
        .addItem("Check Boxes of All Sk1 & Sk2", "fCSMenuClearSk1Sk2CheckBoxes")
        .addItem("Monsters", "fCSMenuClearMonsters")
        .addItem("Morphs", "fCSMenuClearMorphs")
        .addItem("Nish", "fCSMenuClearNish")
        .addItem("Roll Log", "fCSMenuClearRollLog")
    )
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu("Refresh")
        .addItem("Abilities and Gear", "fCSMenuRefreshAbilities")
        .addItem("Monsters", "fCSMenuLoadMonsters")
        .addItem("Vit", "fCSMenuCalcVitMax")
    )
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu("Set to Max")
        .addItem("Luck", "fCSMenuLucktoFull")
        .addItem("Meta", "fCSMenuMetaToFull")
        .addItem("Reset All to Max", "fCSMenuRestAllToMax")
    )
    .addToUi();
  SpreadsheetApp.getUi()
    .createMenu("Gear")
    .addItem("Hide Gear", "fCSMenuHideGear")
    .addItem("Show Gear", "fCSMenuShowGear")
    .addSeparator()
    .addItem("Loot", "fCSMenuRollRndTreasure")
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu("Specific Treasure")
        .addItem("Artifact", "fCSMenuRollArtifact")
        .addItem("Chaos Crystal", "fCSMenuChaosCrystal")
        .addItem("Gear", "fCSmenuRollGear")
        .addItem("Gem (for socket)", "fCSMenuRollSocketGem")
        .addItem("Socketed Item", "fCSMenuRollSockets")
        .addItem("Valuables", "fCSMenuRollValuables")
    )
    .addToUi();
  SpreadsheetApp.getUi()
    .createMenu("PermMorph")
    .addItem("Refresh PermMorph Table", "fCSMenuRefreshAbilities")
    .addToUi();
  SpreadsheetApp.getUi()
    .createMenu("RaceClass")
    .addItem("Refresh All - Saves Data", "fCSMenuRefreshRaceClass")
    .addSeparator()
    .addItem("Build Race Dropdown", "fCSMenuBuildRaceDropDown")
    .addItem("Update Race Notes & Pic", "fCSMenuUpdateRaceNotesPic")
    .addItem(
      `Generate Race Gender to Senses`,
      "fCSMenuGenerateRaceGenderToVision"
    )
    .addSeparator()
    .addItem("Create KeyLine", "fCSMenuCreateKL")
    .addItem("Refresh - URLs", "fCSMenuRefreshRuleBookURLs")
    .addToUi();
  SpreadsheetApp.getUi()
    .createMenu("System ***")
    .addItem("1 - Authorize Script", "fCSMenuAuthorize")
    .addItem("2 - Save My URL", "fCSMenuSaveCSIDURL")
    .addSeparator()
    .addItem("Refresh Menus", "fCSMenuRefreshAll")
    .addToUi();

  g.id.mycs = SpreadsheetApp.getActiveSpreadsheet().getId();
  g.id.mykl = gGetVal("mycs", "Data", "myKLID", "Val");
  if (gGetVal("mycs", "Data", "designer", "Val") === true) {
    SpreadsheetApp.getUi()
      .createMenu("DESIGNER")
      .addItem("Hide All", "fCSMenuHideAll")
      .addItem("Un-Hide All", "fCSMenuUn_HideAll")
      .addToUi();
  }
} // end fCSCreateMenu

// Menu Functions //////////////////////////////////////////////////////////////////////////////////////////////////
// <Game> Menu
// function fCSMenuTurboUI()  Is Ran in local sheet
function fCSMenuRoll() {
  fCSRunMenuOrButton("Roll");
}
function fCSMenuRollLucked() {
  fCSRunMenuOrButton("RollLucked");
}
function fCSMenuRollFree() {
  fCSRunMenuOrButton("RollFree");
}
function fCSMenuShowSidebar() {
  fCSRunMenuOrButton("ShowSidebar");
}
function fCSMenuScrollToNish() {
  fCSRunMenuOrButton("ScrollToNish");
}
function fCSMenuScrollToGear() {
  fCSRunMenuOrButton("ScrollToGear");
}
function fCSMenuNishStartOrEnd() {
  fCSRunMenuOrButton("NishStartOrEnd");
}
function fCSMenuLucktoFull() {
  fCSRunMenuOrButton("LuckToFull");
}
function fCSMenuMetaToFull() {
  fCSRunMenuOrButton("MetaToFull");
}
function fCSMenuMetaFlood() {
  fCSRunMenuOrButton("MetaFlood");
}
function fCSMenuRefreshAbilities() {
  fCSRunMenuOrButton("RefreshAbilities");
}
function fCSMenuLoadMonsters() {
  fCSRunMenuOrButton("LoadMonsters");
}
function fCSMenuGMAward() {
  fCSRunMenuOrButton("GMAward");
}
function fCSMenuRollRndTreasure() {
  fCSRunMenuOrButton("RollRndTreasure");
}
function fCSMenuChaosCrystal() {
  fCSRunMenuOrButton("ChaosCrystal");
}
function fCSMenuRollArtifact() {
  fCSRunMenuOrButton("RollArtifact");
}
function fCSmenuRollGear() {
  fCSRunMenuOrButton("RollGear");
}
function fCSMenuRollSockets() {
  fCSRunMenuOrButton("RollSockets");
}
function fCSMenuRollSocketGem() {
  fCSRunMenuOrButton("RollSocketGem");
}
function fCSMenuRollValuables() {
  fCSRunMenuOrButton("RollValuables");
}
function fCSMenuNaturalHealing() {
  fCSRunMenuOrButton("NaturalHealing");
}
function fCSMenuHealAll() {
  fCSRunMenuOrButton("HealAll");
}
function fCSMenuClearRollLog() {
  fCSRunMenuOrButton("ClearRollLog");
}
function fCSMenuRest() {
  fCSRunMenuOrButton("Rest");
}
function fCSMenuSleep() {
  fCSRunMenuOrButton("Sleep");
}
function fCSMenuNewGameSession() {
  fCSRunMenuOrButton("NewGameSession");
}
function fCSMenuRestAllToMax() {
  fCSRunMenuOrButton("RestAllToMax");
}
function fCSMenuCalcVitMax() {
  fCSRunMenuOrButton("CalcVitMax");
}
function fCSMenuClearChanneledMeta() {
  fCSRunMenuOrButton("ClearChanneledMeta");
}
function fCSMenuClearSk1Sk2CheckBoxes() {
  fCSRunMenuOrButton("ClearSk1Sk2CheckBoxes");
}
function fCSMenuClearMonsters() {
  fCSRunMenuOrButton("ClearMonsters");
}
function fCSMenuClearMorphs() {
  fCSRunMenuOrButton("ClearMorphs");
}
function fCSMenuClearNish() {
  fCSRunMenuOrButton("ClearNish");
}
// Gear
function fCSMenuShowGear() {
  fCSRunMenuOrButton("ShowGear");
}
function fCSMenuHideGear() {
  fCSRunMenuOrButton("HideGear");
}
// RaceClass Menu
function fCSMenuRefreshRaceClass() {
  fCSRunMenuOrButton("RefreshRaceClass");
}
function fCSMenuBuildRaceDropDown() {
  fCSRunMenuOrButton("BuildRaceDropDown");
}
function fCSMenuUpdateRaceNotesPic() {
  fCSRunMenuOrButton("UpdateRaceNotesPic");
}
function fCSMenuGenerateRaceGenderToVision() {
  fCSRunMenuOrButton("GenerateRaceGenderToVision");
}
function fCSMenuRefreshRuleBookURLs() {
  fCSRunMenuOrButton("RefreshRuleBookURLs");
}
// PermMorph Menu
// None as it is simply fMenuRefreshAbilities from the // <Game> Menu
// System Menu
function fCSMenuAuthorize() {
  fCSRunMenuOrButton("Authorize");
}
function fCSMenuSaveCSIDURL() {
  fCSRunMenuOrButton("SaveCSIDURL");
}
function fCSMenuCreateKL() {
  fCSRunMenuOrButton("CreateKL");
}
function fCSMenuRefreshAll() {
  fCSRunMenuOrButton("RefreshMenu");
}
// Designer Menu
function fCSMenuHideAll() {
  fCSRunMenuOrButton("HideAll");
}
function fCSMenuUn_HideAll() {
  fCSRunMenuOrButton("Un_HideAll");
}
// End Menu Functions

// Image Button Functions //////////////////////////////////////////////////////////////////////////////////////////////////
function fCSButtond20() {
  fCSRunMenuOrButton("Roll");
}
function fCSButtonRefreshAbilities() {
  fCSRunMenuOrButton("RefreshAbilities");
}
function fCSButtonScrollToNish() {
  fCSRunMenuOrButton("ScrollToNish");
}
function fCSButtonScrollToGear() {
  fCSRunMenuOrButton("ScrollToGear");
}
function fCSButtonMonsters() {
  fCSRunMenuOrButton("LoadMonsters");
}
function fCSButtonRefreshRaceClass() {
  fCSRunMenuOrButton("RefreshRaceClass");
}
// RefreshPermMorph is the is the same as and uses fCSButtonRefreshAbilities
// End Button Functions

// fCSRunMenuOrButton //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> To run all menu & button choices inside a try-catch-error
function fCSRunMenuOrButton(menuChoice) {
  try {
    g.id.mycs = SpreadsheetApp.getActiveSpreadsheet().getId();
    g.id.mykl = gGetVal("mycs", "Data", "myKLID", "Val");

    switch (menuChoice) {
      // Game Menu
      // case 'TurboUI': is ran in local sheet
      case "Roll":
        fCSMasterRoller();
        break;
      case "RollLucked":
        fCSMasterRoller("Lucked");
        break;
      case "RollFree":
        fCSMasterRoller("Free");
        break;
      case "ShowSidebar":
        break; // Note done here, done on CS local code
      case "ScrollToNish":
        fCSScrollToNish();
        break;
      case "ScrollToGear":
        fCSScrollToGear();
        break;
      case "NishStartOrEnd":
        fCSNishStartOrEnd();
        break;
      case "LuckToFull":
        fCSLuckToMax();
        break;
      case "MetaToFull":
        fCSMetaToFull();
        break;
      case "MetaFlood":
        fCSMetaFlood();
        break;
      case "RefreshAbilities":
        fCSRefreshAbilities();
        break;
      case "LoadMonsters":
        fCSLoadMonsters();
        break;
      case "GMAward":
        fCSGMAward();
        break;
      case "RollRndTreasure":
        fCSRollRndTreasure();
        break;
      case "ChaosCrystal":
        fCSChaosCrystal();
        break;
      case "RollArtifact":
        fCSRollArtifact();
        break;
      case "RollGear":
        fCSRollGear();
        break;
      case "RollSockets":
        fCSRollSockets();
        break;
      case "RollSocketGem":
        fCSRollSocketGem();
        break;
      case "RollValuables":
        fCSRollValuables();
        break;
      case "NaturalHealing":
        fCSNaturalHealing();
        break;
      case "HealAll":
        fCSHealAll();
        break;
      case "ClearRollLog":
        fCSClearRollLog();
        break;
      case "Rest":
        fCSRest();
        break;
      case "Sleep":
        fCSSleep();
        break;
      case "NewGameSession":
        fCSNewGameSession();
        break;
      case "RestAllToMax":
        fCSRestAllToMax();
        break;
      case "CalcVitMax":
        fCSCalcVitMax();
        break;
      case "ClearChanneledMeta":
        fCSMetaChnlClear();
        break;
      case "ClearSk1Sk2CheckBoxes":
        fCSClearSk1Sk2CheckBoxes();
        break;
      case "ClearMonsters":
        fCSClearMonsters();
        break;
      case "ClearMorphs":
        fCSClearMorphs();
        break;
      case "ClearNish":
        fCSClearNish();
        break;
      // Gear
      case "ShowGear":
        fCSShowGear(true);
        break;
      case "HideGear":
        fCSShowGear(false);
        break;
      // RaceClass Menu
      case "RefreshRaceClass":
        fCSRefreshRaceClass();
        break;
      case "BuildRaceDropDown":
        fCSBuildRaceDropDown();
        break;
      case "UpdateRaceNotesPic":
        fCSUpdateRaceNotesPic();
        break;
      case "GenerateRaceGenderToVision":
        fCSGenerateRaceGenderToVision();
        break;
      case "RefreshRuleBookURLs":
        fCSRefreshRuleBookURLs();
        break;
      // PermMorph menu
      // None as it is simply RefreshAbilities from <Game> Menu
      // System Menu
      case "Authorize":
        SpreadsheetApp.getUi().alert(
          `AUTHORIZED`,
          `Script Authorized!`,
          SpreadsheetApp.getUi().ButtonSet.OK
        );
        fCSSaveCSID();
        break;
      case "SaveCSIDURL":
        fCSSaveCSID();
        break;
      case "CreateKL":
        fCSCreateKL();
        break;
      case "RefreshMenu":
        fCSCreateMenu();
        break;
      // Designer Menu
      case "HideAll":
        gHideAll("mycs");
        break;
      case "Un_HideAll":
        gUn_HideAll("mycs");
        break;
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert(error); // NOTE: an error of End or end will simply end the program.
  }
  return "This is it";
} // end fCSRunMenuOrButton

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  Game  (end g. Menu)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// fCSScrollToNish //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Scrolls <Game> sheet to Nish (1st ability row)
function fCSScrollToNish() {
  let game = getObjCSGame();

  gScrollToCell(game.nishAtr_R, game.abilNameID_C);
} // End fCSScrollToNish

// fCSScrollToGear //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Scrolls <Game> sheet to Nish (1st ability row)
function fCSScrollToGear() {
  let game = getObjCSGame();

  gScrollToCell(game.gearTblStart_R + 20, game.abilNameID_C);
} // End fCSScrollToGear

// fCSMetaToFull //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Fills Meta
function fCSMetaToFull(updateRollLog = true) {
  // Set <Game> Constants
  const gameRef = gSheetRef("mycs", "game");
  const gameArr = gArr("mycs", "game");
  const meta_R = gKeyR("mycs", "game", "meta");
  const metaRed_C = gHeaderC("mycs", "game", "R");
  const metaBlue_C = gHeaderC("mycs", "game", "B");

  // Set the section in gameArr to [5, 4, 3, 2, 1]
  const maxMeta = [5, 4, 3, 2, 1];
  for (let i = 0; i < maxMeta.length; i++) {
    gameArr[meta_R][metaRed_C + i] = maxMeta[i];
  }

  // Save the updated gameArr back to the sheet
  gSaveArraySectionToSheet(
    gameRef,
    gameArr,
    meta_R,
    meta_R,
    metaRed_C,
    metaBlue_C
  );

  if (updateRollLog) fCSPostMyRoll(`${g.bluBS}Meta to Full${g.endS}<br>`);
} // End fCSMetaToFull

// fCSMetaFlood //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Meta Flood Roll
function fCSMetaFlood(returnResult = false) {
  // Set <Game> Constants
  const gameRef = gSheetRef("mycs", "game");
  const gameArr = gArr("mycs", "game");
  const meta_R = gKeyR("mycs", "game", "meta");
  const metaRed_C = gHeaderC("mycs", "game", "R");
  const metaBlue_C = gHeaderC("mycs", "game", "B");

  // Set the section in gameArr to [5, 4, 3, 2, 1]
  const maxMeta = [5, 4, 3, 2, 1];
  const metaColor = ["R", "O", "Y", "G", "B"];

  // Roll the dice to determine metaRoll
  const metaRoll = fCSd(5) - 1;

  // Update gameArr based on metaRoll
  for (let i = metaRoll; i >= 0; i--) {
    gameArr[meta_R][metaRed_C + i] += metaRoll - i + 1;
  }

  // Ensure that gameArr values do not exceed the corresponding maxMeta values
  for (let i = 0; i < maxMeta.length; i++) {
    gameArr[meta_R][metaRed_C + i] = Math.min(
      gameArr[meta_R][metaRed_C + i],
      maxMeta[i]
    );
  }

  // Save the updated gameArr back to the sheet
  gSaveArraySectionToSheet(
    gameRef,
    gameArr,
    meta_R,
    meta_R,
    metaRed_C,
    metaBlue_C
  );
  const result = `${g.bluBS}Meta Flood:${g.endS} ${metaColor[metaRoll]}<br>`;
  if (returnResult) return result;
  // Else...
  fCSPostMyRoll(result);
} // End fCSMetaFlood

// fCSMetaChnlClear //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Clears channeled Meta
function fCSMetaChnlClear() {
  // Set <Game> Constants
  const game = getObjCSGame();

  gFillArraySection(
    game.arr,
    game.metaChnl_R,
    game.metaChnl_R,
    game.metaR_C,
    game.metaB_C,
    ""
  );

  // Save the updated chnl section of gameArr back to the sheet
  gSaveArraySectionToSheet(
    game.ref,
    game.arr,
    game.metaChnl_R,
    game.metaChnl_R,
    game.metaR_C,
    game.metaB_C
  );
} // End fCSMetaChnlClear

// fCSNishStartOrEnd //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Performs Nish Start or End Meta, Channel, and On Maintenance
function fCSNishStartOrEnd() {
  const result = fCSMetaFlood(true);
  fCSMetaChnlClear();
  fCSResetAbilOn();
  fCSClearRollLog();
  fCSPostMyRoll(
    `${g.bluBS}Nish Starts or Ends:${g.endS}<br>${result}<br>----------------------------------------`
  );
} // End fCSNishStartOrEnd

// fCSResetAbilOn //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Resets the ON column to all "."
function fCSResetAbilOn() {
  let game = getObjCSGame();

  gFillArraySection(
    game.arr,
    game.abilTblStart_R,
    game.abilTblEnd_R,
    game.on_C,
    game.on_C,
    "."
  );
  gFillArraySection(
    game.arr,
    game.gearTblStart_R,
    game.gearTblEnd_R,
    game.on_C,
    game.on_C,
    "."
  );
  gSaveArraySectionToSheet(
    game.ref,
    game.arr,
    game.abilTblStart_R,
    game.gearTblEnd_R,
    game.on_C,
    game.on_C
  );
} // End fCSResetAbilOn

// fCSRefreshAbilities //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Refreshes <Game> abilities and gear tables
function fCSRefreshAbilities() {
  let objGame = getObjCSGame();
  let klMyAbil = getObjKLMyAbilities();
  let dbAbil = getObjDBAbilities();
  let dbGear = getObjDBGear();

  const total = {
    sumAP: 0,
    sumEnc: 0,
  };

  // Search every Ability on <Game> within in the specified range
  for (let r = objGame.dataFirst_R; r <= objGame.gearTblEnd_R; r++) {
    // Skip rows between abil and gear tables
    if (r > objGame.abilTblEnd_R && r < objGame.gearTblStart_R) {
      r = objGame.gearTblStart_R - 1;
      continue;
    }

    const abil = {
      row: objGame.arr[r],
      isElem: false,
      isMyAbil: false,
      isGear: false,
      klMyAbil_R: -1,
      dbAbil_R: -1,
      dbGear_R: -1,
    };

    // set elemID to false if not in KL's <MyAbilities> or DB's <Gear> otherwise store the element's ID - note can be in both such as a known artifact
    let elemID = abil.row[objGame.abilNameID_C];
    abil.isElem = gTestID("db", "Elements", elemID);
    let testMyID = gTestID("mykl", "MyAbilities", elemID);
    if (testMyID) {
      abil.klMyAbil_R = gKeyR("mykl", "MyAbilities", testMyID);
      const learned = klMyAbil.arr[abil.klMyAbil_R][klMyAbil.learn_C];
      if (learned) {
        abil.isMyAbil = true;
        abil.dbAbil_R = gKeyR("db", "Abilities", testMyID);
      }
    }

    testMyID = gTestID("db", "Gear", elemID);
    if (testMyID) {
      abil.isGear = true;
      abil.dbGear_R = gKeyR("db", "Gear", testMyID);
    }

    if (abil.isMyAbil || abil.isGear) {
      fFillGameAbilAndGearRow(
        r,
        objGame,
        klMyAbil,
        dbAbil,
        dbGear,
        total,
        abil
      );
    } else {
      fClearGameAbilAndGearRow(objGame, r, abil.isElem);
    }
  }

  // Save Gear Totals (Enc, AP)
  objGame.arr[objGame.possEncTot_R][objGame.possEncTot_C] = total.sumEnc;
  objGame.arr[objGame.possAPTot_R][objGame.possGrandAPTot_C] = total.sumAP;

  // Calculates Max Vit (self saves to <Game>), MR, MaxSockets
  fCSCalcVitMax();
  fCSCalcMR(objGame, total.sumEnc); // This uses total.sumEnc, assuming this is the correct argument.
  fCSCalcMaxSocketedItems(objGame);

  // Save the updated Abilities back to <Game>
  gSaveArraySectionToSheet(
    objGame.ref,
    objGame.arr,
    objGame.dataFirst_R,
    objGame.dataLast_R,
    objGame.abilTableFirst_C,
    objGame.last_C
  );

  // Apply Conditions
  fCSRefreshPermMorph();
} // end fCSRefreshAbilities

// fFillGameAbilAndGearRow //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Fills the <Game> ability and gear tables appropriately
// Purpose -> Assumes abil is either isMyAbil or isGear (or both) BUT not neither
function fFillGameAbilAndGearRow(
  r,
  objGame,
  klMyAbil,
  dbAbil,
  dbGear,
  total,
  abil
) {
  if (!abil.isMyAbil) abil.row[objGame.morph1_C] = ",";
  abil.row[objGame.sk1Typ_C] = abil.isMyAbil
    ? dbAbil.arr[abil.dbAbil_R][dbAbil.sk1Typ_C]
    : "";
  abil.row[objGame.sk1_C] = abil.isMyAbil
    ? klMyAbil.arr[abil.klMyAbil_R][klMyAbil.trainedSk1_C]
    : "";
  abil.row[objGame.abilNameID_C] =
    r === objGame.nishAtr_R
      ? "Nish                                                                                  _k97cmz"
      : abil.isMyAbil
      ? dbAbil.arr[abil.dbAbil_R][dbAbil.nameID_C]
      : dbGear.arr[abil.dbGear_R][dbGear.nameID_C];
  if (!abil.isMyAbil) abil.row[objGame.condition_C] = "";
  abil.row[objGame.sk2_C] = abil.isMyAbil
    ? klMyAbil.arr[abil.klMyAbil_R][klMyAbil.trainedSk2_C]
    : "";
  abil.row[objGame.sk2Typ_C] = abil.isMyAbil
    ? dbAbil.arr[abil.dbAbil_R][dbAbil.sk2Typ_C]
    : "";
  if (!abil.isMyAbil) abil.row[objGame.morph2_C] = ",";
  abil.row[objGame.ver_C] = abil.isMyAbil
    ? `v${klMyAbil.arr[abil.klMyAbil_R][klMyAbil.ver_C]}`
    : "";
  abil.row[objGame.notes_C] = abil.isMyAbil
    ? dbAbil.arr[abil.dbAbil_R][dbAbil.notes_C]
    : dbGear.arr[abil.dbGear_R][dbGear.notes_C];
  abil.row[objGame.pic_C] = abil.isMyAbil
    ? dbAbil.arr[abil.dbAbil_R][dbAbil.pic_C]
    : dbGear.arr[abil.dbGear_R][dbGear.pic_C];
  abil.row[objGame.act_C] = abil.isMyAbil
    ? klMyAbil.arr[abil.klMyAbil_R][klMyAbil.act_C]
    : "";
  abil.row[objGame.dur_C] = abil.isMyAbil
    ? klMyAbil.arr[abil.klMyAbil_R][klMyAbil.dur_C]
    : "";
  abil.row[objGame.rng_C] = abil.isMyAbil
    ? klMyAbil.arr[abil.klMyAbil_R][klMyAbil.rng_C]
    : "";
  abil.row[objGame.meta_C] = abil.isMyAbil
    ? klMyAbil.arr[abil.klMyAbil_R][klMyAbil.meta_C]
    : "";

  // Ensure existingUses and maxUses are numbers
  const existingUses = abil.row[objGame.uses_C];
  const maxUses = abil.isMyAbil
    ? klMyAbil.arr[abil.klMyAbil_R][klMyAbil.uses_C]
    : "";
  abil.row[objGame.uses_C] = existingUses === "" ? maxUses : existingUses;
  abil.row[objGame.regain_C] = abil.isMyAbil
    ? klMyAbil.arr[abil.klMyAbil_R][klMyAbil.regain_C]
    : "";

  // Fill Gear Enc Table
  let numOfGear = !abil.isGear
    ? ""
    : abil.row[objGame.possNum_C] === ""
    ? 1
    : abil.row[objGame.possNum_C];
  abil.row[objGame.possNum_C] = numOfGear;
  abil.row[objGame.possName_C] = abil.isGear
    ? gSimpleName(abil.row[objGame.abilNameID_C])
    : "";
  if (!abil.isGear) abil.row[objGame.possWorn_C] = false;

  abil.row[objGame.possEnc_C] = !abil.isGear
    ? ""
    : abil.row[objGame.possWorn_C]
    ? dbGear.arr[abil.dbGear_R][dbGear.wornEnc_C] * numOfGear
    : dbGear.arr[abil.dbGear_R][dbGear.itemEnc_C] * numOfGear;

  total.sumEnc += +abil.row[objGame.possEnc_C];

  const percOffMult = !abil.isGear
    ? 1
    : abil.row[objGame.possPerOff_C] === ""
    ? 1
    : (100 - abil.row[objGame.possPerOff_C]) / 100;
  const dbGearCrEa = !abil.isGear
    ? ""
    : dbGear.arr[abil.dbGear_R][dbGear.itemCR_C];
  abil.row[objGame.possCrEa_C] = !abil.isGear ? "" : dbGearCrEa;
  abil.row[objGame.possCrTot_C] = !abil.isGear
    ? ""
    : isNaN(Number(dbGearCrEa))
    ? dbGearCrEa
    : dbGearCrEa * percOffMult;
  abil.row[objGame.possIsArtifact_C] = !abil.isGear
    ? false
    : dbGear.arr[abil.dbGear_R][dbGear.isArtifact_C];
  abil.row[objGame.possAPEa_C] = !abil.isGear
    ? ""
    : dbGear.arr[abil.dbGear_R][dbGear.apCost_C];
  abil.row[objGame.possAPTot_C] = !abil.isGear
    ? ""
    : numOfGear * dbGear.arr[abil.dbGear_R][dbGear.apCost_C];

  total.sumAP += +abil.row[objGame.possAPTot_C];
} // end fFillGameAbilAndGearRow

// fClearGameAbilAndGearRow //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Clear's a row on <Game> both the ability and gear table completely if isElem (e.g. it is a non MyAbil and non Gear Element) or if the abilNameID_C is empty, else clears tables as a custom ability
function fClearGameAbilAndGearRow(objGame, r, isElem) {
  abilRow = objGame.arr[r];
  const clearAll = isElem || abilRow[objGame.abilNameID_C] === "";

  // Clear Abil Table
  abilRow[objGame.permMorph1_C] = "";
  if (clearAll) abilRow[objGame.morph1_C] = ",";
  if (clearAll) abilRow[objGame.sk1Typ_C] = "";
  if (clearAll) abilRow[objGame.sk1_C] = "";
  if (clearAll) abilRow[objGame.on_C] = ".";
  abilRow[objGame.sk1ChkBox_C] = false;
  if (clearAll) abilRow[objGame.abilNameID_C] = "";
  abilRow[objGame.condition_C] = "";
  abilRow[objGame.sk2ChkBox_C] = false;
  if (clearAll) abilRow[objGame.sk2_C] = "";
  if (clearAll) abilRow[objGame.sk2Typ_C] = "";
  if (clearAll) abilRow[objGame.morph2_C] = ",";
  abilRow[objGame.permMorph2_C] = "";
  abilRow[objGame.ver_C] = "";
  abilRow[objGame.notes_C] = "";
  abilRow[objGame.pic_C] = "";
  if (clearAll) abilRow[objGame.act_C] = "";
  if (clearAll) abilRow[objGame.dur_C] = "";
  if (clearAll) abilRow[objGame.rng_C] = "";
  if (clearAll) abilRow[objGame.meta_C] = "";
  if (clearAll) abilRow[objGame.uses_C] = "";
  if (clearAll) abilRow[objGame.regain_C] = "";

  // Gear Table
  abilRow[objGame.possNum_C] = "";
  abilRow[objGame.possName_C] = "";
  abilRow[objGame.possWorn_C] = false;
  abilRow[objGame.possEnc_C] = "";
  abilRow[objGame.possCrEa_C] = "";
  abilRow[objGame.possPerOff_C] = "";
  abilRow[objGame.possCrTot_C] = "";
  abilRow[objGame.possIsArtifact_C] = false;
  abilRow[objGame.possAPEa_C] = "";
  abilRow[objGame.possAPTot_C] = "";
} // end fClearGameAbilAndGearRow

// fCSCalcMR //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Calculates MR based on Gear Tbl and "Str" and "Spe" and fills out MR info
function fCSCalcMR(objGame, encTotSum) {
  // Set Speed and Strength first so the MR&Carry table is ready
  const atrSpe = objGame.arr[objGame.agi_R][objGame.sk1_C];
  const atrStr = objGame.arr[objGame.str_R][objGame.sk1_C];
  const plusMR =
    parseInt(objGame.arr[objGame.gearPlusMR_R][objGame.gearPlusMR_C], 10) || 0;
  const plusCarry =
    parseInt(
      objGame.arr[objGame.gearPlusCarry_R][objGame.gearPlusCarry_C],
      10
    ) || 0;

  objGame.arr[objGame.gearSpe_R][objGame.gearSpeAndStr_C] = atrSpe;
  objGame.arr[objGame.gearStr_R][objGame.gearSpeAndStr_C] = atrStr;

  // Set MR Table's Enc CheckMarks
  let foundMR = false;
  let finalMR = 0;
  const mrMultArr = [1, 0.9, 0.8, 0.65, 0.5, 0.25];
  const carryMultArr = [1, 2, 4, 8, 12, 16];

  for (
    let c = objGame.gearMRCol1_C, tableCol = 0;
    c <= objGame.gearMRCol6_C;
    c += 2, tableCol++
  ) {
    objGame.arr[objGame.gearMRTbl_R][c] = Math.round(
      (6 + atrSpe / 5 + plusMR) * mrMultArr[tableCol]
    );
    objGame.arr[objGame.gearCarryTbl_R][c] = Math.round(
      (50 + atrStr + plusCarry) * carryMultArr[tableCol]
    );
    objGame.arr[objGame.gearMRCheckBox_R][c] = false;

    if (!foundMR && encTotSum <= objGame.arr[objGame.gearCarryTbl_R][c]) {
      objGame.arr[objGame.gearMRCheckBox_R][c] = true;
      foundMR = true;
      finalMR = objGame.arr[objGame.gearMRTbl_R][c];
    }
  }

  if (finalMR === 0)
    SpreadsheetApp.getUi().alert("You are over encumbered, MR is 0.");

  // Save finalMR to <Game>
  gSaveVal("mycs", "Game", "MR", "MR", finalMR);
} // End fCSCalcMR

// fCSCalcMaxSocketedItems //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Calculates the maximum socketed Items, based on Char Level, and grey's out the other rows on the cs <Game> socketed Item table
function fCSCalcMaxSocketedItems(objGame) {
  // Load cs <RaceClass> data
  const charLevel = gCharLvl();

  // Calculate and save maximum Socketed Items
  const maxSocketItems = Math.min(
    12,
    2 * Math.round(3.2 * Math.log10(charLevel + 1))
  );

  // Set Socket Item table's rows to white or light grey
  const totalCols = objGame.last_C - objGame.socketedGear_C + 1;
  if (maxSocketItems > 0) {
    objGame.ref
      .getRange(
        objGame.socketTblStart_R + 1,
        objGame.socketedGear_C + 1,
        maxSocketItems,
        totalCols
      )
      .setBackground("white");
  }
  if (maxSocketItems < 12) {
    objGame.ref
      .getRange(
        objGame.socketTblStart_R + maxSocketItems + 1,
        objGame.socketedGear_C + 1,
        12 - maxSocketItems,
        totalCols
      )
      .setBackground("lightgrey");
  }
} // END fCSCalcMaxSocketedItems

// fCSLoadMonsters //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Loads db <GMScreen> Active Monsters into cs <Game> Monster Table - Leaves Custom Monster row alone
function fCSLoadMonsters() {
  let gmScreen = getObjDBGMScreen();
  let game = getObjCSGame();

  let allMonstersFound = false;
  let db_R = gmScreen.monFirst_R;

  let cs_R = game.monsterFirst_R;

  // Clear <Game> Monster Table
  gFillArraySection(
    game.arr,
    game.monsterFirst_R,
    game.monsterLast_R,
    game.numOfMonsters_C,
    game.monsterSize_C,
    ""
  );

  // Populate <Game> Monster Table
  while (cs_R <= game.monsterLast_R) {
    if (!allMonstersFound && db_R <= gmScreen.monLast_R) {
      if (gmScreen.arr[db_R][gmScreen.monActiveTF_C]) {
        // Monster is active, proceed with copying its details
        game.arr[cs_R][game.monsterName_C] = gmScreen.arr[db_R][
          gmScreen.monCustName_C
        ]
          ? gSimpleName(gmScreen.arr[db_R][gmScreen.monCustName_C])
          : gSimpleName(gmScreen.arr[db_R][gmScreen.monName_C]);

        // Copy columns from Vit to Pic
        for (let i = 0; i <= game.monsterPic_C - game.monsterVit_C; i++) {
          game.arr[cs_R][game.monsterVit_C + i] =
            gmScreen.arr[db_R][gmScreen.monVit_C + i];
        }

        game.arr[cs_R][game.monsterSize_C] = fCSExtractMonSize(
          gmScreen.arr[db_R][gmScreen.monNotes_C]
        );
        cs_R++; // Move to next row in gameArr only if a monster was added
      }
      db_R++;
      if (db_R > gmScreen.monLast_R) allMonstersFound = true;
    } else {
      // No more monsters to process, break out of the loop
      break;
    }

    if (allMonstersFound) break;
  }

  // After loop, check if not all monsters could be added
  if (cs_R > game.monsterLast_R && !allMonstersFound) {
    SpreadsheetApp.getUi().alert(
      `Tell the GM that there are more than ${
        game.monsterLast_R - game.monsterFirst_R + 1
      } Active Monsters. And the rest are not shown.`
    );
  }

  gSaveArraySectionToSheet(
    game.ref,
    game.arr,
    game.monsterFirst_R,
    game.monsterLast_R,
    game.numOfMonsters_C,
    game.monsterSize_C
  );
} // End fCSLoadMonsters

// fCSExtractMonSize //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Extracts Monster's Size from notes and returns the result
function fCSExtractMonSize(notesText) {
  // If notesText is empty, return it as is
  if (notesText === "") return "";

  // Extract the monster size from the notes
  let monSize = "";
  if (notesText.includes("\nSIZE:")) {
    let startIndex = notesText.indexOf("\nSIZE:") + 6;
    let endIndex = notesText.indexOf("\n", startIndex);
    monSize =
      endIndex !== -1
        ? notesText.substring(startIndex, endIndex)
        : notesText.substring(startIndex);
  } else {
    return notesText;
  }

  // Trim any preceding or trailing blanks
  monSize = monSize.trim();

  // Return the extracted monster size
  return monSize;
} // End fCSExtractMonSize

// fCSGMAward //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Rolls GM Award
function fCSGMAward() {
  //DB Constants
  const critArr = gArr("db", "Crit");
  const gmAward_C = gHeaderC("db", "Crit", "GMAward");
  const dataFirst_R = gDataFirst_R("db", "Crit");
  const dataLast_R = gGetVal("db", "Crit", "maxRow", gmAward_C);

  // Pick a random result
  const randomRow =
    Math.floor(Math.random() * (dataLast_R - dataFirst_R + 1)) + dataFirst_R;
  const gmAwardTxt = critArr[randomRow][gmAward_C];

  fCSPostMyRoll(`${g.bluBS}GM Award${g.endS}<br>${gmAwardTxt}<br>`);
} // End fGMAfCSGMAwardward

// fCSRest //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Calcs Rest: Meta Flood
function fCSRest() {
  fCSMetaFlood();
  fCSPostMyRoll("-- REST --");
} // End fCSRest

// fCSSleep //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Calcs Sleep: Meta Flood, Natural Healing
function fCSSleep() {
  fCSMetaFlood();
  fCSNaturalHealing();
  fCSPostMyRoll("-- SLEEP --");
} // End fCSSleep

// fCSNewGameSession //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: New Game Session: = Luck to Max
function fCSNewGameSession() {
  fCSLuckToMax();
} // End fCSNewGameSession

// fCSCalcVitMax //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Calculates Max Vit
function fCSCalcVitMax() {
  // Load <Game> data
  let game = getObjCSGame();

  const vitPlus = parseInt(game.arr[game.vitTbl_R][game.vitPlus_C] || 0);
  const atrVit = parseInt(game.arr[game.for_R][game.sk1_C] || 0);

  const vitMax = Math.round(atrVit + vitPlus);

  gSaveVal("mycs", "game", game.vitTbl_R, game.vitMax_C, vitMax);
} // End fCSCalcVitMax

// fCSClearRollLog //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Clears the roll log
function fCSClearRollLog() {
  let game = getObjCSGame();

  // Clear Roll Log cells
  gPutDocCache("rollLog", "");

  // Once mepty, run the full fPostMyRoll so it clears the Roll Result and GMScreen
  fCSPostMyRoll(`${g.bluBS}Roll Log Cleared${g.endS}`);
} // End fCSClearRollLog

// fCSClearMonsters //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Clears Monsters including custom Monster
function fCSClearMonsters() {
  let game = getObjCSGame();

  // Clear Monster Table
  gFillArraySection(
    game.arr,
    game.custMon_R,
    game.monsterLast_R,
    game.numOfMonsters_C,
    game.monsterSize_C,
    ""
  );

  // Save Monsters to <Game>
  gSaveArraySectionToSheet(
    game.ref,
    game.arr,
    game.custMon_R,
    game.monsterLast_R,
    game.numOfMonsters_C,
    game.monsterSize_C
  );
} // End fCSClearMonsters

// fCSClearMorphs //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Clears All Morphs on <Game> to just ','
function fCSClearMorphs() {
  // Set cs <Game> data
  const game = getObjCSGame();

  // Clear Morph Tables
  gFillArraySection(
    game.arr,
    game.dataFirst_R,
    game.abilTblEnd_R,
    game.morph1_C,
    game.morph1_C,
    ","
  );
  gFillArraySection(
    game.arr,
    game.dataFirst_R,
    game.abilTblEnd_R,
    game.morph2_C,
    game.morph2_C,
    ","
  );
  gFillArraySection(
    game.arr,
    game.gearTblStart_R,
    game.gearTblEnd_R,
    game.morph1_C,
    game.morph1_C,
    ","
  );
  gFillArraySection(
    game.arr,
    game.gearTblStart_R,
    game.gearTblEnd_R,
    game.morph2_C,
    game.morph2_C,
    ","
  );

  // Save Monsters to <Game>
  gSaveArraySectionToSheet(
    game.ref,
    game.arr,
    game.dataFirst_R,
    game.gearTblEnd_R,
    game.morph1_C,
    game.morph1_C
  );
  gSaveArraySectionToSheet(
    game.ref,
    game.arr,
    game.dataFirst_R,
    game.gearTblEnd_R,
    game.morph2_C,
    game.morph2_C
  );
} // End fCSClearMorphs

// fCSClearSk1Sk2CheckBoxes //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Clears All Ability Check Boxes (Sk1 and Sk2)
function fCSClearSk1Sk2CheckBoxes() {
  // Set cs <Game> data
  const gameObj = getObjCSGame();

  // Clear Sk1 and Sk2 Check Boxes
  gFillArraySection(
    gameObj.arr,
    gameObj.abilTblStart_R,
    gameObj.abilTblEnd_R,
    gameObj.sk1ChkBox_C,
    gameObj.sk1ChkBox_C,
    false
  );
  gFillArraySection(
    gameObj.arr,
    gameObj.abilTblStart_R,
    gameObj.abilTblEnd_R,
    gameObj.sk2ChkBox_C,
    gameObj.sk2ChkBox_C,
    false
  );
  gFillArraySection(
    gameObj.arr,
    gameObj.gearTblStart_R,
    gameObj.gearTblEnd_R,
    gameObj.sk1ChkBox_C,
    gameObj.sk1ChkBox_C,
    false
  );
  gFillArraySection(
    gameObj.arr,
    gameObj.gearTblStart_R,
    gameObj.gearTblEnd_R,
    gameObj.sk2ChkBox_C,
    gameObj.sk2ChkBox_C,
    false
  );

  // Save Sk1 and Sk2 Check Boxes to <Game>
  gSaveArraySectionToSheet(
    gameObj.ref,
    gameObj.arr,
    gameObj.abilTblStart_R,
    gameObj.gearTblEnd_R,
    gameObj.sk1ChkBox_C,
    gameObj.sk1ChkBox_C
  );
  gSaveArraySectionToSheet(
    gameObj.ref,
    gameObj.arr,
    gameObj.abilTblStart_R,
    gameObj.gearTblEnd_R,
    gameObj.sk2ChkBox_C,
    gameObj.sk2ChkBox_C
  );
} // End fCSClearSk1Sk2CheckBoxes

// fCSRestAllToMax //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Reset All To Max: = Clear Nish, Meta to Full, Channel clear, Luck to Max, Act to Max, Calc MR, Wounds Healed, clear Monster Table, Reset Daily Use Artifacts, Clear roll Log
function fCSRestAllToMax() {
  const ui = SpreadsheetApp.getUi(); // Access the UI of the Spreadsheet
  const promptReply = ui.alert(
    "This will reset",
    "Nish to blank\n" +
      "Meta to Full\n" +
      "Channeled Meta to blank\n" +
      "Luck to Max\n" +
      "Act to Max\n" +
      "Calc MR\n" +
      "Calc Vit Max\n" +
      "All Wounds Healed\n" +
      "Clear Sk1 & Sk2 Check Boxes\n" +
      "Clear Monster Table\n" +
      "Clear Roll Log\n\n" +
      "Continue?",
    ui.ButtonSet.YES_NO
  );

  if (promptReply !== ui.Button.YES)
    throw new Error("Operation canceled by user.");

  const updateRollLog = false;

  fCSClearNish(updateRollLog);
  fCSRefreshAbilities(updateRollLog);
  fCSMetaToFull(updateRollLog);
  fCSMetaChnlClear();
  fCSLuckToMax(updateRollLog);
  fCSActToMax();
  fCSCalcVitMax();
  fCSHealAll();
  fCSClearSk1Sk2CheckBoxes();
  fCSClearMonsters();
  fCSClearMorphs();
  fCSClearRollLog();
} // End fCSRestAllToMax

// fCSActToMax //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Sets Act (Actions) to Max
function fCSActToMax() {
  //Set cs <Game> Data
  const gameArr = gArr("mycs", "game");
  const gameAct_R = gKeyR("mycs", "game", "Act");
  const gameActPlus_R = gKeyR("mycs", "game", "ActPlus");
  const gameAct_C = gHeaderC("mycs", "game", "ActTot");

  let actPlus = gameArr[gameActPlus_R][gameAct_C];
  if (actPlus === "") actPlus = 0;

  if (!Number.isInteger(actPlus) || actPlus < 0)
    throw new Error(
      `"Act +" must be 0 or a positive integer, but is currently "${actPlus}"`
    );

  const maxAct = 5 + actPlus;

  gSaveVal("mycs", "game", gameAct_R, gameAct_C, maxAct);
} // End fCSActToMax

// fCSClearNish //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Clears Nish
function fCSClearNish(updateRollLog = true) {
  //Set cs <Game> Data
  const gameRef = gSheetRef("mycs", "game");
  const gameArr = gArr("mycs", "game");
  const gameNish_R = gKeyR("mycs", "game", "Nish");
  const gameNish_C = gHeaderC("mycs", "game", "Nish");

  gSaveVal("mycs", "game", gameNish_R, gameNish_C, "");

  if (updateRollLog) fCSPostMyRoll(`${g.bluBS}Nish Box Cleared${g.endS}<br>`);
} // End fCSClearNish

// fCSHealAll //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Removes all 1st, 2nd, 3rd deg wounds
function fCSHealAll() {
  // Load <Game> data
  const gameRef = gSheetRef("mycs", "game");
  const gameArr = gArr("mycs", "game");
  const vitTbl_R = gKeyR("mycs", "game", "VitTbl");
  const vit1st_C = gHeaderC("mycs", "game", "vit1st");
  const vit3rd_C = gHeaderC("mycs", "game", "vit3rd");

  // Save and Reload gameObj
  gFillArraySection(gameArr, vitTbl_R, vitTbl_R, vit1st_C, vit3rd_C, 0);
  gSaveArraySectionToSheet(
    gameRef,
    gameArr,
    vitTbl_R,
    vitTbl_R,
    vit1st_C,
    vit3rd_C
  );
  getObjCSGame(true);

  fCSPostMyRoll(`${g.bluBS}All Wounds Healed${g.endS}<br>`);
} // End fCSHealAll

// fCSNaturalHealing //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Applies a Natural Healing Roll
function fCSNaturalHealing(updateRollLog = true) {
  // Load <Game> data
  const gameRef = gSheetRef("mycs", "game");
  const gameArr = gArr("mycs", "game");
  const vitTbl_R = gKeyR("mycs", "game", "VitTbl");
  const vitMax_C = gHeaderC("mycs", "game", "vitMax");
  const vit1st_C = gHeaderC("mycs", "game", "vit1st");
  const vit2nd_C = gHeaderC("mycs", "game", "vit2nd");
  const vit3rd_C = gHeaderC("mycs", "game", "vit3rd");

  let vitMax = gameArr[vitTbl_R][vitMax_C];
  let vit1stWnds = gameArr[vitTbl_R][vit1st_C];
  let vit2ndWnds = gameArr[vitTbl_R][vit2nd_C];
  let vit3rdWnds = gameArr[vitTbl_R][vit3rd_C];

  let heal1st = Math.min(
    vit1stWnds,
    Math.round((fCSd(vitMax) + fCSd(vitMax)) / 4)
  );
  let heal2nd = Math.min(
    vit2ndWnds,
    Math.round((fCSd(vitMax) + fCSd(vitMax)) / 8)
  );
  let heal3rd = Math.min(
    vit3rdWnds,
    Math.round((fCSd(vitMax) + fCSd(vitMax)) / 16)
  );

  gameArr[vitTbl_R][vit1st_C] -= heal1st;
  gameArr[vitTbl_R][vit2nd_C] -= heal2nd;
  gameArr[vitTbl_R][vit3rd_C] -= heal3rd;

  // Save and Reload gameObj
  gSaveArraySectionToSheet(
    gameRef,
    gameArr,
    vitTbl_R,
    vitTbl_R,
    vit1st_C,
    vit3rd_C
  );
  getObjCSGame(true);

  if (updateRollLog && vit1stWnds + vit2ndWnds + vit3rdWnds > 0) {
    fCSPostMyRoll(
      `${g.bluBS}Natural Healing:${g.endS}<br>${heal1st} -> 1st Deg<br>${heal2nd} -> 2nd Deg<br>${heal3rd} -> 3rd Deg<br>`
    );
  }
} // End fCSNaturalHealing

// fCSLuckToMax //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Sets Luck to Max
function fCSLuckToMax(updateRollLog = true) {
  // Load <Game> data
  let game = getObjCSGame();

  const luckPlus = parseInt(game.arr[game.luckBoxPlus_R][game.luckBox_C] || 0);
  if (!Number.isInteger(luckPlus) || luckPlus < 0)
    throw new Error(
      `"Luck +" must be 0 or a positive integer, but is currently "${luckPlus}"`
    );

  const luckVerString = game.arr[game.arc_R][game.ver_C] || 1;
  const luckVer = parseInt(luckVerString.substring(1));

  const finalLuckBoxes = 4 + luckVer + luckPlus;

  gSaveVal("mycs", "game", game.luckBox_R, game.luckBox_C, finalLuckBoxes);

  if (updateRollLog)
    fCSPostMyRoll(`${g.bluBS}Luck to Max:${g.endS} ${finalLuckBoxes}<br>`);
} // End fCSLuckToMax

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  Gear (end g. game)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// fCSShowGear //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Shows or Hides the columns of the Gear Tbl on <Game>
function fCSShowGear(colShow) {
  const objGame = getObjCSGame();
  const firstCol = objGame.possNum_C + 1;
  const lastCol = objGame.possAPTot_C + 2; // Also hides the column after Gear Tbl
  const numCols = lastCol - firstCol + 1;

  if (colShow) {
    objGame.ref.showColumns(firstCol, numCols);
  } else {
    objGame.ref.hideColumns(firstCol, numCols);
  }
} // End fCSShowGear

// fCSRollRndTreasure //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Rolls random treasure
function fCSRollRndTreasure() {
  const roll = fCSd(12);
  const rnd = true;

  // Select the treasure to roll based on the roll
  switch (true) {
    case roll <= 2:
      fCSRollValuables(rnd);
      break;
    case roll <= 4:
      fCSRollGear(rnd);
      break;
    case roll <= 6:
      fCSChaosCrystal(rnd);
      break;
    case roll <= 8:
      fCSRollSocketGem(rnd);
      break;
    case roll <= 10:
      fCSRollArtifact(rnd);
      break;
    default:
      fCSRollSockets(rnd);
  }
} // End fCSRollRndTreasure

// fCSChaosCrystal //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Rolls a Choas Crystal
function fCSChaosCrystal(rnd = false) {
  //DB Constants
  const critArr = gArr("db", "Crit");
  const chaosCrystalMaxUses_C = gHeaderC("db", "Crit", "ChaosCrytstalMaxUses");
  const chaosCrystal_C = gHeaderC("db", "Crit", "ChaosCrystal");
  const dataFirst_R = gDataFirst_R("db", "Crit");
  const dataLast_R = gGetVal("db", "Crit", "maxRow", chaosCrystal_C);

  // Pick a random result
  const randomRow =
    Math.floor(Math.random() * (dataLast_R - dataFirst_R + 1)) + dataFirst_R;
  const chaosCrystalUses = fCSd(critArr[randomRow][chaosCrystalMaxUses_C]);
  let chaosCrystalAbility = critArr[randomRow][chaosCrystal_C];
  chaosCrystalAbility = chaosCrystalAbility.replace(/[\r\n]/g, " ");

  fCSPostMyRoll(
    `${rnd ? "Loot:<br>" : "--- SPECIFIC TREASURE ---<br>"}${
      g.bluBS
    }Chaos Crystal${g.endS} - ${
      chaosCrystalUses === 1 ? "Use" : "Uses"
    }(${chaosCrystalUses}): <br>${chaosCrystalAbility}<br>`
  );
} // End fCSChaosCrystal

// fCSRollArtifact //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Rolls an Artifact
function fCSRollArtifact(rnd = false) {
  // CS Constants
  const charLvl = gCharLvl();

  // DB Constants
  const gearArr = gArr("db", "Gear");
  const isArtifact_C = gHeaderC("db", "Gear", "IsArtifact");
  const artName_C = gHeaderC("db", "Gear", "Name");
  const apCost_C = gHeaderC("db", "Gear", "AP");
  const artDataFirst_R = gDataFirst_R("db", "Gear");
  let artDataLast_R = gDataLast_R("db", "Gear");

  // Filter gearArr to keep only artifacts and update artDataLast_R
  for (let r = artDataFirst_R; r <= artDataLast_R; r++) {
    if (!gearArr[r][isArtifact_C]) {
      gearArr.splice(r, 1); // Remove the non-artifact row
      r--; // Adjust index due to row removal
      artDataLast_R--; // Decrement last row index
    }
  }

  // Pick a random result that meets the AP cost condition but stop on any result after 10 tries
  let loopCount = 0;
  let randomRow;
  do {
    randomRow =
      Math.floor(Math.random() * (artDataLast_R - artDataFirst_R + 1)) +
      artDataFirst_R;
    loopCount++;
  } while (
    gearArr[randomRow][apCost_C] > Math.max(5, charLvl) &&
    loopCount < 10
  );

  const artName = gearArr[randomRow][artName_C];
  const apCost = gearArr[randomRow][apCost_C];

  fCSPostMyRoll(
    `${rnd ? "Loot:<br>" : "--- SPECIFIC TREASURE ---<br>"}${
      g.bluBS
    }${artName}${g.endS} (${apCost}AP)<br>`
  );
} // End fCSRollArtifact

// fCSRollGear //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Rolls an Artifact
function fCSRollGear(rnd = false) {
  // DB Constants
  const gearArr = gArr("db", "Gear");
  const isArtifact_C = gHeaderC("db", "Gear", "IsArtifact");
  const itemCr_C = gHeaderC("db", "Gear", "ItemCR");
  const itemName_C = gHeaderC("db", "Gear", "Name");
  const itemDataFirst_R = gDataFirst_R("db", "Gear");
  let itemDataLast_R = gDataLast_R("db", "Gear");

  // Set CS Data
  const charLevel = gCharLvl();
  const maxCr =
    charLevel <= 4
      ? 100
      : charLevel <= 9
      ? 250
      : charLevel <= 20
      ? 1000
      : 10000;

  // Keep gearArr rows above artDataFirst_R but sort gearArr between rows artDataFirst_R and artDataLast_R keeping only those rows for which gearArr[r][isArtifact_C] === false and gearArr[r][itemCr_C] is a number less than maxCr, Decrement artDataLast_R for each row eliminated
  for (let r = itemDataFirst_R; r <= itemDataLast_R; r++) {
    const itemCr = parseInt(gearArr[r][itemCr_C]);
    if (gearArr[r][isArtifact_C] || isNaN(itemCr) || itemCr > maxCr) {
      gearArr.splice(r, 1); // Remove the row
      r--; // Adjust index to account for removed row
      itemDataLast_R--; // Decrement last row index
    }
  }

  // Pick a random result
  const randomRow =
    Math.floor(Math.random() * (itemDataLast_R - itemDataFirst_R + 1)) +
    itemDataFirst_R;
  const itemName = gearArr[randomRow][itemName_C];
  const conditionArr = [
    "Useless",
    "Broken",
    "Used",
    "Good",
    "New",
    "High Quality (Cr*2)",
  ];
  // Random selection of conditionArr
  const condition =
    conditionArr[Math.floor(Math.random() * conditionArr.length)];

  fCSPostMyRoll(
    `${rnd ? "Loot:<br>" : "--- SPECIFIC TREASURE ---<br>"}${
      g.bluBS
    }${itemName}${g.endS}<br>Condition: ${condition}<br>`
  );
} // End fCSRollGear

// fCSRollSockets //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Rolls a socketed item with socket colors
function fCSRollSockets(rnd = false) {
  // Set CS Data
  const charLevel = gCharLvl();
  const socketColors = ["R", "O", "Y", "G", "B"];

  // Determine maxGemColor based on charLevel (which is socketColors index + 1)
  const maxGemColor =
    charLevel <= 4 ? 2 : charLevel <= 9 ? 3 : charLevel <= 20 ? 4 : 5;

  // Determine maxGemSlots based on the cube root of charLevel
  const maxGemSlots = Math.floor(Math.cbrt(charLevel));

  const numGemSlots = Math.min(fCSd(maxGemSlots), fCSd(maxGemSlots));
  let slotColorArr = [];

  for (let i = 0; i < numGemSlots; i++) {
    let colorIndex = Math.min(fCSd(maxGemColor) - 1, fCSd(maxGemColor) - 1);
    slotColorArr.push(socketColors[colorIndex]);
  }

  // Select Socketed Item
  const weaponArr = [
    "Rnd Owned Body Weapon",
    "Selected Body Weapon",
    "Rnd Owned Melee Weapon",
    "Selected Melee Weapon",
    "Rnd Owned Hurled Weapon",
    "Selected Hurled Weapon",
    "Rnd Owned Ranged Weapon",
    "Selected Ranged Weapon",
  ];
  const itemArr = [
    "Boots",
    "Gloves",
    "Gauntlet",
    "Rnd Owned Armor",
    "Selected Armor",
    "Pants",
    "Belt",
    "Tunic/Shirt",
    "Ring",
    "Arm Band",
    "Helmet",
    "Tiara",
    "Necklace",
    "Amulet",
    "Pendant/Broach",
  ];

  // Randomly choose between weaponArr and itemArr
  const itemRolled =
    fCSd(3) === 1
      ? weaponArr[Math.floor(Math.random() * weaponArr.length)]
      : itemArr[Math.floor(Math.random() * itemArr.length)];

  // Convert slotColorArr to a comma-separated string
  const slotColorsText = slotColorArr.join(", ");

  // Post the result
  fCSPostMyRoll(
    `${rnd ? "Loot:<br>" : "--- SPECIFIC TREASURE ---<br>"}${
      g.bluBS
    }Socketed Item${g.endS}:<br>${itemRolled} with ${slotColorsText} ${
      numGemSlots === 1 ? "slot" : "slots"
    }<br>`
  );
} // End fCSRollSockets

// fCSRollSocketGem //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Rolls level-appropriate Gems
function fCSRollSocketGem(rnd = false) {
  // DB Constants
  const critArr = gArr("db", "Crit");
  const gemColor_C = gHeaderC("db", "Crit", "GemColor");
  const gemDesc_C = gHeaderC("db", "Crit", "GemDesc");
  const dataFirst_R = gDataFirst_R("db", "Crit");
  const dataLast_R = gGetVal("db", "Crit", "maxRow", gemDesc_C);

  // Set CS Data
  const charLevel = gCharLvl();
  const socketColors = ["R", "O", "Y", "G", "B"];

  // Determine maxGemColor based on charLevel (which is socketColors index + 1)
  let maxGemColor =
    charLevel <= 4 ? 2 : charLevel <= 9 ? 3 : charLevel <= 20 ? 4 : 5;
  maxGemColor = Math.min(fCSd(maxGemColor) - 1, fCSd(maxGemColor) - 1);
  let gemFound = false;
  let gemFound_R;

  // Loop until gemFound is true
  while (!gemFound) {
    let r =
      Math.floor(Math.random() * (dataLast_R - dataFirst_R + 1)) + dataFirst_R;
    if (socketColors.indexOf(critArr[r][gemColor_C]) <= maxGemColor) {
      gemFound = true;
      gemFound_R = r;
    }
  }

  fCSPostMyRoll(
    `${rnd ? "Loot:<br>" : "--- SPECIFIC TREASURE ---<br>"}${
      g.bluBS
    }Socket Gem:${g.endS}<br>${critArr[gemFound_R][gemColor_C]} Gem:<br>${
      critArr[gemFound_R][gemDesc_C]
    }<br>`
  );
} // End fCSRollSocketGem

// fCSRollValuables //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Rolls random valuables (credits, art, jewelry, notes, trash, etc.)
function fCSRollValuables(rnd = false) {
  // DB Constants
  const critArr = gArr("db", "Crit");
  const treasureColor_C = gHeaderC("db", "Crit", "TreasureColor");
  const treasureDesc_C = gHeaderC("db", "Crit", "Treasure");
  const dataFirst_R = gDataFirst_R("db", "Crit");
  const dataLast_R = gGetVal("db", "Crit", "maxRow", treasureDesc_C);

  // Set CS Data
  const charLevel = gCharLvl();
  const treasureColorArr = ["R", "O", "Y", "G", "B"];

  // Determine maxTreasureColor based on charLevel (which is treasureColorArr index + 1)
  let maxTreasureColor =
    charLevel <= 4 ? 2 : charLevel <= 9 ? 3 : charLevel <= 20 ? 4 : 5;
  maxTreasureColor = Math.min(
    fCSd(maxTreasureColor) - 1,
    fCSd(maxTreasureColor) - 1
  );
  let treasureFound = false;
  let treasureFound_R;

  // Loop until treasureFound is true
  while (!treasureFound) {
    let r =
      Math.floor(Math.random() * (dataLast_R - dataFirst_R + 1)) + dataFirst_R;
    if (
      treasureColorArr.indexOf(critArr[r][treasureColor_C]) <= maxTreasureColor
    ) {
      treasureFound = true;
      treasureFound_R = r;
    }
  }

  fCSPostMyRoll(
    `${rnd ? "Loot:<br>" : "--- SPECIFIC TREASURE ---<br>"}${g.bluBS}Treasure:${
      g.endS
    }<br>${critArr[treasureFound_R][treasureDesc_C]}<br>`
  );
} // End fCSRollValuables

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  RaceClass (end g. gear)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// fCSRefreshRaceClass //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Builds the Race dropdown
function fCSRefreshRaceClass() {
  let game = getObjCSGame();

  // Get rollLog
  const oldRollLog = gGetDocCache("rollLog") || "";
  const rollLog = oldRollLog === "" ? fCSCreateObjRollLog() : oldRollLog;

  const loadAllTF = true;
  fCSLoadRollLogStats(rollLog, game, loadAllTF);

  gPutDocCache("rollLog", rollLog);

  fCSPostMyRoll(`${g.bluBS}RaceClass Tab${g.endS}<br>Data Saved<br>`);
} // End fCSRefreshRaceClass

// fCSBuildRaceDropDown //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Builds the Race dropdown
function fCSBuildRaceDropDown() {
  // db <Race> constants
  const dbRaceArr = gArr("db", "race");
  const dbRaceFirst_R = gDataFirst_R("db", "race");
  const dbRaceLast_R = gDataLast_R("db", "race");
  const dbRaceNameID_C = gHeaderC("db", "race", "Name_ID");

  // cs <RaceClass> constants
  const csRCRef = gSheetRef("mycs", "raceclass");
  const csRCArr = gArr("mycs", "raceclass");
  const csRCRace_R = gKeyR("mycs", "raceclass", "race");
  const csRCRace_C = gHeaderC("mycs", "raceclass", "Val");
  const csRCNotes_C = gHeaderC("mycs", "raceclass", "Notes");
  const csRCPic_C = gHeaderC("mycs", "raceclass", "Pic");

  let raceDropDown = ["."];

  // Loop through dbRaceArr and push names onto raceDropDown
  for (let r = dbRaceFirst_R; r <= dbRaceLast_R; r++) {
    if (dbRaceArr[r][dbRaceNameID_C]) {
      raceDropDown.push(dbRaceArr[r][dbRaceNameID_C]);
    }
  }

  // Alphabetize raceDropDown after element 0
  raceDropDown = [raceDropDown[0]].concat(raceDropDown.slice(1).sort());

  // Create a dropdown and save raceDropDown to csRCRef
  // Get the cell where the dropdown will be set
  let dropdownCell = csRCRef.getRange(csRCRace_R + 1, csRCRace_C + 1);

  // Create a data validation rule with the raceDropDown array
  let validationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(raceDropDown, true)
    .build();

  // Save the drop down and set its initial value to '.'
  dropdownCell.setDataValidation(validationRule);
  dropdownCell.setValue(".");

  // Clear Race selection, Notes, Pic
  csRCArr[csRCRace_R][csRCNotes_C] = "";
  csRCArr[csRCRace_R][csRCPic_C] = "";

  gSaveArraySectionToSheet(
    csRCRef,
    csRCArr,
    csRCRace_R,
    csRCRace_R,
    csRCNotes_C,
    csRCPic_C
  );
} // End fCSBuildRaceDropDown

// fCSUpdateRaceNotesPic //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Update's Race's Notes and Pic
function fCSUpdateRaceNotesPic() {
  // db <Race> constants
  const dbRaceArr = gArr("db", "race");
  const dbRaceNotes_C = gHeaderC("db", "race", "Notes");
  const dbRacePic_C = gHeaderC("db", "race", "Pic");

  // cs <RaceClass> constants
  const csRCRef = gSheetRef("mycs", "raceclass");
  const csRCArr = gArr("mycs", "raceclass");
  const csRCRace_R = gKeyR("mycs", "raceclass", "race");
  const csRCRace_C = gHeaderC("mycs", "raceclass", "Val");
  const csRCNotes_C = gHeaderC("mycs", "raceclass", "Notes");
  const csRCPic_C = gHeaderC("mycs", "raceclass", "Pic");

  const myRaceNameID = csRCArr[csRCRace_R][csRCRace_C];
  const raceID = gTestID("db", "race", myRaceNameID);
  if (!raceID)
    throw new Error(
      `Choose a new race from the dropdown list. I cannot find the race "${myRaceNameID}"`
    );
  const dbRace_R = gKeyR("db", "race", raceID);

  csRCArr[csRCRace_R][csRCNotes_C] = dbRaceArr[dbRace_R][dbRaceNotes_C];
  csRCArr[csRCRace_R][csRCPic_C] = dbRaceArr[dbRace_R][dbRacePic_C];

  gSaveArraySectionToSheet(
    csRCRef,
    csRCArr,
    csRCRace_R,
    csRCRace_R,
    csRCNotes_C,
    csRCPic_C
  );
} // End fCSUpdateRaceNotesPic

// fCSGenerateRaceGenderToVision //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Calls fUpdateRaceNotesPic, Then Fills in Gender to Vision
function fCSGenerateRaceGenderToVision() {
  // Update Notes and Pic
  fCSUpdateRaceNotesPic();

  // db <Race> constants
  const dbRaceArr = gArr("db", "race");
  const dbAgeRoll_C = gHeaderC("db", "race", "AgeRoll");
  const dbAgeMaxRoll_C = gHeaderC("db", "race", "AgeMaxRoll");
  const dbAgeStatic_C = gHeaderC("db", "race", "AgeStatic");
  const dbAgeSuffix_C = gHeaderC("db", "race", "AgeSuffix");
  const dbOldAgeCheck_C = gHeaderC("db", "race", "OldAgeCheck");
  const dbGenderChoices_C = gHeaderC("db", "race", "GenderChoices");
  const dbHeightRoll_C = gHeaderC("db", "race", "HeightRoll");
  const dbHeightMaxRoll_C = gHeaderC("db", "race", "HeightMaxRoll");
  const dbHeightStatic_C = gHeaderC("db", "race", "HeightStatic");
  const dbHeightSuffix_C = gHeaderC("db", "race", "HeightSuffix");
  const dbWeightRoll_C = gHeaderC("db", "race", "WeightRoll");
  const dbWeightMaxRoll_C = gHeaderC("db", "race", "WeightMaxRoll");
  const dbWeightStatic_C = gHeaderC("db", "race", "WeightStatic");
  const dbWeightSuffix_C = gHeaderC("db", "race", "WeightSuffix");
  const dbOtherRoll_C = gHeaderC("db", "race", "OtherRoll");
  const dbOtherMaxRoll_C = gHeaderC("db", "race", "OtherMaxRoll");
  const dbOtherSuffix_C = gHeaderC("db", "race", "OtherSuffix");
  const dbEyeChoices_C = gHeaderC("db", "race", "EyeChoices");
  const dbHairChoices_C = gHeaderC("db", "race", "HairChoices");
  const dbHairSuffix_C = gHeaderC("db", "race", "HairSuffix");
  const dbDiet_C = gHeaderC("db", "race", "Diet");
  const dbSleep_C = gHeaderC("db", "race", "Sleep");
  const dbVision_C = gHeaderC("db", "race", "Vision");

  // cs <RaceClass> constants
  const csRCRef = gSheetRef("mycs", "raceclass");
  const csRCArr = gArr("mycs", "raceclass");
  const csRCRace_R = gKeyR("mycs", "raceclass", "race");
  const csRCVal_C = gHeaderC("mycs", "raceclass", "Val");
  const csRCGender_R = gKeyR("mycs", "raceclass", "gender");
  const csRCHair_R = gKeyR("mycs", "raceclass", "hair");
  const csRCEyes_R = gKeyR("mycs", "raceclass", "eyes");
  const csRCAge_R = gKeyR("mycs", "raceclass", "age");
  const csRCOldAge_R = gKeyR("mycs", "raceclass", "oldAge");
  const csRCHeight_R = gKeyR("mycs", "raceclass", "height");
  const csRCWeight_R = gKeyR("mycs", "raceclass", "weight");
  const csRCOther_R = gKeyR("mycs", "raceclass", "other");
  const csRCDiet_R = gKeyR("mycs", "raceclass", "diet");
  const csRCSleep_R = gKeyR("mycs", "raceclass", "sleep");
  const csRCVision_R = gKeyR("mycs", "raceclass", "vision");

  const myRaceNameID = csRCArr[csRCRace_R][csRCVal_C];
  const raceID = gTestID("db", "race", myRaceNameID);
  if (!raceID)
    throw new Error(
      `Choose a new race from the dropdown list. I cannot find the race "${myRaceNameID}"`
    );
  const dbRace_R = gKeyR("db", "race", raceID);

  csRCArr[csRCGender_R][csRCVal_C] = fCSDDmaker(
    csRCGender_R,
    csRCVal_C,
    dbRaceArr,
    dbRace_R,
    dbGenderChoices_C
  );
  csRCArr[csRCHair_R][csRCVal_C] = fCSDDmaker(
    csRCHair_R,
    csRCVal_C,
    dbRaceArr,
    dbRace_R,
    dbHairChoices_C,
    dbHairSuffix_C
  );
  csRCArr[csRCEyes_R][csRCVal_C] = fCSDDmaker(
    csRCEyes_R,
    csRCVal_C,
    dbRaceArr,
    dbRace_R,
    dbEyeChoices_C
  );
  csRCArr[csRCAge_R][csRCVal_C] = fCSStatRoller(
    dbRaceArr,
    dbRace_R,
    dbAgeRoll_C,
    dbAgeMaxRoll_C,
    dbAgeStatic_C,
    dbAgeSuffix_C
  );
  csRCArr[csRCOldAge_R][csRCVal_C] = dbRaceArr[dbRace_R][dbOldAgeCheck_C];
  csRCArr[csRCHeight_R][csRCVal_C] = fCSStatRoller(
    dbRaceArr,
    dbRace_R,
    dbHeightRoll_C,
    dbHeightMaxRoll_C,
    dbHeightStatic_C,
    dbHeightSuffix_C
  );
  csRCArr[csRCWeight_R][csRCVal_C] = fCSStatRoller(
    dbRaceArr,
    dbRace_R,
    dbWeightRoll_C,
    dbWeightMaxRoll_C,
    dbWeightStatic_C,
    dbWeightSuffix_C
  );
  csRCArr[csRCOther_R][csRCVal_C] = fCSStatRoller(
    dbRaceArr,
    dbRace_R,
    dbOtherRoll_C,
    dbOtherMaxRoll_C,
    "",
    dbOtherSuffix_C
  );
  csRCArr[csRCDiet_R][csRCVal_C] = dbRaceArr[dbRace_R][dbDiet_C];
  csRCArr[csRCSleep_R][csRCVal_C] = dbRaceArr[dbRace_R][dbSleep_C];
  csRCArr[csRCVision_R][csRCVal_C] = dbRaceArr[dbRace_R][dbVision_C];

  gSaveArraySectionToSheet(
    csRCRef,
    csRCArr,
    csRCGender_R,
    csRCVision_R,
    csRCVal_C,
    csRCVal_C
  );
} // End fCSGenerateRaceGenderToVision

// fCSDDmaker //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Creates Gender or Hair or Eyes dropdown and returns the choice of '.'
function fCSDDmaker(cs_R, cs_C, dbArr, db_R, db_C, dbSuffix_C = "") {
  let choicesArr = ["."];

  // Concatenate choices with the split values from dbArr
  choicesArr = choicesArr.concat(dbArr[db_R][db_C].split("|"));

  // If dbSuffix_C is provided, add it to all but the '.' choice
  const suffix = dbSuffix_C ? " " + dbArr[db_R][dbSuffix_C] : "";
  if (suffix) {
    for (let i = 1, len = choicesArr.length; i < len; i++) {
      choicesArr[i] += suffix;
    }
  }

  // Alphabetize choicesArr after element 0
  choicesArr = [choicesArr[0]].concat(choicesArr.slice(1).sort());

  // Create a dropdown and save choicesArr to csRCRef
  const csRCRef = gSheetRef("mycs", "raceclass");
  // Get the cell where the dropdown will be set
  let dropdownCell = csRCRef.getRange(cs_R + 1, cs_C + 1);

  // Create a data validation rule with the choicesArr array
  let validationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(choicesArr, true)
    .build();

  // Save the dropdown and set its initial value to '.'
  dropdownCell.setDataValidation(validationRule);
  dropdownCell.setValue(".");

  return ".";
} // End fCSDDmaker

// fCSStatRoller //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Rolls many of the Racial stats such as age, height, weight, etc.
function fCSStatRoller(
  dbArr,
  db_R,
  dbRoll_C,
  dbMax_C,
  dbStatic_C = "",
  dbSuffix_C = ""
) {
  // If a static value exists, use it
  let rollResult = dbStatic_C ? dbArr[db_R][dbStatic_C] : "";

  // If not static && there is something to roll (not a blank OtherRoll) then do all this..
  if (!rollResult && dbArr[db_R][dbRoll_C]) {
    let rollStr = dbArr[db_R][dbRoll_C];
    let rollMax = dbArr[db_R][dbMax_C];

    // Remove whitespace from rollStr
    rollStr = rollStr.replace(/\s+/g, "");

    // Verify rollStr format is like "D8+15"
    const rollRegex = /D(\d+)\+(\d+)/;
    const rollMatch = rollRegex.exec(rollStr);
    if (!rollMatch) {
      throw new Error(
        `In fCSStatRoller, the incorrect formula "${rollStr}" was passed.`
      );
    }

    // Extract rtg and bonus
    const rtg = parseInt(rollMatch[1]);
    const bonus = parseInt(rollMatch[2]);

    // Calculate rollResult and use lesser of it or rollMax
    rollResult = Math.min(fCSD(rtg) + bonus, rollMax);
  }

  // Append suffix if provided (and if it starts with 'ftin' convert rollResult to #'#")
  if (dbSuffix_C) {
    let suffixStr = dbArr[db_R][dbSuffix_C];
    const ftinTF = suffixStr.startsWith("ftin");
    if (suffixStr.includes("|")) {
      suffixStr = suffixStr.split("|")[1];
    }

    if (ftinTF) {
      rollResult = fCSConvertFtIn(rollResult);
    }
    if (suffixStr === "ftin") suffixStr = "";
    if (suffixStr) rollResult += " " + suffixStr;
  }

  return rollResult;
} // End fStatRoller

// Helper function to simulate a dice roll
function fCSD(sides) {
  return Math.floor(Math.random() * sides) + 1;
} // End fCSStatRoller

// fCSConvertFtIn //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Converts inches to feet and inches in the format 6'6"
function fCSConvertFtIn(inches) {
  // Validate input
  if (isNaN(inches) || inches < 0)
    throw new Error(
      `Error in fCSConvertFtIn: Incorrect value passed - ${inches} inches.`
    );

  // Ensure inches is treated as a number
  inches = Number(inches);

  // Convert inches to feet and inches
  const feet = Math.floor(inches / 12);
  const remainingInches = inches % 12;

  // Format the result, omitting 0 feet or 0 inches
  let result = "";
  if (feet > 0) result += `${feet}'`;
  if (remainingInches > 0) result += `${remainingInches}"`;

  return result || '0"'; // Return '0"' if both feet and inches are 0
} // End fCSConvertFtIn

// fCSCreateKL //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Runs fCSSaveCSID, then Creates a copy of Master KeyLine for this character
function fCSCreateKL() {
  fCSSaveCSID();

  // Get ss constants
  const ssRef = gSSRef("mycs");
  const ssName = ssRef.getName();
  const folder = DriveApp.getFileById(g.id.mycs).getParents().next();

  // Get <RaceClass> constants
  const raceClassRef = gSheetRef("mycs", "RaceClass");
  const raceClassKLURL_R = gKeyR("mycs", "RaceClass", "KLURL");
  const raceClassKLURL_C = gHeaderC("mycs", "RaceClass", "KLURL");
  const oldKLID = g.id.mykl;

  // Prompt user if oldKLID already is a non '' id
  if (oldKLID) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "WARNING!",
      "This will replace your current KeyLine with a new one. Do you wish to continue?",
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      throw new Error("New KeyLine creation aborted!");
    }
  }

  // Check if file with the new name already exists, if so rename it to 'Old' + name
  const newFileName = "KeyLine - " + ssName;
  const existingFiles = folder.getFilesByName(newFileName);
  if (existingFiles.hasNext()) {
    const existingFile = existingFiles.next();
    existingFile.setName("Old " + existingFile.getName());
  }

  // Copy the master keyline
  const masterKLFile = DriveApp.getFileById(g.id.masterkl);
  const newKLFile = masterKLFile.makeCopy(newFileName, folder);
  const newKLID = newKLFile.getId();
  g.id.mykl = newKLID;

  // Save the new kl's ID to cs
  gSaveVal("mycs", "data", "myKLID", "Manual", newKLID);
  const urlMyKL = gURLFromID(newKLID, "sheet");
  raceClassRef
    .getRange(raceClassKLURL_R + 1, raceClassKLURL_C + 1)
    .setFormula('=HYPERLINK("' + urlMyKL + '", "My KeyLine")');

  // Save cs ID to kl's <Data>
  const csID = gGetVal("mycs", "data", "myCSID", "Val");
  gSaveVal("mykl", "data", "myCSID", "Manual", csID);
  gSaveVal("mykl", "data", "myKLID", "Manual", newKLID);

  // Open the new spreadsheet in a new tab
  const newKLUrl =
    "https://docs.google.com/spreadsheets/d/" + newKLID + "/edit";
  const htmlOutput = HtmlService.createHtmlOutput(
    '<script>window.open("' +
      newKLUrl +
      '");google.script.host.close();</script>'
  );
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Opening KeyLine");
} // End fCSCreateKL

// fCSRefreshRuleBookURLs //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Refresh URL links for CM, EM, RB, SG, PS
function fCSRefreshRuleBookURLs() {
  // Set cs <RaceClass> constants

  const raceClassRef = gSheetRef("mycs", "RaceClass");
  const raceClassCMandRB_R = gKeyR("mycs", "RaceClass", "CMandRB");
  const raceClassEMandSG_R = gKeyR("mycs", "RaceClass", "EMandSG");
  const raceClassPS_R = gKeyR("mycs", "RaceClass", "PS");
  const raceClassCMandEM_C = gHeaderC("mycs", "RaceClass", "CMandEM");
  const raceClassRBandSG_C = gHeaderC("mycs", "RaceClass", "RBandSG");
  const raceClassPS_C = gHeaderC("mycs", "RaceClass", "PS");

  // Create Document URLs
  const urlCM = gURLFromID(g.id.cm, "doc");
  const urlEM = gURLFromID(g.id.em, "doc");
  const urlSG = gURLFromID(g.id.sg, "doc");
  const urlRB = gURLFromID(g.id.rb, "doc");
  const urlPL = gURLFromID(g.id.ps, "sheet");

  // Save Document URLs to <RaceClass>
  raceClassRef
    .getRange(raceClassCMandRB_R + 1, raceClassCMandEM_C + 1)
    .setFormula('=HYPERLINK("' + urlCM + '", "Character Manual")');
  raceClassRef
    .getRange(raceClassCMandRB_R + 1, raceClassRBandSG_C + 1)
    .setFormula('=HYPERLINK("' + urlRB + '", "Rule Book")');
  raceClassRef
    .getRange(raceClassEMandSG_R + 1, raceClassCMandEM_C + 1)
    .setFormula('=HYPERLINK("' + urlEM + '", "Equipment Manual")');
  raceClassRef
    .getRange(raceClassEMandSG_R + 1, raceClassRBandSG_C + 1)
    .setFormula('=HYPERLINK("' + urlSG + '", "Setting Guide")');
  raceClassRef
    .getRange(raceClassPS_R + 1, raceClassPS_C + 1)
    .setFormula('=HYPERLINK("' + urlPL + '", "Player Screen")');
} // End fCSRefreshRuleBookURLs

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  PermMorph (end g. Gear)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// NOTE: this section is ONLY called as a helper to <Game> menu's function fRefreshAbilities()

// fCSRefreshPermMorph //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Refreshes the <PermMorph> table
function fCSRefreshPermMorph() {
  // Load perm, game, myAbil objects
  let perm = getObjCSPermMorph();

  // Blank any empty Ability rows or rows without both a Sk1 Morph and Sk2 Morph
  fCSBlankInvalidPermRows(perm);

  // Blank any duplicate ID + Condition rows, then sort perm.arr on abilities then condition
  fCSRemoveDuplicatesAndSort(perm);

  // Calculate ID, OldSk1, NewSk1, Roll1Morph, then same for Sk2
  fCSCalcNewPermData(perm);
  fCSBlankInvalidPermRows(perm);
  fCSSortPermRows(perm);

  // Save perm.arr changes to <PermMorph>
  gSaveArraySectionToSheet(
    perm.ref,
    perm.arr,
    perm.dataFirst_R,
    perm.dataLast_R,
    perm.first_C,
    perm.last_C
  );
  perm = getObjCSPermMorph(true);

  // *** Apply to current <Game> sheet
  let game = getObjCSGame();
  fCSBuildconditionDropDown(perm, game);
  fCSApplyPermToGame(perm, game);
} // End fCSRefreshPermMorph

// fCSBlankInvalidPermRows //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Blank any empty Perm Morph Ability rows or rows without both a Sk1 Morph and Sk2 Morph
function fCSBlankInvalidPermRows(perm) {
  // Blank any empty Ability rows or rows without both a Sk1 Morph and Sk2 Morph
  for (let r = perm.dataFirst_R; r <= perm.dataLast_R; r++) {
    if (
      !perm.arr[r][perm.abilNameID_C] ||
      (!perm.arr[r][perm.morph1_C] && !perm.arr[r][perm.morph2_C])
    ) {
      gFillArraySection(perm.arr, r, r, perm.first_C, perm.last_C, "");
    }
  }
} // End fCSBlankInvalidPermRows

// fCSSortPermRows //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Sort perm.arr
function fCSSortPermRows(perm) {
  gSortArraySection(
    perm.arr,
    perm.dataFirst_R,
    perm.dataLast_R,
    perm.first_C,
    perm.last_C,
    perm.condition_C
  );
  gSortArraySection(
    perm.arr,
    perm.dataFirst_R,
    perm.dataLast_R,
    perm.first_C,
    perm.last_C,
    perm.abilNameID_C
  );
} // End fCSSortPermRows

// fCSRemoveDuplicatesAndSort //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Sort perm.arr, then Blank any duplicate ID + Condition rows
function fCSRemoveDuplicatesAndSort(perm) {
  // Sort perm.arr so abilities are contiguous with conditions in sorted order next to them
  fCSSortPermRows(perm);

  for (let r = perm.dataLast_R; r >= perm.dataFirst_R - 1; r--) {
    if (
      perm.arr[r][perm.abilNameID_C] === perm.arr[r - 1][perm.abilNameID_C] &&
      perm.arr[r][perm.condition_C] === perm.arr[r - 1][perm.condition_C]
    ) {
      gFillArraySection(perm.arr, r, r, perm.first_C, perm.last_C, "");
    }
  }

  // Re-sort after deletions to maintain order and contiguous data
  fCSSortPermRows(perm);
} // End fCSRemoveDuplicatesAndSort

// fCSCalcNewPermData //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Calculate ID, OldSk1, NewSk1, Roll1Morph, then same for Sk2
function fCSCalcNewPermData(perm) {
  let myAbil = getObjKLMyAbilities();

  // Calculate ID, OldSk1, NewSk1, Roll1Morph, then same for Sk2
  for (
    let r = perm.dataFirst_R;
    r <= perm.dataLast_R && perm.arr[r][perm.abilNameID_C];
    r++
  ) {
    const abilID = perm.arr[r][perm.abilNameID_C].slice(-6); // Assume last 6 characters are the abilID
    perm.arr[r][perm.id_C] = abilID;
    const myAbil_R = gKeyR("mykl", "MyAbilities", abilID); // returns false if not found

    if (myAbil_R) {
      perm.arr[r][perm.oldSk1_C] = myAbil.arr[myAbil_R][myAbil.trainedSk1_C];
      perm.arr[r][perm.oldSk2_C] = myAbil.arr[myAbil_R][myAbil.trainedSk2_C];

      const [newSk1Morph, newSk1, newMorphRoll1] = fCSCalcNewSkAndRollMorph(
        perm,
        r,
        "Sk1"
      );
      perm.arr[r][perm.morph1_C] = newSk1Morph;
      perm.arr[r][perm.newSk1_C] = newSk1;
      perm.arr[r][perm.morphRoll1_C] = newMorphRoll1;

      const [newSk2Morph, newSk2, newMorphRoll2] = fCSCalcNewSkAndRollMorph(
        perm,
        r,
        "Sk2"
      );
      perm.arr[r][perm.morph2_C] = newSk2Morph;
      perm.arr[r][perm.newSk2_C] = newSk2;
      perm.arr[r][perm.morphRoll2_C] = newMorphRoll2;
    }
  }
} // End fCSCalcNewPermData

// fCSCalcNewSkAndRollMorph //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Calculates newMorphStr, newSk and newRollMorph
function fCSCalcNewSkAndRollMorph(perm, r, sk1or2) {
  const isSk1 = sk1or2 === "Sk1";

  perm._morphCol = isSk1 ? perm.morph1_C : perm.morph2_C;
  perm._oldSkCol = isSk1 ? perm.oldSk1_C : perm.oldSk2_C;
  perm._newSkCol = isSk1 ? perm.newSk1_C : perm.newSk2_C;
  perm._morphRollCol = isSk1 ? perm.morphRoll1_C : perm.morphRoll2_C;

  // If there is no oldSk (which is MyAbil.trained)
  if (!perm.arr[r][perm._oldSkCol]) return ["", "", ""];

  const oldMorphStr = perm.arr[r][perm._morphCol];
  const heading = isSk1 ? "Sk1 Morph" : "Sk2 Morph";

  let newMorphStr = perm.arr[r][perm._morphCol]
    ? fCSCleanPermMorph(oldMorphStr, heading)
    : "";
  let newSk = perm.arr[r][perm._oldSkCol];
  let newRollMorph = "";

  if (newMorphStr)
    [newSk, newRollMorph] = fCSCalcNewSkNewMorphRoll(perm, r, newMorphStr);

  return [newMorphStr, newSk, newRollMorph];
} // End fCSCalcNewSkAndRollMorph

// fCSCleanPermMorph //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> throws errors if morphStr contains bad elements
function fCSCleanPermMorph(morphStr, heading) {
  if (!morphStr) return "";

  // Remove spaces, tabs, returns, reduce multiple commas to one, and remove leading/trailing commas
  let newMorphStr = morphStr
    .replace(/\s+/g, "") // Remove all whitespace characters
    .replace(/,+/g, ",") // Reduce multiple commas to a single comma
    .replace(/^,|,$/g, ""); // Remove leading and trailing commas

  const morphArr = newMorphStr.toUpperCase().split(",");

  let num; // Declare num here

  // For each element of morphArr verify that it matches one of the cases else throw an error
  for (const element of morphArr) {
    switch (true) {
      // Test just a Combine # (If there is just X where X is a positive number (e.g. 5, 8, 12.34))
      case /^\d+(\.\d+)?$/.test(element.toString()):
        break;

      // Test two letter + Integer Morphs
      case element.startsWith("=="):
      case element.startsWith("++"):
      case element.startsWith("--"):
        num = Number(element.substring(2));
        if (!Number.isInteger(num) || num < 1)
          throw new Error(
            `(1) ${heading} has an invalid argument: "${element}"`
          );
        break;

      // Test two letter + Positive Number Morphs
      case element.startsWith("**"):
      case element.startsWith("//"):
        num = Number(element.substring(2));
        if (isNaN(num) || num <= 0)
          throw new Error(
            `(2) ${heading} has an invalid argument: "${element}"`
          );
        break;

      // Test +1d, -1d, +1c, -1c, +1t, -1t (where 1 could be any integer including two or more digits)
      case /^[\+\-][1-9]\d*[DCT]$/.test(element):
        num = Number(element.substring(1, element.length - 1));
        if (!Number.isInteger(num) || num < 1)
          throw new Error(
            `(3) ${heading} has an invalid argument: "${element}"`
          );
        break;

      // Test one letter + Integer Morphs
      case element.startsWith("="):
      case element.startsWith("+"):
      case element.startsWith("-"):
        num = Number(element.substring(1));
        if (!Number.isInteger(num) || num < 1)
          throw new Error(
            `(4) ${heading} has an invalid argument: "${element}"`
          );
        break;

      // Test one letter + Positive Number Morphs
      case element.startsWith("*"):
      case element.startsWith("/"):
        num = Number(element.substring(1));
        if (isNaN(num) || num <= 0)
          throw new Error(
            `(5) ${heading} has an invalid argument: "${element}"`
          );
        break;

      default:
        throw new Error(`(6) ${heading} has an invalid argument: "${element}"`);
    }
  }

  return `,${newMorphStr}`;
} // End fCSCleanPermMorph

// fCSCalcNewSkNewMorphRoll //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Returns [newSk,newRollMorph]
// Purpose: Calculates NewSk from oldSk and MorphStr
// Purpose: Calculates newRollMorph based on "est Two letter + Interger Morphs"
function fCSCalcNewSkNewMorphRoll(perm, r, morphStr) {
  const oldsk = perm.arr[r][perm._oldSkCol];

  if (!morphStr) return [oldsk, ""];

  const morph = {
    base: oldsk,
    combineArr: [],
    plusMinus: 0,
    multDiv: 1,
    dctTotal: 0,

    // The keyname+2 format below is for ==, ++, --, //, **
    equals2: 0,
    plusMinus2: 0,
    multDiv2: 1,
  };

  const capMorphStr = morphStr.toUpperCase().replace(/^,/, ""); // All CAPS and remove leading ',' if it exists
  const morphArr = capMorphStr.split(",");

  morphArr.forEach((element) => {
    switch (true) {
      // Combine: If there is just X where X is a positive number (e.g. 5, 8, 12.34) then truncated to an integer and added to morph.combineArr
      case /^\d+(\.\d+)?$/.test(element.toString()): {
        morph.combineArr.push(Math.trunc(Number(element)));
        break;
      }

      // Test two letter + Integer Morphs
      case element.startsWith("=="):
        morph.equals2 = Math.max(morph.equals2, Number(element.substring(2)));
        break;
      case element.startsWith("++"):
        morph.plusMinus2 += Number(element.substring(2));
        break;
      case element.startsWith("--"):
        morph.plusMinus2 -= Number(element.substring(2));
        break;

      // Calc two letter + Positive Number Morphs
      case element.startsWith("**"):
        morph.multDiv2 *= Number(element.substring(2));
        break;
      case element.startsWith("//"):
        morph.multDiv2 /= Number(element.substring(2));
        break;

      // Calc +1d, -1d, +1c, -1c, +1t, -1t (where 1 could be any integer including two or more digits)
      case /^[\+\-][1-9]\d*[DCT]$/.test(element): {
        const isPositive = element.startsWith("+");
        let num = parseInt(element.substring(1));

        if (element.endsWith("T")) {
          num *= 10;
          morph.multDiv *= isPositive ? num : 1 / num;
        } else {
          // Else endsWith D or C
          num = element.endsWith("D") ? num : 3 * num;
          morph.dctTotal += isPositive ? num : -num;
        }
        break;
      }

      // Calc one letter + Integer Morphs
      case element.startsWith("="):
        morph.base = Number(element.substring(1));
        break;
      case element.startsWith("+"):
        morph.plusMinus += Number(element.substring(1));
        break;
      case element.startsWith("-"):
        morph.plusMinus -= Number(element.substring(1));
        break;
      case element.startsWith("*"):
        morph.multDiv *= Number(element.substring(1));
        break;
      case element.startsWith("/"):
        morph.multDiv /= Number(element.substring(1));
        break;

      default:
        throw new Error(
          `In fCSCalcNewSkNewMorphRoll, Morph not found: "${element}"`
        );
    }
  });

  // Calculate newSk
  let newSk = fCSCombineBaseAndArr(morph.base, morph.combineArr);
  newSk = Math.max(
    1,
    Math.round(
      fCSCalcdTiers(newSk, morph.dctTotal) * morph.multDiv + morph.plusMinus
    )
  );

  // Calculate newMorphRoll
  let newMorphRoll = "";
  newMorphRoll += morph.equals2 !== 0 ? `,==${morph.equals2}` : "";
  newMorphRoll += morph.plusMinus2 > 0 ? `,++${morph.plusMinus2}` : "";
  newMorphRoll += morph.plusMinus2 < 0 ? `,--${morph.plusMinus2}` : "";
  if (morph.multDiv2 !== 1) {
    const multiplier = parseFloat(morph.multDiv2.toFixed(3));
    newMorphRoll += `,**${multiplier}`;
  }

  return [newSk, newMorphRoll];
} // End fCSCalcNewSkNewMorphRoll

// fCSBuildconditionDropDown //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Builds Condition Drop Downs for <Game> and removes unused Conditions
function fCSBuildconditionDropDown(perm, game) {
  for (let r = game.abilTblStart_R; r <= game.gearTblEnd_R; r++) {
    // Skip rows between abil and gear tables
    if (r > game.abilTblEnd_R && r < game.gearTblStart_R) {
      r = game.gearTblStart_R - 1;
      continue;
    }

    const abilNameID = game.arr[r][game.abilNameID_C];
    const abilID = gTestID("mycs", "PermMorph", abilNameID);
    const gameCond = game.arr[r][game.condition_C];
    let gameCondFound = false;
    let dropDownList = [];

    // If abilID exists in <PermMorph>
    if (abilID) {
      let perm_R = gKeyR("mycs", "PermMorph", abilID);
      do {
        const permCond = perm.arr[perm_R][perm.condition_C];
        gameCondFound = gameCondFound || gameCond === permCond;
        dropDownList.push(permCond);
        perm_R--;
      } while (
        perm_R >= perm.dataFirst_R &&
        perm.arr[perm_R][perm.id_C] === abilID
      );
    }

    // Set the DropDown for <Game> "Condition" at the specified range
    const dropDownRange = game.ref.getRange(r + 1, game.condition_C + 1);
    if (dropDownList.length > 0) {
      // Create a data validation rule
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(dropDownList, true) // Only allow values from dropDownList
        .setAllowInvalid(false) // Prevent invalid data
        .build();
      dropDownRange.setDataValidation(rule);
    } else {
      // Clear any existing data validation from the cell.
      dropDownRange.clearDataValidations();
    }

    if (!gameCondFound) game.arr[r][game.condition_C] = "";
  }

  gSaveArraySectionToSheet(
    game.ref,
    game.arr,
    game.abilTblStart_R,
    game.gearTblEnd_R,
    game.condition_C,
    game.condition_C
  );
} // End fCSBuildconditionDropDown

// fCSApplyPermToGame //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Applies <PermMorph> to <Game>
function fCSApplyPermToGame(perm, game) {
  for (let r = game.abilTblStart_R; r <= game.gearTblEnd_R; r++) {
    // Skip rows between abil and gear tables
    if (r > game.abilTblEnd_R && r < game.gearTblStart_R) {
      r = game.gearTblStart_R - 1;
      continue;
    }

    const abilNameID = game.arr[r][game.abilNameID_C];
    const abilID = gTestID("mycs", "PermMorph", abilNameID);

    if (abilID) {
      let perm_R = gKeyR("mycs", "PermMorph", abilID); // Finds the LAST match
      let foundMatch = false;
      const gameCond = game.arr[r][game.condition_C];
      do {
        if (perm.arr[perm_R][perm.condition_C] === gameCond) {
          foundMatch = true;
          break;
        }
        perm_R--;
      } while (
        !foundMatch &&
        perm_R >= perm.dataFirst_R &&
        perm.arr[perm_R][perm.id_C] === abilID
      );

      if (foundMatch) {
        game.arr[r][game.permMorph1_C] = perm.arr[perm_R][perm.morphRoll1_C];
        game.arr[r][game.sk1_C] = perm.arr[perm_R][perm.newSk1_C];
        game.arr[r][game.permMorph2_C] = perm.arr[perm_R][perm.morphRoll2_C];
        game.arr[r][game.sk2_C] = perm.arr[perm_R][perm.newSk2_C];
      }
    }
  }

  gSaveArraySectionToSheet(
    game.ref,
    game.arr,
    game.abilTblStart_R,
    game.gearTblEnd_R,
    game.abilTableFirst_C,
    game.abilTableLast_C
  );
} // End fCSApplyPermToGame

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  Systems (end g. Gear)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// fCSSaveCSID //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Save the character sheet's URL and ID to <Data>
function fCSSaveCSID() {
  const csRef = gSSRef("mycs");
  const csID = csRef.getId();

  gSaveVal("mycs", "data", "myCSID", "Manual", csID);
} // End fCSSaveCSID

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  Die Rolls (end system)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// fCSd //////////////////////////////////////////////////////////////////////////////////////////////////
function fCSd(die) {
  die = Number(die);
  return Math.floor(Math.random() * die) + 1;
} // End fCSd

// fCSdBetween //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose ->  dBetween(6,8) = 6 or 7 or 8
function fCSdBetween(low, high) {
  low = Number(low);
  high = Number(high);
  return fCSd(high - low + 1) + (low - 1); // dBetween(6,8) = 6 or 7 or 8
} // End fCSdBetweenfCSDDmg

// fCSD //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Make an old fassioned, but smooth, MetaScape d16 roll (no T or C)
function fCSD(die) {
  die = Number(die);
  return Math.floor((die * Math.random() * 1) / Math.random()) + 1;
} // End fCSD

// fCSDSk //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Returns the average of a d(sk*2) and an old-fashioned, but smooth, MetaScape d16 roll (no T or C)
function fCSDSk(die) {
  die = Number(die);
  return Math.round((fCSd(die * 2) + fCSD(die)) / 2);
} // End fCSDSk

// fCSDDmg //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Rolls PC Dmg, which is a d(regular fDSk(die))
function fCSDDmg(pcDmg) {
  return fCSDSk(pcDmg);
} // End fCSDDmg

// fCSDArmor //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Rolls PC Armor which is a fDSk(armor) roll divided by 15 rounded to the 10ths place such as 1.2 with a minimum AR of .5
function fCSDArmor(pcAR) {
  return Math.max(1, Math.trunc((fCSDSk(pcAR) / 15) * 10) / 10);
} // End fCSDArmor

// fCSCalcdTiers //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Calculates applying +/- X number of tiers to dieBase such as +2d or -3d
function fCSCalcdTiers(dieBase, numOfd) {
  if (numOfd === 0) return dieBase;

  // Check if numOfd is an integer
  if (!Number.isInteger(numOfd))
    throw new Error(`In fCSdTiers, +/- # dtc has illegal # of: "${numOfd}"`);

  let newBase = dieBase;
  const isPositive = numOfd > 0;

  // Loop for the absolute value of numOfd

  const onedVal = 1 + 1 / 3; // 1.33333 is the +1d increment
  const difOneDVal = 1 / onedVal; // is the -1d increment

  for (let i = 0; i < Math.abs(numOfd); i++) {
    newBase =
      newBase === 1 && isPositive
        ? 2
        : newBase === 2 && !isPositive
        ? 1
        : newBase * (isPositive ? onedVal : difOneDVal);
    newBase = Math.max(1, Math.round(newBase));
  }

  return newBase;
} // End fCSCalcdTiers

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  Roll Log  (end g. Roll Helpers)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// fCSPostMyRoll //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Posts a string to the <Game> Roll Result cell, then prepends RollResult to Roll Log, and finally posts the Roll Log to <GMScreen> and to PS <PartyLog>
function fCSPostMyRoll(rollString) {
  let game = getObjCSGame();
  let psPartyLog = getObjPSPartyLog();
  let gmScreen = getObjDBGMScreen();

  // Get rollLog and trim it if rollLog.html is too long
  let rollLog = gGetDocCache("rollLog") || fCSCreateObjRollLog();
  if (rollLog.html.length > 50000) {
    rollLog.html = `${g.bluBS}Roll Log Cleared${g.endS}<br>Due to length.`;
    rollLog.plain = fCSStripHTMLFrom(rollLog.html);
  }

  // Prepend the new rollString and trim to a max of 10,000 characters
  rollLog.html = `${rollString}<br>${rollLog.html}`;
  rollLog.plain = `${fCSStripHTMLFrom(rollString)}\n${rollLog.plain}`;

  const loadAllTF = rollLog.stats.slot === "";
  fCSLoadRollLogStats(rollLog, game, loadAllTF);

  gPutDocCache("rollLog", rollLog);

  // Update GM Screen and Party Log with shared data
  const gmScreenLog_C = gHeaderC("db", "GMScreen", rollLog.stats.slot);
  const psPartyLog_C = gHeaderC("ps", "PartyLog", rollLog.stats.slot);

  // Array of rollLog.stats properties (without .slot) and adding in rollLog.plain as the last statsArray element
  const statsArray = [
    rollLog.stats.url,
    rollLog.stats.rc,
    rollLog.stats.level,
    rollLog.stats.vit,
    rollLog.stats.nish,
    rollLog.stats.playerCharNames,
    rollLog.plain,
  ];

  // Save to gmScreen.arr and psPartyLog.arr
  for (let i = 0; i < 7; i++) {
    gmScreen.arr[gmScreen.url_R + i][gmScreenLog_C] = statsArray[i];
    psPartyLog.arr[psPartyLog.url_R + i][psPartyLog_C] = statsArray[i];
  }

  // Save to gmScreen and psPartyLog
  gSaveArraySectionToSheet(
    psPartyLog.ref,
    psPartyLog.arr,
    psPartyLog.url_R,
    psPartyLog.url_R + 6,
    psPartyLog_C,
    psPartyLog_C
  );
  gSaveArraySectionToSheet(
    gmScreen.ref,
    gmScreen.arr,
    gmScreen.url_R,
    gmScreen.url_R + 6,
    gmScreenLog_C,
    gmScreenLog_C
  );
} // End fCSPostMyRoll

// fCSCreateObjRollLog //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Factory Function for CacheServices in fCSPostMyRoll
function fCSCreateObjRollLog() {
  return {
    html: "",
    plain: "",
    stats: {
      slot: "",
      url: "",
      rc: "",
      level: 1,
      vit: "",
      nish: 0,
      playerCharNames: "",
    },
  };
} // End fCSCreateObjRollLog

// fCSStripHTMLFrom //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose:
// - Removes <span any text> and </span> tags, leaving just the contents
// - Converts <br> to \n
// - Converts &amp; to '&' and &gt; or &gt to '>'
function fCSStripHTMLFrom(htmlString) {
  let plainTxt = htmlString
    .replace(/<\/?span.*?>/gi, "") // Removes all <span> and </span> tags
    .replace(/<br\s*\/?>/gi, "\n") // Converts <br> (and <br />) to \n
    .replace(/&amp;/g, "&") // Converts &amp; to '&'
    .replace(/&gt;?/g, ">"); // Converts &gt; or &gt to '>'
  return plainTxt;
} // End fCSStripHTMLFrom

// fCSLoadRollLogStats //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Populates the rollLog.stats object based on the provided RaceClass and Game data
function fCSLoadRollLogStats(rollLog, game, loadAllTF) {
  // Populate rollLog.stats with relatively static information
  if (loadAllTF) {
    let raceClass = getObjCSRaceClass();

    // Ensure rollLog is an object
    if (!rollLog || typeof rollLog !== "object" || !rollLog.stats)
      throw new Error("Invalid rollLog object passed to fCSLoadRollLogStats.");

    // Save Roll Log to <GMScreen>
    let mySlotNum = raceClass.arr[raceClass.slot_R][raceClass.value_C]
      .trim()
      .slice(-1);
    if (!/^[1-9]$/.test(mySlotNum))
      throw new Error(
        `Your <RaceClass> "GM Roll Slot" is not valid. It should look like "Slot #" where # is 1 to 9.`
      );

    const mySlot = `Slot${mySlotNum}`;
    const sheetURL = gURLFromID(g.id.mycs, "sheet");
    const rc = raceClass.arr[raceClass.rc_R][raceClass.value_C];
    const level = raceClass.arr[raceClass.charLevel_R][raceClass.value_C];
    const player = raceClass.arr[raceClass.playerName_R][raceClass.value_C];
    const char = raceClass.arr[raceClass.charName_R][raceClass.value_C];
    const playerChar = `${player} (${char})`;

    rollLog.stats.slot = mySlot;
    rollLog.stats.url = sheetURL;
    rollLog.stats.level = level;
    rollLog.stats.rc = rc;
    rollLog.stats.playerCharNames = playerChar;
  }

  // Populate rollLog.stats with highly voliltile information
  const vitNow = game.arr[game.vitTbl_R][game.vitNow_C];
  const vitPercent = game.arr[game.vitTbl_R][game.vitNowPercent_C];
  const vitMax = game.arr[game.vitTbl_R][game.vitMax_C];
  const vitString = `${vitNow} of ${vitMax} (${vitPercent})`;
  const nish = game.arr[game.nish_R][game.nish_C];

  rollLog.stats.vit = vitString;
  rollLog.stats.nish = nish;
} // End fCSLoadRollLogStats
