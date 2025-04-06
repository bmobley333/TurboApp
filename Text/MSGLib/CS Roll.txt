////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  Master Roller ()
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// fCSMasterRoller //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Determines the type of roll (Sk1, Sk2, RPG, or MSO) and rolls. Works across all CS sheets
// Purpose -> Prioritize speed of normal skill roll as they are most critical, most common, and most complex
function fCSMasterRoller(luckedOrFreeStr = "") {
  let game = getObjCSGame();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getSheetName();

  const activeCellRef = sheet.getActiveCell();

  const activeCell = {
    val: String(activeCellRef.getValue()).trim(),
    r: activeCellRef.getRow() - 1, // Adjust row to zero-based index
    c: activeCellRef.getColumn() - 1, // Adjust column to zero-based index
  };

  // Set just a roll.log up for basic recoding. roll object is greatly expanded if the roll is an Ability
  let roll = { log: "" };

  // Determine if the roll is a <Game> Sk1/Sk2 roll
  let isSk1,
    isSk2 = false;
  [isSk1, isSk2] = fIsSk1OrSk2ChkBox(game, activeCell);
  if (isSk1 || isSk2) {
    roll = fCSFFRoll(game, activeCell.r, isSk1, luckedOrFreeStr);

    fCSCalcMorphEffects(game, roll, activeCell.r);
    roll.isCombat = ["atk", "def", "dmg", "ar"].includes(
      roll.type.toLowerCase()
    );
    fCSCalcRollCosts(game, roll, activeCell.r); // Meta, Action, Uses
    fCSSelectAbilRollFunc(game, roll, activeCell.r);

    // Empty Cell
  } else if (!activeCell.val) {
    throw new Error(`Empty Cell. Nothing to roll.`);

    // Astral Gauntlet
  } else if (
    sheetName === "Game" &&
    activeCell.r >= game.chaos1_R &&
    activeCell.r <= game.chaosWrist_R &&
    activeCell.c >= game.chaosUses_C &&
    activeCell.c <= game.chaosAbility_C
  ) {
    fCSPerformUseChaosGem(game, roll, activeCell.r);

    // If just a number, make an MSO roll
  } else if (!isNaN(activeCell.val)) {
    fCSRollSk(roll, activeCell.val);

    // Else Try an RPG roll
  } else {
    fCSRollRPG(roll, activeCell.val);
  }

  fCSSaveAbilRollEffects(game, roll);
  fCSRestoreCellFormula(game);
} // End fCSMasterRoller

// fIsSk1OrSk2ChkBox //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Determines if the roll is for Sk1 or Sk2 and if either, changes the values of the activeCell object, and unchecks the Sk#ChkBox.
// Purpose -> Ensures only one checkbox is checked; otherwise, calls fThereMayBeOnlyOneChkBox(game).
// Purpose -> If sk1 was checked and there is a sk2, be sure to check Sk2 (e.g., Atk then check Dmg, Def the check AR, etc.)
function fIsSk1OrSk2ChkBox(game, activeCell) {
  let sk1Checked = 0;
  let sk2Checked = 0;
  let sk1Row = -1;
  let sk2Row = -1;

  // Loop through rows of abilTbl and gearTbl
  for (let r = game.abilTblStart_R; r <= game.gearTblEnd_R; r++) {
    // Skip rows between abil and gear tables
    if (r > game.abilTblEnd_R && r < game.gearTblStart_R) {
      r = game.gearTblStart_R - 1;
      continue;
    }

    if (game.arr[r][game.sk1ChkBox_C]) {
      sk1Checked++;
      sk1Row = r;
    }
    if (game.arr[r][game.sk2ChkBox_C]) {
      sk2Checked++;
      sk2Row = r;
    }
  }

  // Check if more than one checkbox is checke
  if (sk1Checked + sk2Checked > 1) {
    fCSClearSk1Sk2CheckBoxes(game);
    throw new Error(
      "Multiple Skills checked.\n\nAll skill checkboxes will be cleared. Please try again."
    );
  }

  // If no checkbox is checked, return [false, false]
  if (sk1Checked + sk2Checked === 0) return [false, false];

  // Determine if Sk1 or Sk2 is checked
  const isSk1 = sk1Checked === 1;
  const isSk2 = sk2Checked === 1;

  // Set activeCell object to the proper Sk#ChkBox
  activeCell.r = isSk1 ? sk1Row : sk2Row;
  activeCell.c = isSk1 ? game.sk1ChkBox_C : game.sk2ChkBox_C;
  activeCell.val = true;

  // Uncheck Sk#ChkBox (and if isSk1 also check sk2ChkBox_C if Sk2_C is non blank)
  if (isSk1) {
    game.arr[sk1Row][game.sk1ChkBox_C] = false;
    game.arr[sk1Row][game.sk2ChkBox_C] = game.arr[sk1Row][game.sk2_C] !== "";
  } else {
    game.arr[sk2Row][game.sk2ChkBox_C] = false;
  }

  return [isSk1, isSk2];
} // End fIsSk1OrSk2ChkBox

// fCSFFRoll //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Creates the roll object for <Game> sheet
function fCSFFRoll(game, r, isSk1, luckedOrFreeStr) {
  const newObj = {
    isSk1: isSk1,
    isCombat: false,
    isFreeRoll: luckedOrFreeStr === "Free",
    isLuckedRoll: luckedOrFreeStr === "Lucked",
    unSkilled: false,
    finalTC: "", // stores any 'T' or 'C' (tremendous or crit)
    wasMetaSpent: false,
    wasActionUsed: false,
    wasUsesUsed: false,
    name: `${g.bluBS}${gSimpleName(game.arr[r][game.abilNameID_C])}${g.endS}`,
    id: gGetIDFromString(game.arr[r][game.abilNameID_C]),
    difficulty: 0,
    focusColor: "",
    dctTotal: 0,
    morphString: "",
    typeOriginal: isSk1
      ? game.arr[r][game.sk1Typ_C]
      : game.arr[r][game.sk2Typ_C],
    type: isSk1 ? game.arr[r][game.sk1Typ_C] : game.arr[r][game.sk2Typ_C],
    base: isSk1 ? game.arr[r][game.sk1_C] : game.arr[r][game.sk2_C],
    morphedBase: 0,
    equals: 0,
    postEquals: 0,
    result: 0,
    morphedResult: 0,
    focusPlus: 0,
    plus: 0,
    mult: 1,
    postPlus: 0,
    postMult: 1,
    combineArr: [], // For all combines
    numOfRolls: 1,
    monWndsAccumultator: 0,
    pcWndsAccumulator: 0,
    totalMonWnds: 0,
    totalPCWnds: 0,
    mon: {},
    log: "",
  };

  return newObj;
} // End fCSFFRoll

// fCSCalcMorphEffects //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> cleans and applies permMorp and Morph to roll
function fCSCalcMorphEffects(game, roll, r) {
  const crudeMorphString = roll.isSk1
    ? `${game.arr[r][game.permMorph1_C]},${game.arr[r][game.morph1_C]}`
    : `${game.arr[r][game.permMorph2_C]},${game.arr[r][game.morph2_C]}`;

  const cleanMorphString = String(crudeMorphString)
    .replace(/\s+/g, "") // Remove all whitespace
    .replace(/,+/g, ",") // Remove any ',,' or ',,,' etc. in morphString
    .replace(/^,|,$/g, "") // Remove any leading or trailing ','
    .toUpperCase(); // Upper Case

  if (cleanMorphString) {
    roll.morphString = cleanMorphString;
    const morphArr = cleanMorphString.split(",");
    fCSApplyMorph(roll, morphArr);
  }
} // End fCSCalcMorphEffects

// fCSApplyMorph //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Verifies and Applies morphArr. Note all text in moprhArr is .toUpperCase via fCSCalcMorphEffects
function fCSApplyMorph(roll, morphArr) {
  morphArr.forEach((element) => {
    switch (true) {
      // Meta Focus Color (if more than one, take best)
      case ["~", "R", "O", "Y", "G", "B"].includes(element): {
        const colors = ["", "~", "R", "O", "Y", "G", "B"];
        const maxIndex = Math.max(
          colors.indexOf(roll.focusColor),
          colors.indexOf(element)
        );
        roll.focusColor = colors[maxIndex];
        break;
      }

      // Calc roll types
      case element === "SKI":
      case element === "SK":
        roll.type = "Sk";
        break;
      case element === "UN":
      case element === "UNSK":
        roll.unSkilled = true;
        break;
      case element === "ATK":
        roll.type = "Atk";
        break;
      case element === "DMG":
        roll.type = "DMG";
        break;
      case element === "DEF":
        roll.type = "Def";
        break;
      case element === "ARM":
      case element === "AR":
        roll.type = "AR";
        break;
      case element === "FREE":
        roll.isFreeRoll = true;
        break;
      case element === "LUCK":
        roll.isLuckedRoll = true;
        break;

      // Calc two letter + Positive Number
      case element.startsWith("**"):
        roll.postMult *= Number(element.substring(2));
        break;
      case element.startsWith("//"):
        roll.postMult /= Number(element.substring(2));
        break;
      case element.startsWith("=="):
        roll.postEquals = Math.max(
          roll.postEquals,
          Math.abs(Number(element.substring(2)))
        );
        break;
      case element.startsWith("++"):
        roll.postPlus += Number(element.substring(2));
        break;
      case element.startsWith("--"):
        roll.postPlus -= Number(element.substring(2));
        break;

      // Calc +1d, -1d, +1c, -1c, +1t, -1t (where 1 could be any integer including two or more digits)
      case /^[\+\-][1-9]\d*[DCT]$/.test(element): {
        const isPositive = element.startsWith("+");
        let num = parseInt(element.substring(1));
        if (element.endsWith("T")) {
          num *= 10;
          roll.mult *= isPositive ? num : 1 / num;
        } else {
          // Else endsWith D or C
          num = element.endsWith("D") ? num : 3 * num;
          roll.dctTotal += isPositive ? num : -num;
        }
        break;
      }

      // If there is a #X where X is a positive number (e.g. 1, 2, 27.89, etc. with 50 as Max) and X is truncated to an integer
      case /^#\d+(\.\d+)?$/.test(element.toString()): {
        const truncatedValue = Math.trunc(
          Number(element.toString().replace(/^#/, ""))
        );
        roll.numOfRolls = Math.min(
          50,
          Math.max(roll.numOfRolls, truncatedValue)
        );
        break;
      }

      // Combine: If there is just X where X is a positive number (e.g. 5, 8, 12.34) then truncated to an integer and added to roll.combineArr
      case /^\d+(\.\d+)?$/.test(element.toString()): {
        roll.combineArr.push(Math.trunc(Number(element)));
        break;
      }

      // Calc one letter + Integer Morphs
      case element.startsWith("="):
        roll.equals = Math.max(
          roll.equals,
          Math.abs(Number(element.substring(1)))
        );
        break;
      case element.startsWith("+"):
        roll.plus += Number(element.substring(1));
        break;
      case element.startsWith("-"):
        roll.plus -= Number(element.substring(1));
        break;
      case element.startsWith("^"):
        roll.difficulty = Math.max(
          roll.difficulty,
          Number(element.substring(1))
        );
        break;
      case element.startsWith("*"):
        roll.mult *= Number(element.substring(1));
        break;
      case element.startsWith("/"):
        roll.mult /= Number(element.substring(1));
        break;

      default:
        throw new Error(`I don't understand this morph: "${element}"`);
    }
  });

  // Calculate +2d, -1c, etc. effects onto roll.mult
  fCSdTiers(roll);

  // Calculate Meta Focus and add to roll.plus
  fCSCalcFocus(roll);
} // End fCSApplyMorph

// fCSdTiers //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Calculates applying +/- X number as a multiplier for +2d or -3d etc. and adjusts roll.mult accordingly
function fCSdTiers(roll) {
  if (roll.dctTotal === 0) return;

  // Check if dctTotal is an integer
  if (!Number.isInteger(roll.dctTotal))
    throw new Error(
      `In fCSdTiers, +/- # dct has illegal # of: "${roll.dctTotal}"`
    );

  let newMult = 1;
  const isPositive = roll.dctTotal > 0;

  // Loop for the absolute value of dctTotal
  const onedVal = 1 + 1 / 3; // 1.33333 is the +1d increment
  const divOneDVal = 1 / onedVal; // is the -1d increment

  for (let i = 0; i < Math.abs(roll.dctTotal); i++) {
    newMult *= isPositive ? onedVal : divOneDVal;
  }

  roll.mult *= newMult;
} // End fCSdTiers

// fCSCombineBaseAndArr //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Combines an oldBase value with combineArr and returns finalBase
function fCSCombineBaseAndArr(oldBase, combineArr) {
  // If combineArr is an empty array return oldBase
  if (combineArr.length === 0) return oldBase;

  // Add oldBase to the list in combineArr and sort from largest to smallest value
  combineArr.push(oldBase);
  combineArr.sort((a, b) => b - a);

  // Calculate finalAR by adding each element divided by increasing powers of 2
  let finalBase = 0;
  for (let i = 0; i < combineArr.length; i++) {
    finalBase += combineArr[i] / 2 ** i;
  }

  return Math.round(finalBase);
} // End fCSCombineBaseAndArr

// fCSCalcFocus //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Calculate Meta Focus and add to roll.focusPlus
function fCSCalcFocus(roll) {
  // If no focus
  if (!roll.focusColor) return;

  const metaColorArr = ["R", "O", "Y", "G", "B"];
  const metaBonus = [5, 10, 15, 20, 25];

  const focusColorUpper = roll.focusColor.toUpperCase();
  if (!metaColorArr.includes(focusColorUpper))
    throw new Error(
      `In fCSCalcFocus, there is an illegal focus color of "${roll.focusColor}"`
    );

  const i = metaColorArr.indexOf(focusColorUpper);
  const finalPlus = roll.unSkilled ? (metaBonus[i] * 2) / 5 : metaBonus[i];

  roll.focusPlus = finalPlus;
} // End fCSCalcFocus

// fCSCalcRollCosts //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Calculates the cost of the roll: Meta, Action, Uses
// Purpose -> NOTE: if roll isFreeRoll or isLuckedRoll there is no cost of any kind
function fCSCalcRollCosts(game, roll, r) {
  if (roll.isFreeRoll) return;
  if (roll.isLuckedRoll) {
    fCSSpendLuckBox(game);
    return;
  }

  const metaColor = roll.isSk1 ? game.arr[r][game.meta_C] : "";

  fCSSpendMetaColor(game, roll, metaColor);
  fCSSpendMetaColor(game, roll, roll.focusColor);

  if (roll.isSk1) {
    fCSSpendAction(game, roll, r);
    fCSSpendUses(game, roll, r);
    fCSSetOn(game, r);
  }
} // End fCSCalcRollCosts

// fCSSpendLuckBox //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Spends one luck box
function fCSSpendLuckBox(game) {
  const luckNum = game.arr[game.luckBox_R][game.luckBox_C];
  if (!Number.isInteger(luckNum) || luckNum <= 0)
    throw new Error(`At "${luckNum}" Luck Boxes, you are out of Luck.`);

  game.arr[game.luckBox_R][game.luckBox_C]--;
} // End fCSSpendLuckBox

// fCSSpendMetaColor //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Spends the passed metaColor, if any
function fCSSpendMetaColor(game, roll, metaColor) {
  const metaColorArr = ["R", "O", "Y", "G", "B"];

  if (metaColor === "~" || metaColor === "") return;
  if (!metaColorArr.includes(metaColor.toUpperCase())) {
    SpreadsheetApp.getUi().alert(
      `In function fCSSpendMetaColor a non meta color of "${metaColor}" was passed.`
    );
    return;
  }

  let metaWasSpent = false;
  let metaSpend_C =
    game.metaR_C + metaColorArr.indexOf(metaColor.toUpperCase());

  while (metaSpend_C <= game.metaB_C) {
    if (game.arr[game.metaChnl_R][metaSpend_C] > 0) {
      const newMetaNum = game.arr[game.metaChnl_R][metaSpend_C] - 1;
      game.arr[game.metaChnl_R][metaSpend_C] =
        newMetaNum === 0 ? "" : newMetaNum;
      metaWasSpent = true;
      break;
    }
    if (game.arr[game.meta_R][metaSpend_C] > 0) {
      const newMetaNum = game.arr[game.meta_R][metaSpend_C] - 1;
      game.arr[game.meta_R][metaSpend_C] = newMetaNum === 0 ? "" : newMetaNum;
      metaWasSpent = true;
      break;
    }
    metaSpend_C++;
  }

  if (metaWasSpent) {
    roll.wasMetaSpent = true;
  } else {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      `You do not have the needed @${metaColor}.\n\nContinue Anyway?`,
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES)
      throw new Error("Roll canceled due to lack of Meta.");
    roll.log += `<br>-------- ${g.redBS}ALERT${g.endS} --------<br>Above Rolled without @${metaColor}<br>`;
  }
} // End fCSSpendMetaColor

// fCSSpendAction //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Spends the passed action cost
function fCSSpendAction(game, roll, r) {
  const actCost = game.arr[r][game.act_C];

  // Check if actCost is one of the valid values
  if (![1, 2, 3, 4, 5].includes(actCost)) return;

  let actionsLeft = game.arr[game.actTotal_R][game.actTotal_C];
  let actWasSpent = false;

  if (actionsLeft >= actCost) {
    actionsLeft -= actCost;
    actWasSpent = true;
  }

  if (actWasSpent) {
    roll.wasActionUsed = true;
    game.arr[game.actTotal_R][game.actTotal_C] = actionsLeft; // Update the game array
  } else {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      `You do not have the needed Act(${actCost}).<br>\nContinue Anyway?`,
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES)
      throw new Error("Roll canceled due to lack of Action.");
    roll.log += `<br>-------- ${g.redBS}ALERT${g.endS} --------<br>Above Rolled without Act(${actCost})<br>`;
    game.arr[game.actTotal_R][game.actTotal_C] = 0; // Set actionsLeft to 0 in the game array
  }
} // End fCSSpendAction

// fCSSpendUses //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Decrements Uses
function fCSSpendUses(game, roll, r) {
  let currentUses = game.arr[r][game.uses_C];

  // Check if currentUses is not an integer or is less than 0
  if (!Number.isInteger(currentUses) || currentUses < 0) return;

  if (currentUses > 0) {
    game.arr[r][game.uses_C]--;
    roll.wasUsesUsed = true;
  } else {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      `Your Uses is at 0.\n\nContinue Anyway?`,
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES)
      throw new Error("Roll canceled due to lack of Uses.");
    roll.log += `<br>-------- ${g.redBS}ALERT${g.endS} --------<br>Above Rolled without enough Uses<br>`;
  }
} // End fCSSpendUses

// fCSSetOn //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Set On
function fCSSetOn(game, r) {
  const dur = game.arr[r][game.dur_C];
  const onNow = game.arr[r][game.on_C];

  if (!onNow || onNow === ".") game.arr[r][game.on_C] = dur;
} // End fCSSetOn

// fCSSelectAbilRollFunc //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Now that Morphs are calculated, selects the type of Ability Roll needed and calls the proper ability rolling function
function fCSSelectAbilRollFunc(game, roll, r) {
  roll.base = roll.base ? Math.round(Math.max(1, roll.base)) : 0;

  // If roll.base is empty and there is a roll.equals then cahnge it to any =# (e.g. =100)
  roll.base =
    !roll.base && roll.equals
      ? Math.round(Math.max(1, roll.equals))
      : roll.base;

  if (roll.base) {
    roll.morphedBase = roll.equals
      ? fCSCombineBaseAndArr(roll.equals, roll.combineArr)
      : fCSCombineBaseAndArr(roll.base, roll.combineArr);
    roll.morphedBase = Math.round(
      Math.max(1, roll.morphedBase * roll.mult + roll.plus)
    );

    const monRoster = roll.isCombat ? fCSFFGetMonsters(game, roll) : {};

    if (roll.isCombat) {
      if (monRoster && monRoster.hasMonsters) {
        fCSCalcCombatVsMonRoll(game, roll, monRoster);
      } else {
        throw new Error(
          `Please select a Monster Table opponent for your ${roll.name} ${roll.type} roll.`
        );
      }
    } else {
      fCSCalcNonCombatRoll(roll);
      if (roll.id === "k97cmz" && r === game.nishAtr_R)
        fCSDoNishExtras(game, roll);
    }
  } else {
    roll.log = `${roll.name}<br>Activated<br>${roll.log}`;
  }
} // End fCSSelectAbilRollFunc

// fCSFFGetMonsters //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Creates a monRoster object of all Monsters on the <Game> Monster table with a # > 0 and sorts by the critical attribute based on PC's Atk, Dmg, Def, AR
function fCSFFGetMonsters(game, roll) {
  const monStatsTemplate = {
    numOf: 0,
    name: "",
    atk: 0,
    dmg: 0,
    def: 0,
    ar: 1,
  };

  const monRoster = {
    hasMonsters: false,
    list: [],
  };

  for (let r = game.custMon_R; r <= game.monsterLast_R; r++) {
    const numOfMon = game.arr[r][game.numOfMonsters_C];
    if (numOfMon) {
      const newMonStats = { ...monStatsTemplate }; // create a new instance of the monStats object
      newMonStats.numOf = numOfMon;
      newMonStats.name =
        r === 1
          ? `Custom: ${game.arr[r][game.monsterName_C]}`
          : game.arr[r][game.monsterName_C] || "NoName";
      newMonStats.atk = isNaN(Number(game.arr[r][game.monsterAtk_C]))
        ? 0
        : Math.max(0, Number(game.arr[r][game.monsterAtk_C]));
      newMonStats.dmg = isNaN(Number(game.arr[r][game.monsterDmg_C]))
        ? 0
        : Math.max(0, Number(game.arr[r][game.monsterDmg_C]));
      newMonStats.def = isNaN(Number(game.arr[r][game.monsterDef_C]))
        ? 0
        : Math.max(0, Number(game.arr[r][game.monsterDef_C]));
      newMonStats.ar = isNaN(Number(game.arr[r][game.monsterAR_C]))
        ? 1
        : Math.max(1, Number(game.arr[r][game.monsterAR_C]));

      monRoster.list.push(newMonStats);
    }
  }

  monRoster.hasMonsters = monRoster.list.length > 0;

  if (!monRoster.hasMonsters) {
  }

  // Sort Monsters based on best attribute versus PC's roll.type so Focus will be versus the best target
  if (monRoster.hasMonsters) fCSSortMonsters(roll, monRoster);

  return monRoster;
} // End fCSFFGetMonsters

// fCSSortMonsters //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Sorts monRoster from largest to smallest based on best attribute versus PC's roll.type so Focus will be versus the best target
function fCSSortMonsters(roll, monRoster) {
  const pcSkType = roll.type.toUpperCase();

  let sortKey = "";

  switch (pcSkType) {
    case "ATK":
      sortKey = "def";
      break;
    case "DMG":
      sortKey = "ar";
      break;
    case "DEF":
      sortKey = "atk";
      break;
    case "AR":
      sortKey = "dmg";
      break;
    default:
      throw new Error(`In fCSSortMonsters, unknown skill type: ${pcSkType}`);
  }

  // Sort monRoster from largest to smallest value of monRoster.list[i][sortKey]
  monRoster.list.sort((a, b) => b[sortKey] - a[sortKey]);
} // End fCSSortMonsters

// fCSCalcCombatVsMonRoll //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Rolls a combat roll versus monsters:
// Assumes -> roll.base > 0, is a combat roll (atk, dmg, def, ar), monsters have been selected by player and stored in sorted order in monRoster.list
function fCSCalcCombatVsMonRoll(game, roll, monRoster) {
  let thisRollLog = "";

  const firstMonName = monRoster.list[0].name;

  for (const monster of monRoster.list) {
    const isFirstMonType = monster.name === firstMonName;

    // Roll in order so can total results if ever wanted
    for (let i = 1; i <= monster.numOf; i++) {
      const isFirstMon = i === 1 && isFirstMonType;

      // Monster Name
      thisRollLog +=
        i === 1 && monster.name ? `|<br>--- ${monster.name} ---<br>` : "";

      const rollCounter = monster.numOf === 1 ? "" : `${i}&gt `;

      // if isFirstMon apply Meta Focus to tempMophedBase
      const tempMorphedBase = isFirstMon
        ? roll.morphedBase + roll.focusPlus
        : roll.morphedBase;
      rollBaseStr =
        roll.base === tempMorphedBase
          ? `(${roll.base})`
          : `(${roll.base}-&gt${tempMorphedBase})`;

      // Based on roll.type and monster stats calculate: rollResultStr, difficultyStr
      const [rollResultStr, difficultyStr] =
        fCSCalcCombatVsMonResultStringAndDiffStrings(roll, monster);

      roll.finalTC = isFirstMon ? fCStc() : "";

      thisRollLog += `${rollCounter}${roll.unSkilled ? "UnSk" : ""}${
        roll.type
      }${rollBaseStr} ${rollResultStr}${difficultyStr} ${roll.finalTC}<br>`;
    }

    if (roll.pcWndsAccumulator) {
      thisRollLog += `${monster.name} Wounds to PC = ${g.redItS}-${roll.pcWndsAccumulator}${g.endS}<br>`;
      roll.totalPCWnds += roll.pcWndsAccumulator;
      roll.pcWndsAccumulator = 0;
    } else if (roll.monWndsAccumultator) {
      thisRollLog += `${monster.name} Wounds = ${g.grnBS}+${roll.monWndsAccumultator}${g.endS}<br>`;
      roll.totalMonWnds += roll.monWndsAccumultator;
      roll.monWndsAccumultator = 0;
    }
  }

  if (roll.totalPCWnds) {
    thisRollLog += `|<br>Total Wounds to PC = ${g.redItS}-${roll.totalPCWnds}${g.endS}<br>`;
    game.arr[game.vitTbl_R][game.vit1st_C] =
      parseFloat(game.arr[game.vitTbl_R][game.vit1st_C]) || 0;
    game.arr[game.vitTbl_R][game.vit1st_C] += roll.totalPCWnds;
    game.arr[game.vitTbl_R][game.vitNow_C] =
      parseFloat(game.arr[game.vitTbl_R][game.vitNow_C]) || 0;
    game.arr[game.vitTbl_R][game.vitNow_C] -= roll.totalPCWnds;
    thisRollLog += `Vit Now = ${game.arr[game.vitTbl_R][game.vitNow_C]}/${
      game.arr[game.vitTbl_R][game.vitMax_C]
    }<br>`;
    if (game.arr[game.vitTbl_R][game.vitNow_C] < 0)
      thisRollLog += `${g.redBS}WARNING: CRITICAL DMG${g.endS}<br>`;
  }

  const freeOrLucked =
    roll.isFreeRoll || roll.isLuckedRoll
      ? roll.isLuckedRoll
        ? `${g.redBS}LUCKED ROLL${g.endS}<br>`
        : `${g.redBS}FREE ROLL${g.endS}<br>`
      : "";
  const rollTypeString =
    roll.typeOriginal === roll.type
      ? roll.type
      : `${roll.typeOriginal} as ${roll.type}`;
  const morphString = roll.morphString ? `${roll.morphString}<br>` : "";
  roll.log = `${freeOrLucked}${roll.name} ${rollTypeString}<br>${morphString}${thisRollLog}${roll.log}`;
} // End fCSCalcCombatVsMonRoll

// fCSCalcCombatVsMonResultStringAndDiffStrings //////////////////////////////////////////////////////////////////////////////////////////////////
// Helper of fCSCalcCombatVsMonRoll
// Purpose -> for a combat roll with monsters, calculates [rollResultStr, difficultyStr]
// Full thisRollLog: `${rollCounter}${roll.type}${rollBaseStr} ~${rollResultStr}${difficultyStr} ${roll.finalTC}<br>`
// Example rollResultStr: ~${rollResultStr} looks like: ~16, ~16->18, ~16/2
// Example difficultyStr: `^${roll.difficulty} = +${difResult}` looks like ^12 = +3, /2 = +9, ^12 = +3->+8, /2 = +9->+18
function fCSCalcCombatVsMonResultStringAndDiffStrings(roll, monster) {
  let rollResultStr = "";
  let difficultyStr = "";

  switch (roll.type.toUpperCase()) {
    case "ATK":
      const pcAtk = (roll.result = roll.postEquals
        ? Math.round(Math.max(1, roll.postEquals))
        : roll.unSkilled
        ? fCSRollUnSkilled(roll.base)
        : fCSDSk(roll.base));
      const pcPostAtk = (roll.morphedResult =
        pcAtk * roll.postMult + roll.postPlus);
      rollResultStr =
        pcAtk === pcPostAtk ? `~${pcAtk}` : `~${pcAtk}-&gt${pcPostAtk}`;

      // Get the correct monster stat
      const monDef = (roll.difficulty = monster.def);
      const pcAtkResult = pcPostAtk - monDef;
      difficultyStr =
        pcAtkResult >= 0
          ? `^${monDef} = ${g.grnBS}+${pcAtkResult}${g.endS}`
          : `^${monDef} = ${g.redItS}${pcAtkResult}${g.endS}`; // Adds the '+' for positive results
      break;

    case "DMG":
      // roll custom fCSDDmg
      const pcDmg = (roll.result = roll.postEquals
        ? Math.round(Math.max(1, roll.postEquals))
        : roll.unSkilled
        ? fCSRollUnSkilledDmg(roll.base)
        : fCSDDmg(roll.base));

      // Don't apply post roll adjustments (roll.postMult + roll.postPlus) until after /AR so... there is no roll.morphedResult
      rollResultStr = `~${pcDmg}`;

      // Get the correct monster stat
      const monArmor = !monster.ar
        ? 1
        : Math.max(1, Math.round((monster.ar / 10) * 10) / 10);
      const monWounds = Math.round(pcDmg / monArmor);
      const monPostWounds = monWounds
        ? monWounds * roll.postMult + roll.postPlus
        : 0;
      const monWndSlideStr =
        monWounds === monPostWounds
          ? monPostWounds >= 0
            ? `${g.grnBS}+${monPostWounds}${g.endS}`
            : `${g.redItS}${monPostWounds}${g.endS}`
          : monPostWounds >= 0
          ? `${monWounds}-&gt+${g.grnBS}${monPostWounds}${g.endS}`
          : `${monWounds}-&gt${g.redItS}${monPostWounds}${g.endS}`;
      roll.monWndsAccumultator += monPostWounds;
      difficultyStr = `/${monArmor} = ${monWndSlideStr}`;
      break;

    case "DEF":
      const pcDef = (roll.result = roll.postEquals
        ? Math.round(Math.max(1, roll.postEquals))
        : roll.unSkilled
        ? fCSRollUnSkilled(roll.base)
        : fCSDef(roll.base));
      const pcPostDef = (roll.morphedResult =
        pcDef * roll.postMult + roll.postPlus);
      rollResultStr =
        pcDef === pcPostDef ? `~${pcDef}` : `~${pcDef}-&gt${pcPostDef}`;

      // Get the correct monster stat
      const monAtk = (roll.difficulty = monster.atk);
      const pcDefResult = pcPostDef - monAtk;
      difficultyStr =
        pcDefResult >= 0
          ? `^${monAtk} = ${g.grnBS}+${pcDefResult}${g.endS}`
          : `^${monAtk} = ${g.redItS}${pcDefResult}${g.endS}`; // Adds the '+' for positive results
      break;

    case "AR":
      // Get the correct monster stat
      const monDmg = monster.dmg;
      rollResultStr = `${monDmg}`;

      // roll custom fCSDArmor rounded to 10th's place
      const pcAR = (roll.result = roll.postEquals
        ? Math.round(Math.max(1, roll.postEquals))
        : roll.unSkilled
        ? fCSRollUnSkilledArmor(roll.base)
        : fCSDArmor(roll.base));
      const pcPostAR = Math.max(1, pcAR * roll.postMult + roll.postPlus);
      const pcARSlideStr =
        pcAR === pcPostAR ? `~${pcPostAR}` : `~(${pcAR}-&gt${pcPostAR})`;
      const pcWnds = Math.round(monDmg / pcPostAR);
      roll.pcWndsAccumulator += pcWnds >= 0 ? pcWnds : 0;
      difficultyStr =
        pcWnds > 0
          ? `/${pcARSlideStr} = ${g.redItS}-${pcWnds}${g.endS}`
          : `/${pcARSlideStr} = ${g.grnBS}0${g.endS}`;
      break;
  }

  return [rollResultStr, difficultyStr];
} // End fCSCalcCombatVsMonResultStringAndDiffStrings

// fCSDoNishExtras //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Add to roll.log: copying final Nish roll to Nish box, channel Meta, Max Actions, listing any Nish TC
function fCSDoNishExtras(game, roll) {
  game.arr[game.nish_R][game.nish_C] = roll.morphedResult;

  fCSMetaRollChnl(game, roll);

  // Add a Nish TC Text Result if needed
  if (roll.finalTC) roll.log += `|<br>${fCSNishCritTrem(roll.finalTC)}`;
  roll.log += `${g.bluBS}==========================${g.endS}<br>`;

  // Set Each Round's Action Total
  const actPlus = Number(game.arr[game.actPlus_R][game.actTotal_C]);
  const finalActPlus = Number.isInteger(actPlus) && actPlus > 0 ? actPlus : 0;
  game.arr[game.actTotal_R][game.actTotal_C] = 5 + finalActPlus;
  fCalcNishOnEffects(game);
} // End fCSDoNishExtras

// fCSMetaRollChnl //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Roll a channeled Meta, updates game.arr and returns the color such as R
function fCSMetaRollChnl(game, roll) {
  const metaColor = ["R", "O", "Y", "G", "B"];

  // Roll the dice to determine channel roll
  const chnlRoll = Math.min(fCSd(5) - 1, fCSd(5) - 1);

  game.arr[game.metaChnl_R][game.metaR_C + chnlRoll]++;

  roll.log = `Channeled: ${metaColor[chnlRoll]}<br>${roll.log}`;
} // End fCSMetaRollChnl

// fCSNishCritTrem //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Generates a Nish Critical or Tremendous text if warranted
function fCSNishCritTrem(nishtc) {
  // Validate nishtc, return '' if not 'c' or 't'
  const isCrit = nishtc === `${g.redBS}C${g.endS}`;
  const isTrem = nishtc === `${g.treS}T${g.endS}`;
  if (!isCrit && !isTrem) return "";

  //DB Constants
  const critArr = gArr("db", "Crit");
  const ct_C = isCrit
    ? gHeaderC("db", "Crit", "NishCrit")
    : gHeaderC("db", "Crit", "NishTrem");
  const dataFirst_R = gDataFirst_R("db", "Crit");
  const dataLast_R = gGetVal("db", "Crit", "maxRow", ct_C);

  // Pick a random result
  const randomRow =
    Math.floor(Math.random() * (dataLast_R - dataFirst_R + 1)) + dataFirst_R;
  let nishCTTxt = isCrit
    ? `${g.redBS}CRITICAL NISH${g.endS}<br>`
    : `${g.treS}TREMENDOUS NISH${g.endS}<br>`;
  nishCTTxt += critArr[randomRow][ct_C];

  return `${nishCTTxt}<br>`;
} // End fCSNishCritTrem

// fCalcNishOnEffects //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Calculates Nish ON Effects (i.e. decrements rounds, etc.)
function fCalcNishOnEffects(game) {
  for (let r = game.abilTblStart_R; r <= game.gearTblEnd_R; r++) {
    // Skip rows between abil and gear tables
    if (r > game.abilTblEnd_R && r < game.gearTblStart_R) {
      r = game.gearTblStart_R - 1;
      continue;
    }

    let abilOn =
      typeof game.arr[r][game.on_C] === "string"
        ? game.arr[r][game.on_C].toUpperCase()
        : game.arr[r][game.on_C];

    switch (true) {
      case abilOn === "E":
        abilOn = "E";
        break;
      case abilOn === "*":
        abilOn = "*";
        break;
      case !isNaN(abilOn) && abilOn >= "2":
        abilOn = (Number(abilOn) - 1).toString();
        break;

      case abilOn === "1":
      case abilOn === "~":
      case abilOn === "Y":
      case abilOn === "N":
      case abilOn === "I":
      case abilOn === "P":
      default:
        abilOn = ".";
    }

    game.arr[r][game.on_C] = abilOn;
  }
} // End fCalcNishOnEffects

// fCSCalcNonCombatRoll //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Rolls a Combat Ability (Atk, Dmg, Def, AR) roll that does have a roll.base > 0
function fCSCalcNonCombatRoll(roll) {
  let thisRollLog = "";

  // Roll in order so can total results if ever wanted
  for (let i = 1; i <= roll.numOfRolls; i++) {
    // Adjust bases if i = 1 for Focus
    const rollCounter = roll.numOfRolls === 1 ? "" : `${i}&gt `;
    const tempMorphedBase =
      i === 1 ? roll.morphedBase + roll.focusPlus : roll.morphedBase;
    const rollBaseStr =
      roll.base === tempMorphedBase
        ? `(${roll.base})`
        : `(${roll.base}-&gt${tempMorphedBase})`;

    // roll.result is either roll.postEquals or UnSkilled (fCSRollUnSkilled) or regular skill (fCSDSk)
    roll.result = roll.postEquals
      ? Math.round(Math.max(1, roll.postEquals))
      : roll.unSkilled
      ? fCSRollUnSkilled(roll.base)
      : fCSDSk(roll.base);
    roll.morphedResult = roll.result * roll.postMult + roll.postPlus;

    const rollResultStr =
      roll.result === roll.morphedResult
        ? `~${roll.result}`
        : `~${roll.result}-&gt${roll.morphedResult}`;
    let difficultyStr = "";
    if (roll.difficulty) {
      let difResult = roll.morphedResult - roll.difficulty;
      difficultyStr =
        difResult >= 0
          ? `^${roll.difficulty} =  ${g.grnBS}+${difResult}${g.endS}`
          : `^${roll.difficulty} = ${g.redItS}${difResult}${g.endS}`; // Adds the '+' for positive results
    }

    roll.finalTC = i === 1 ? fCStc() : "";

    thisRollLog += `${rollCounter}${roll.unSkilled ? "UnSk" : ""}${
      roll.type
    }${rollBaseStr} ${rollResultStr}${difficultyStr} ${roll.finalTC}<br>`;
  }

  const freeOrLucked =
    roll.isFreeRoll || roll.isLuckedRoll
      ? roll.isLuckedRoll
        ? `${g.redBS}LUCKED ROLL${g.endS}<br>`
        : `${g.redBS}FREE ROLL${g.endS}<br>`
      : "";
  const rollTypeString =
    roll.typeOriginal === roll.type
      ? roll.type
      : `${roll.typeOriginal} as ${roll.type}`;
  const morphString = roll.morphString ? `${roll.morphString}<br>` : "";
  roll.log = `${freeOrLucked}${roll.name} ${rollTypeString}<br>${morphString}${thisRollLog}${roll.log}`;
} // End fCSCalcNonCombatRoll

// fCSCalcNoBaseRoll //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Performs a roll where no Sk Number exists (such as an enhancement)
function fCSCalcNoBaseRoll(roll) {
  roll.log += `${roll.name}<br>Activated`;
} // End fCSCalcNoBaseRoll

// fCSPerformUseChaosGem //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Rolls Astral Gauntlet Chaos Gem and decrements uses
function fCSPerformUseChaosGem(game, roll, r) {
  let chaosUses = game.arr[r][game.chaosUses_C];
  const chaosAbilDesc = game.arr[r][game.chaosAbility_C];

  if (!Number.isInteger(chaosUses) || chaosUses <= 0)
    throw new Error(
      `With "${chaosUses} uses," you can't use "${chaosAbilDesc}"`
    );
  chaosUses--;
  game.arr[r][game.chaosUses_C] = chaosUses;

  roll.log += `${g.bluBS}Chaos Crystal${g.endS} (${chaosUses} left):<br>${chaosAbilDesc}<br>`;
} // End fCSPerformUseChaosGem

// fCSRollSk //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Rolls a skill roll and posts it to the roll log
function fCSRollSk(roll, base) {
  const rollResult = fCSDSk(base);

  // Post the result
  roll.finalTC = fCStc();
  roll.log += `${g.bluBS}Custom Sk Roll${g.endS}<br>Sk(${base}) ~${rollResult}${roll.finalTC}<br>`;
} // End fCSRollSk

// fCSRollRPG //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Makes a typical RPG style die roll (e.g., d8, 2d4, 3d6+5, 7d29-14)
function fCSRollRPG(roll, dieString) {
  // Regex pattern to match RPG die notation
  const diePattern = /^(\d*)d(\d+)([+-]\d+)?$/;
  const match = dieString.match(diePattern);

  if (!match) throw new Error(`I can't roll the cell contents "${dieString}"`);

  const numberOfDice = match[1] ? parseInt(match[1]) : 1; // Default to 1 die if no number specified
  const sidesOfDie = parseInt(match[2]);
  const modifier = match[3] ? parseInt(match[3]) : 0;

  let rollResult = 0;
  for (let i = 0; i < numberOfDice; i++) {
    rollResult += fCSd(sidesOfDie); // Roll each die
  }
  rollResult += modifier; // Apply any modifier

  roll.log += `${g.bluBS}Custom d Roll${g.endS}<br>${dieString} ~ ${rollResult}<br>`;
} // End fCSRollRPG

// fCSSaveAbilRollEffects //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Saves all effects of the ability roll
function fCSSaveAbilRollEffects(game, roll) {
  // game.arr[game.rollLog_R][game.rollLog_C] = `${roll.log}<br>${game.arr[game.rollLog_R][game.rollLog_C]}`;

  // Save nearly the entire sheet except rows 0 and 1
  gSaveArraySectionToSheet(
    game.ref,
    game.arr,
    game.meta_R,
    game.dataLast_R,
    game.rollLog_C,
    game.monsterSize_C
  );

  // // Do NOT post the roll to <Game> as that was done above
  const normalLineFeed = false;
  fCSPostMyRoll(roll.log, normalLineFeed);
} // End fCSSaveAbilRollEffects

// fCSRestoreCellFormula //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Restores <Game> cell formula after fCSSaveAbilRollEffects
// Purpose -> Example entry: Formula: AR4,=sum(AN4:AN5)-sum(AO4:AQ5)
function fCSRestoreCellFormula(game) {
  const a1Notes = game.ref.getRange("A1").getNote();

  // Search a1Notes for lines starting with 'Formula: '
  const lines = a1Notes.split("\n");

  lines.forEach((line) => {
    if (line.startsWith("Formula: ")) {
      // Remove "Formula: " from thisLine
      const thisLine = line.replace("Formula: ", "");
      // Split the line into cell reference and formula
      const [myCell, myFormula] = thisLine.split(",=");
      // In game.ref assign the formula myFormula to cell myCell
      game.ref.getRange(myCell).setFormula(`=${myFormula}`);
    }
  });
} // End fCSRestoreCellFormula

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  Die Rolls (end Master Roller)
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
} // End fCSdBetween

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

// fCSDef //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Returns a PC Def roll
function fCSDef(die) {
  die = Number(die);
  return Math.round(fCSD(die) * 0.75);
} // End fCSDef

// fCSRollUnSkilled //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Rolls an ability as UnSkilled - Cuts passed die down a lot, then returns the average of a d(sk*2) + d(sk*2) + an old-fashioned, but smooth, MetaScape d16 roll (no T or C)
function fCSRollUnSkilled(die) {
  const newDie = Math.max(2, Math.min(die / 2, 5 + Math.sqrt(die)));
  return Math.round((fCSd(newDie * 2) + fCSd(newDie * 2) + fCSD(newDie)) / 3);
} // End fCSRollUnSkilled

// fCSRollUnSkilledDmg //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Rolls PC Unskilled Dmg
function fCSRollUnSkilledDmg(pcDmg) {
  return fCSRollUnSkilled(pcDmg);
} // End fCSRollUnSkilledDmg

// fCSRollUnSkilledArmor //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Rolls PC Unskilled Armor
function fCSRollUnSkilledArmor(pcAR) {
  return Math.max(1, Math.trunc((fCSRollUnSkilled(pcAR) / 8) * 10) / 10);
} // End fCSRollUnSkilledArmor

// fCSDDmg //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Rolls PC Dmg, which is a d(regular fDSk(die))
function fCSDDmg(pcDmg) {
  return fCSDSk(pcDmg);
} // End fCSDDmg

// fCSDArmor //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Rolls PC Armor which is a fDSk(armor) roll divided by 8 rounded to the 10ths place such as 1.2 with a minimum AR of .5
function fCSDArmor(pcAR) {
  return Math.max(1, Math.trunc((fCSDSk(pcAR) / 8) * 10) / 10);
} // End fCSDArmor

// fCStc //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> returns "" or 'T' or 'C'
function fCStc() {
  const roll = fCSd(16);
  return roll === 1
    ? `${g.treS}T${g.endS}`
    : roll === 2
    ? `${g.redBS}C${g.endS}`
    : "";
} // End fCStc

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
////////////////////                                    (end g. Die Rolls)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
