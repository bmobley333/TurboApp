
// KL

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  Menu (end initialize)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



// fKLCreateMenu //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Create MetaScape menu
function fKLCreateMenu() {
  SpreadsheetApp.getUi()
  .createMenu('*** MyKLs')
    .addItem('Refresh My KLs', 'fKLMenuRefreshMyKLs')
    .addItem('Calc KLs & AP', 'fKLMenuCalcKLAndAP')
  .addToUi();
  SpreadsheetApp.getUi()
  .createMenu('MyAbilities')
    .addItem('Load Abilities from DB', 'fKLMenuLoadAbilitiesFromDB')
    .addItem('Build MyAbilities from Kits and Rogue', 'fKLMenuBuildMyAbilities')
  .addToUi();
  SpreadsheetApp.getUi()
  .createMenu('Rogue')
    .addItem('Refresh', 'fKLMenuRefreshRogueAbil')
  .addToUi();
  SpreadsheetApp.getUi()
  .createMenu('System ***')
    .addItem('1 - Authorize Script', 'fKLMenuAuthorize')
    .addSeparator()
    .addItem('Refresh Menus', 'fKLMenuRefreshAll')
  .addToUi();


  g.id.mykl = SpreadsheetApp.getActiveSpreadsheet().getId();
  g.id.mycs = gGetVal('mykl','Data','myKLID','Val');
  if (gGetVal('mykl','Data','designer','Val') === true) {
    SpreadsheetApp.getUi()
      .createMenu('DESIGNER')
        .addItem('Hide All', 'fKLMenuHideAll')
        .addItem('Un-Hide All', 'fKLMenuUn_HideAll')
        .addSubMenu(SpreadsheetApp.getUi().createMenu('KLs')
          .addItem('Un-check All CheckBoxes', 'fKLMenuClearAllCheckBoxes')
          .addItem('Update Element Names', 'fKLMenuUpdateKLElementNames')
          .addItem('Set AP Costs', 'fKLMenuSetKLAPCosts')
          .addItem('Set Notes', 'fKLMenuSetKLNotes')
          .addItem('Hide Other RC Tabs', 'fKLMenuHideOtherRCTabs')
          .addItem('Un-Hide All RC Tabs', 'fKLMenuUn_HideAllRCTabs')
        )
      .addToUi();
  } 
} // End fKLCreateMenu



// Menu Functions //////////////////////////////////////////////////////////////////////////////////////////////////
// Kit Menu
function fKLMenuRefreshMyKLs() {fKLRunMenuOrButton('RefreshMyKLs');}
function fKLMenuCalcKLAndAP() {fKLRunMenuOrButton('CalcKLAndAP');}
// Abilies Menu
function fKLMenuLoadAbilitiesFromDB() {fKLRunMenuOrButton('LoadAbilitiesFromDB');}
// Rogue Menu
function fKLMenuRefreshRogueAbil() {fKLRunMenuOrButton('RefreshRogueAbil');}
function fKLMenuBuildMyAbilities() {fKLRunMenuOrButton('BuildMyAbilities');}
// System Menu
function fKLMenuAuthorize() {fKLRunMenuOrButton('Authorize');}
function fKLMenuRefreshAll() {fKLRunMenuOrButton('RefreshMenu');}
// Designer Menu
function fKLMenuHideAll() {fKLRunMenuOrButton('HideAll');}
function fKLMenuUn_HideAll() {fKLRunMenuOrButton('Un_HideAll');}
function fKLMenuUpdateKLElementNames() {fKLRunMenuOrButton('UpdateKLElementNames');}
function fKLMenuClearAllCheckBoxes() {fKLRunMenuOrButton('ClearAllCheckBoxes');}
function fKLMenuSetKLAPCosts() {fKLRunMenuOrButton('SetKLAPCosts');}
function fKLMenuSetKLNotes() {fKLRunMenuOrButton('SetKLNotes');}
function fKLMenuHideOtherRCTabs() {fKLRunMenuOrButton('HideOtherRCTabs');}
function fKLMenuUn_HideAllRCTabs() {fKLRunMenuOrButton('Un_HideOtherRCTabs');}
// End Menu Functions


// Image Button Functions //////////////////////////////////////////////////////////////////////////////////////////////////
function fKLButtonRefreshKits() {fKLRunMenuOrButton('RefreshMyKLs');}
function fKLButtonBuildMyAbilities() {fKLRunMenuOrButton('BuildMyAbilities');}
function fKLButtonRefreshRogue() {fKLRunMenuOrButton('RefreshRogueAbil');}
function fKLButtonLoadAllAbilities() {fKLRunMenuOrButton('LoadAbilitiesFromDB');}
function fKLButtonCalcKLAndAP() {fKLRunMenuOrButton('CalcKLAndAP');}
// End Button Functions


// fKLRunMenuOrButton //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> To run all menu & button choices inside a try-catch-error
function fKLRunMenuOrButton(menuChoice) { 
  try {

    g.id.mykl = SpreadsheetApp.getActiveSpreadsheet().getId();
    g.id.mycs = gGetVal('mykl','Data','myCSID','Val');

    switch (menuChoice) {
      // Kits Menu
      case 'RefreshMyKLs': fKLRefreshMyKLs(); break;
      case 'CalcKLAndAP': fKLCalcKLAndAP(); break;
      // Abilities Menu
      case 'LoadAbilitiesFromDB': fKLLoadAbilitiesFromDB(); break;
      // Rogue Menu
      case 'RefreshRogueAbil': fKLRefreshRogueAbil(); break;
      case 'BuildMyAbilities': fKLBuildMyAbilities(); break;
      // System Menu
      case 'Authorize': SpreadsheetApp.getUi().alert(`AUTHORIZED`, `Script Authorized!`, SpreadsheetApp.getUi().ButtonSet.OK); break;
      case 'RefreshMenu': fKLCreateMenu(); break;
      // Designer Menu
      case 'HideAll': gHideAll('mykl'); break;
      case 'Un_HideAll': gUn_HideAll('mykl'); break;
      case 'UpdateKLElementNames': fKLUpdateKLElementNames(); break;
      case 'ClearAllCheckBoxes': fKLClearAllCheckBoxes(); break;
      case 'SetKLAPCosts': fKLSetKLAPCosts(); break;
      case 'SetKLNotes': fKLSetKLNotes(); break;
      case 'HideOtherRCTabs': fKLGetRCAndHideOtherRCTabs(); break;
      case 'Un_HideOtherRCTabs': fKLUn_HideOtherRCTabs(); break;
    }
  } catch (error) {
      SpreadsheetApp.getUi().alert(error); // NOTE: an error of End or end will simply end the program.

  }
} // End fKLRunMenuOrButton


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  KLs  (end Menu)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


/**
 * Purpose: A master function to perform a full calculation of all KL sheet AP values and update the header information.
 * Input:
 * -- none
 * Output: Void (modifies the KL sheets directly).
 */
function fKLCalcKLAndAP() {

    // Get the character's active RC tab name, hiding other RC tabs in the process.
    const myRCTabName = fKLGetRCAndHideOtherRCTabs();
    const myUsedRCTabs = ['All', myRCTabName];

    // Verify RC Checkboxes
    const tierString = gGetVal('mykl', 'All', 'Tier', 'Tier');
    const tierMatch = String(tierString).match(/\d+/);
    const tierNum = tierMatch ? parseInt(tierMatch[0], 10) : 0;
    fKLVerifyRCCheckedBoxes(myUsedRCTabs,tierNum);

    // Calculate AP Spent
    const apSpent = {
        combat: 0,
        base: 0,
    };
    fKLAPSpentCalc(myUsedRCTabs, apSpent);

    // Build and save final header info
    const headerInfo = fKLBuildHeaderObject(apSpent);
    fKLSaveRCHeaderInfo(myUsedRCTabs, headerInfo);


} // End fKLCalcKLAndAP








/**
 * Purpose: Calculates the total spent combat and base AP from a given list of KL sheets by summing costs from checked abilities.
 * Notes: This function modifies the apSpent object directly (pass by reference).
 * Input:
 * -- myRCTabName - {string[]} An array of KL sheet names to process.
 * -- apSpent - {object} An object with 'combat' and 'base' properties to accumulate the costs into.
 * Output: Void (modifies the input apSpent object).
 */
function fKLAPSpentCalc(myRCTabName, apSpent) {

    // Reset the AP counters to zero before calculation.
    apSpent.combat = 0;
    apSpent.base = 0;

    // Iterate through each provided sheet name.
    myRCTabName.forEach(tabName => {
        const currentTab = getObjKL_KLTab(tabName, true);

        const dataFirstTF_C = 1;
        const dataLastTF_C = currentTab.arr[0].length - 1;

        for (let r = currentTab.dataFirst_R; r <= currentTab.dataLast_R; r++) {
            for (let c = dataFirstTF_C; c < dataLastTF_C; c++) {

                // If a checkbox in the current cell is checked, process the adjacent card.
                if (currentTab.arr[r][c] === true) {

                    const cardText = currentTab.arr[r][c + 1];
                    if (!cardText) continue;

                    const klCard = fGetKLCardObj(cardText, tabName, r, c + 1);

                    // Add the card's AP cost to the appropriate counter based on its type.
                    switch (klCard.apType) {
                        case 'C': apSpent.combat += klCard.apCost; break;
                        case 'B': apSpent.base += klCard.apCost; break;
                        case 'F': break;
                        default: throw new Error(`In sheet "${tabName}" the KL card "${cardText}" has an illegal AP Type of "${klCard.apType}"`);
                    }
                }
            }
        }
    });

} // End fKLAPSpentCalc





/**
 * Purpose: Calculates all header AP, Level, and Tier values and assembles them into a single object using pre-calculated spent AP.
 * Assumptions: The apSpent object is provided and contains valid 'combat' and 'base' number properties.
 * Notes: This refactor makes the function dependent on an input object for spent AP values rather than reading them from the sheet directly.
 * Input:
 * -- apSpent - {object} An object containing the spent AP values, e.g., { combat: 10, base: 5 }.
 * Output: An object containing all calculated header values.
 */
function fKLBuildHeaderObject(apSpent) {

    //
    // First, calculate all the necessary values in sequence.
    const myLevel = gCharLvl();
    const levelAP = 10 * myLevel;
    const tierString = gGetVal('mykl', 'All', 'Tier', 'Tier');
    const tierMatch = String(tierString).match(/\d+/);
    const tierNum = tierMatch ? parseInt(tierMatch[0], 10) : 0;
    const rcName_ID = gGetVal('mykl', 'All', 'RC', 'RC');
    const bnsAP = gGetVal('mykl', 'All', 'BnsAP', 'APCount');
    const totalAP = levelAP + bnsAP;
    const apCombat = totalAP;
    const apBase = Math.round(0.3 * totalAP);
    const apCombatSpent = apSpent.combat;
    const apBaseSpent = apSpent.base;
    const apCombatRemaining = apCombat - apCombatSpent;
    const apBaseRemaining = apBase - apBaseSpent;

    //
    // Then, assemble the final object using the calculated constants.
    const header = {
        myLevel,
        tierString,
        tierNum,
        rcName_ID,
        levelAP,
        bnsAP,
        totalAP,
        apCombat,
        apBase,
        apCombatSpent,
        apBaseSpent,
        apCombatRemaining,
        apBaseRemaining,
    };

    return header;

} // End fKLBuildHeaderObject






/**
 * Purpose: Saves the calculated header AP, Level, and Tier info to a given list of RC sheets.
 * Assumptions: The header object 'h' has been pre-calculated by fKLBuildHeaderObject().
 * Input:
 * -- myUsedRCTabs - {string[]} An array of KL sheet names to update.
 * -- h - {object} The header object containing the values to be saved.
 * Output: Void (modifies the specified KL sheets directly).
 */
function fKLSaveRCHeaderInfo(myUsedRCTabs, h) {

    // For each provided sheet name, set all header values from the header object, then save the sheet.
    myUsedRCTabs.forEach(tabName => {

        gSetVal('mykl', tabName, 'MyLvl', 'MyLvl', h.myLevel);
        gSetVal('mykl', tabName, 'Tier', 'Tier', h.tierString);
        gSetVal('mykl', tabName, 'RC', 'RC', h.rcName_ID);
        gSetVal('mykl', tabName, 'LevelAP', 'APCount', h.levelAP);
        gSetVal('mykl', tabName, 'BnsAP', 'APCount', h.bnsAP);
        gSetVal('mykl', tabName, 'TotalAP', 'APCount', h.totalAP);
        gSetVal('mykl', tabName, 'APCombat', 'APTotal', h.apCombat);
        gSetVal('mykl', tabName, 'APBase', 'APTotal', h.apBase);
        gSetVal('mykl', tabName, 'APCombat', 'APSpent', h.apCombatSpent);
        gSetVal('mykl', tabName, 'APBase', 'APSpent', h.apBaseSpent);
        gSetVal('mykl', tabName, 'APCombat', 'APRemaining', h.apCombatRemaining);
        gSetVal('mykl', tabName, 'APBase', 'APRemaining', h.apBaseRemaining);

        gSaveSheet('mykl', tabName);
    });

} // End fKLSaveRCHeaderInfo




/**
 * Purpose: Verifies all checkboxes in the data rows of specified KL sheets, unchecking any that are invalid and checking any that are 'Free' and valid.
 * Assumptions: The header object 'h' contains the character's current tier number (h.tierNum).
 * Notes: This function enforces the rule that a higher-tier ability cannot be selected if the tier directly above it is not selected.
 * Input:
 * -- myUsedRCTabs - {string[]} An array of KL sheet names to process.
 * -- tierNum - containing the character's tier.
 * Output: Void (modifies the specified KL sheets directly).
 */
function fKLVerifyRCCheckedBoxes(myUsedRCTabs, tierNum) {

    // Iterate through each provided sheet name.
    myUsedRCTabs.forEach(tabName => {
        const currentTab = getObjKL_KLTab(tabName, true);

        const lastCol = currentTab.arr[0].length - 2; // Loop until the second to last column to safely access c+1

        for (let r = currentTab.dataFirst_R; r <= currentTab.dataLast_R; r++) {
            for (let c = 1; c <= lastCol; c++) {

                // If there is a checkbox (true or false) in the current cell, verify it.
                if (currentTab.arr[r][c] === true || currentTab.arr[r][c] === false) {

                    const cardText = currentTab.arr[r][c + 1];
                    if (!cardText) continue;

                    const klCard = fGetKLCardObj(cardText, tabName, r, c + 1);

                    // Check if the ability directly above this one in the same column is checked.
                    const isDependentAboveUn_Checked = (currentTab.arr[r - 1][c] === false);

                    // Verify checkbox is allowed based on character tier and dependency on the ability above it.
                    if (klCard.tier > tierNum || isDependentAboveUn_Checked) {
                        currentTab.arr[r][c] = false;
                    } else if (klCard.apType === 'F') {
                      
                        // Automatically check 'Free' abilities (already know they are at or below PC's tier)
                        currentTab.arr[r][c] = true;
                    }
                }
            }
        }

        gSaveSheet('mykl', tabName);
    });

} // End fKLVerifyRCCheckedBoxes




/**
 * Updates element names in specified 'KL' sheets.
 * It iterates through each sheet, finds rows marked with 'true' in the first column,
 * looks up the element ID from the second column in a central database,
 * and writes the corresponding element name back into that cell.
 */
function fKLUpdateKLElementNames() {

  g.klRCSheetNames.forEach(tabName => {
    // Get the object for the current tab, forcing a reload to ensure data is current.
    const currentTab = getObjKL_KLTab(tabName, true);

    // Loop through each data row of the current tab's array.
    // Note: Array row 'r' is 0-indexed, while Sheet rows are 1-indexed.
    const dataFirstTF_C = 1;
    const dataLastTF_C = currentTab.arr[0].length -1;
    for (let r = currentTab.dataFirst_R; r <= currentTab.dataLast_R; r++) {
      for (let c = dataFirstTF_C; c < dataLastTF_C; c++) {
      
        // Check if the cell is a boolean value.
        if (currentTab.arr[r][c] === true || currentTab.arr[r][c] === false) {
          
          const cardText = currentTab.arr[r][c+1]; // Get value from the next column

          const klCard = fGetKLCardObj(cardText,tabName,r,c+1);
          
          // Check if the card's ID exists in the 'Elements' database.
          if (gTestID('db', 'Elements', klCard.id)) {
            // If it exists, get the official name from the database.
            const newName = gGetVal('db', 'Elements', klCard.id, 'Name');
            
            // Write the new name back to the sheet.
            fKLNewKLName(currentTab, klCard, newName, r, c+1);
          } else {
            // If the ID is not found, throw a detailed error.
            throw new Error(`ID lookup failed for sheet "${tabName}". The ID "${klCard.id}" (derived from cell ${r+1},${c+2} with value "${cardText}") was not found in the 'Elements' database.`);
          }
        }
      }
    }
    gSaveSheet('mykl', tabName);
  });

} // End fKLUpdateKLElementNames




/**
 * Purpose: Parses a formatted 3-line string from a "KL Card" into a structured object.
 * Assumptions: cardText must be a string with exactly two newline characters, creating three lines of text.
 * Input: 
 *   cardText - A formatted string containing the card's name, cost, tier, ID, and buff info.
 *   tabName - The name of the sheet where the cardText is located.
 *   r - The 0-indexed row of the cell.
 *   c - The 0-indexed column of the cell.
 * Output: A structured object (cardObj) with properties for name, apType, apCost, tier, id, buffVerType, and buffVerNum.
 */
function fGetKLCardObj(cardText, tabName, r, c) {

  // Input Validation
  if (typeof cardText !== 'string' || !cardText.trim()) {
    throw new Error(`In fGetKLCardObj, the cardText from sheet "${tabName}" at cell ${r},${c} was empty or not a string.`);
  }

  // Test for exactly two newline characters, which creates an array of three lines.
  const lines = cardText.split('\n');
  if (lines.length !== 3) {
    throw new Error(`In fGetKLCardObj, KLCard from sheet "${tabName}" at cell ${r},${c} does not contain exactly two newline characters. It must be three separate lines.`);
  }

  // Store each line in its own variable for simpler parsing.
  const nameLine = lines[0];
  const costTierLine = lines[1];
  const idVerLine = lines[2];

  const cardObj = {};
  
  // --- Parse each line individually ---

  // Parse Name (from the first line)
  cardObj.name = nameLine.trim();

  // Parse AP Type and Cost (from the second line)
  // Updated regex to handle 'Free' as well as 'F', 'C', and 'B' (case-insensitive)
  const apMatch = costTierLine.match(/\$(Free|[CBF])(\d{1,3})?/i);
  if (!apMatch) throw new Error(`In fGetKLCardObj, the cost/tier line from sheet "${tabName}" at cell ${r},${c} ("${costTierLine}") has an invalid or missing AP cost block (e.g., $C5, $Free, $F).`);
  
  let [, apType, apCostStr] = apMatch;
  apType = apType.toUpperCase(); // Normalize to uppercase

  // Check if the type is 'F' (for Free)
  if (apType === 'F' || apType === 'FREE') {  // Previously Normalize to uppercase
    cardObj.apType = 'F'; // Standardize the output type to 'F'
    cardObj.apCost = 0;
  } else { // Must be 'C' or 'B'
    cardObj.apType = apType;
    cardObj.apCost = parseInt(apCostStr, 10);
    if (isNaN(cardObj.apCost) || cardObj.apCost < 1 || cardObj.apCost > 100) {
      throw new Error(`In fGetKLCardObj, the cost/tier line from sheet "${tabName}" at cell ${r},${c} ("${costTierLine}") has an apCost that is not an integer between 1 and 100.`);
    }
  }

  // Parse Tier (from the second line)
  const tierMatch = costTierLine.match(/T(\d+)/);
  if (!tierMatch) throw new Error(`In fGetKLCardObj, the cost/tier line from sheet "${tabName}" at cell ${r},${c} ("${costTierLine}") is missing the Tier indicator (e.g., T2).`);
  cardObj.tier = parseInt(tierMatch[1], 10);

  // Parse ID (from the third line)
  const idMatch = idVerLine.match(/🔑(.{6})/);
  if (!idMatch) throw new Error(`In fGetKLCardObj, the ID/version line from sheet "${tabName}" at cell ${r},${c} ("${idVerLine}") is missing the '🔑' ID indicator.`);
  cardObj.id = idMatch[1];

  // Parse Buff/Version Type and Number (from the third line)
  const buffVerMatch = idVerLine.match(/\.([bv])([1-9])$/);
  if (!buffVerMatch) throw new Error(`In fGetKLCardObj, the ID/version line from sheet "${tabName}" at cell ${r},${c} ("${idVerLine}") has an invalid or missing buff/version suffix (e.g., .b1, .v9).`);
  
  [, cardObj.buffVerType, cardObj.buffVerNum] = buffVerMatch;
  cardObj.buffVerNum = parseInt(cardObj.buffVerNum, 10);

  return cardObj;
} // End fGetKLCardObj



/**
 * Purpose: Updates the name of a KL Card object and rebuilds the source string in the provided array.
 * Input:
 *   currentTab - The cached sheet object containing the .arr to be modified.
 *   klCard - The parsed card object to be updated.
 *   elementName - The new name for the card.
 *   r - The 0-indexed row in the array to update.
 *   c - The 0-indexed column in the array to update.
 * Output: Void (modifies the currentTab.arr by reference).
 */
function fKLNewKLName(currentTab, klCard, elementName, r, c) {
  
  // Check if the new name is different and a non-empty string
  if (klCard.name !== elementName && typeof elementName === 'string' && elementName) {
    klCard.name = elementName;
    currentTab.arr[r][c] = fKLCardObjToStr(klCard);
  }
} // End fKLNewKLName




/**
 * Purpose: Converts a KL Card object back into its three-line formatted string representation.
 * Assumptions: The klCard object has all the necessary properties (name, apType, apCost, etc.).
 * Input: klCard - The object representing the card's data.
 * Output: A formatted, three-line string representing the card.
 */
function fKLCardObjToStr(klCard) {
  
  // Conditionally format the AP cost string
  const apCostString = klCard.apType === 'F' ? 'Free' : klCard.apType + klCard.apCost;

  return `${klCard.name}\n$${apCostString} T${klCard.tier}\n🔑${klCard.id}.${klCard.buffVerType}${klCard.buffVerNum}`;
} // End fKLCardObjToStr




/**
 * Purpose: A wrapper function that unchecks all 'true' checkboxes across all standard KL RC sheets.
 * @returns {void}
 */
function fKLClearAllCheckBoxes() {

    // Call the core function with the global list of all RC sheet names.
    fKLClearAllCheckBoxesFrom(g.klRCSheetNames);

} // End fKLClearAllCheckBoxes



/**
 * Purpose: Iterates through a provided list of KL RC sheets and unchecks any cell containing a boolean 'true' value.
 * Notes: This is the core logic function; it is called by other checkbox-clearing functions.
 * @param {(string|string[])} sheetNameList - A single sheet name or an array of sheet names to process.
 * @returns {void}
 */
function fKLClearAllCheckBoxesFrom(sheetNameList) {

    // Ensure the input is an array so .forEach can be used reliably.
    const sheetNames = Array.isArray(sheetNameList) ? sheetNameList : [sheetNameList];

    // For each specified sheet, load it, change all 'true' values to 'false', and save it.
    sheetNames.forEach(tabName => {
        const currentTab = getObjKL_KLTab(tabName, true);
        const lastCol = currentTab.arr[0].length - 1;

        for (let r = currentTab.dataFirst_R; r <= currentTab.dataLast_R; r++) {
            for (let c = 1; c <= lastCol; c++) {
                if (currentTab.arr[r][c] === true) {
                    currentTab.arr[r][c] = false;
                }
            }
        }
        gSaveSheet('mykl', tabName);
    });

} // End fKLClearAllCheckBoxesFrom




/**
 * Purpose: Recalculates the AP Cost for all abilities in the KL RC-style sheets based on their tier progression.
 * Assumptions: The klRCSheetNames array is globally available at g.klRCSheetNames.
 * Input:
 * none
 * Output: Void (modifies the KL sheets directly).
 */
function fKLSetKLAPCosts() {

    const firstTierMap = {};
    const buffAPCost = [2, 4, 8, 16, 32];
    const verAPCost = [5, 5, 9, 16, 25];

    // Alert the user that the <All> sheet has a special condition.
    SpreadsheetApp.getUi().alert('Notice', 'Abilities on the KL <All> sheet will only be updated from row 18 onwards (Skipping Attributes). All other RC sheets will be fully processed.', SpreadsheetApp.getUi().ButtonSet.OK);

    // For each specified sheet, load it, recalculate AP costs, and save it.
    g.klRCSheetNames.forEach(tabName => {
        const currentTab = getObjKL_KLTab(tabName, true);
        const lastCol = currentTab.arr[0].length - 2; // Loop until the second to last column to safely access c+1

        for (let r = currentTab.dataFirst_R; r <= currentTab.dataLast_R; r++) {

            // If on the 'All' sheet, skip the protected Attributes rows.
            if (tabName.toLowerCase() === 'all' && r < 17) {
                continue;
            }

            for (let c = 1; c <= lastCol; c++) {
                
                // Check if the cell contains a boolean value to process.
                if (currentTab.arr[r][c] === true || currentTab.arr[r][c] === false) {

                    const cardText = currentTab.arr[r][c + 1];
                    if (!cardText) throw new Error(`In fKLSetKLAPCosts sheet "${tabName}" at cell ${r},${c + 1} there is no KL Card after this check box.`);

                    const klCard = fGetKLCardObj(cardText, tabName, r, c + 1);

                    // If this is the first time seeing this ability ID, store its tier as the base tier.
                    if (!firstTierMap.hasOwnProperty(klCard.id)) {
                        firstTierMap[klCard.id] = klCard.tier;
                    }

                    // Skip cost calculation for 'Free' abilities.
                    if (klCard.apType === 'F') {
                        continue;
                    }

                    // Calculate and validate the cost tier.
                    const baseTier = firstTierMap[klCard.id];
                    const apCostTier = klCard.tier - baseTier;

                    if (!Number.isInteger(apCostTier) || apCostTier < 0 || apCostTier > 4) {
                        throw new Error(`In fKLSetKLAPCosts sheet "${tabName}" at cell ${r + 1},${c + 2} with value "${cardText}" has an improper Tier progression. The calculated cost tier was ${apCostTier}, but it must be an integer between 0 and 4.`);
                    }

                    // Assign the new AP cost based on the tier and buff/version type.
                    klCard.apCost = (klCard.buffVerType === 'v') ? verAPCost[apCostTier] : buffAPCost[apCostTier];
                    currentTab.arr[r][c + 1] = fKLCardObjToStr(klCard);
                }
            }
        }
        gSaveSheet('mykl', tabName);
    });

} // End fKLSetKLAPCosts


/**
 * Purpose: Populates the notes for the first instance of each unique ability card in the KL RC-style sheets, preserving any existing notes in the header rows.
 * Notes: This function now treats the first 'buff' ('b') and the first 'version' ('v') of an ability as separate instances for the purpose of adding notes.
 * Input:
 * none
 * Output: Void (modifies the notes of the KL sheets directly).
 */
function fKLSetKLNotes() {

    const firstInstanceMap = {};

    g.klRCSheetNames.forEach(tabName => {
        const currentTab = getObjKL_KLTab(tabName, true);
        const numRows = currentTab.arr.length;
        const numCols = currentTab.arr[0].length;

        //
        // Read all existing notes from the sheet to preserve the header notes.
        const noteArr = currentTab.ref.getRange(1, 1, numRows, numCols).getNotes();

        const lastCol = numCols - 2; // Loop until the second to last column to safely access c+1

        for (let r = currentTab.dataFirst_R; r <= currentTab.dataLast_R; r++) {
            for (let c = 1; c <= lastCol; c++) {

                // Erase any old note in the non-header row to ensure a clean slate for this run.
                noteArr[r][c + 1] = null;
                
                // Check if the cell contains a boolean value to process.
                if (currentTab.arr[r][c] === true || currentTab.arr[r][c] === false) {

                    const cardText = currentTab.arr[r][c + 1];
                    if (!cardText) continue;

                    const klCard = fGetKLCardObj(cardText, tabName, r, c + 1);

                    // Create a unique key by combining the ID and the buff/version type ('b' or 'v').
                    const uniqueKey = klCard.id + klCard.buffVerType;

                    // If this is the first time seeing this specific ID + type combination, get the note.
                    if (!firstInstanceMap.hasOwnProperty(uniqueKey)) {
                        firstInstanceMap[uniqueKey] = true; // Mark this combo as seen
                        noteArr[r][c + 1] = gGetVal('db', 'Elements', klCard.id, 'Notes');
                    }
                }
            }
        }

        //
        // Save the entire notes array, now containing both old header notes and new ability notes, in a single operation.
        currentTab.ref.getRange(1, 1, numRows, numCols).setNotes(noteArr);
    });

} // End fKLSetKLNotes



/**
 * Purpose: Hides all RC-related sheets in the KeyLine except for the <All> sheet and the one currently selected on the Character Sheet.
 * Assumptions: The g.klRCSheetNames and g.matchingKLRCIDs global arrays are parallel and correctly populated.
 * Notes: This function will now skip clearing and hiding sheets that are already hidden.
 * Input:
 * -- none
 * Output: myRCTabName
 */
function fKLGetRCAndHideOtherRCTabs() {

    const ssRef = gSSRef('mykl');
    const rcName_ID = gGetVal('mycs', 'RaceClass', 'RC', 'Val');
    gSaveVal('mykl', 'All', 'RC', 'RC', rcName_ID);

    // Validate that a RaceClass has been selected on the Character Sheet.
    if (!rcName_ID || typeof rcName_ID !== 'string') {
        throw new Error(`You need to select a RaceClass on the <RaceClass> tab of your Character Sheet.`);
    }

    // Parse the ID from the RaceClass string and find its index in the global ID list.
    const rcID = gGetIDFromString(rcName_ID);
    const i = g.matchingKLRCIDs.indexOf(rcID);

    // If the ID isn't found in our list, throw an error.
    if (i === -1) {
        throw new Error(`In fKLGetRCAndHideOtherRCTabs the ID "${rcID}" from your selected RaceClass was not found in the g.matchingKLRCIDs list.`);
    }

    // Use the found index to get the corresponding tab name from the parallel array.
    const myRCTabName = g.klRCSheetNames[i];

    // Iterate through all sheets and hide the ones that are in the RC list but are not already hidden or 'All' or the selected RC.
    const sheets = ssRef.getSheets();
    sheets.forEach(sheet => {
        const tabName = sheet.getName();
        if (g.klRCSheetNames.includes(tabName) && tabName !== 'All' && tabName !== myRCTabName && !sheet.isSheetHidden()) {
            fKLClearAllCheckBoxesFrom(tabName);
            sheet.hideSheet();
        } else if (tabName === myRCTabName) {
            sheet.showSheet();
        }
    });

    return myRCTabName;

} // End fKLGetRCAndHideOtherRCTabs




/**
 * Purpose: Un-hides all sheets listed in the g.klRCSheetNames global array.
 * Input:
 * -- none
 * Output: Void (modifies the visibility of sheets).
 */
function fKLUn_HideOtherRCTabs() {

    const ssRef = gSSRef('mykl');

    // Iterate through the global list of RC sheet names and unhide each one.
    g.klRCSheetNames.forEach(tabName => {
        const sheet = ssRef.getSheetByName(tabName);
        if (sheet && sheet.isSheetHidden()) {
            sheet.showSheet();
        }
    });

} // End fKLUn_HideOtherRCTabs


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  Old MyKLs (menu)  (end KLs)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////





// fKLRefreshMyKLs //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Refreshes and Calculates MyKLs choices
function fKLRefreshMyKLs() {

  fKLLoadDBKits();
  // Load KL <MyKLs> Info
  let klMyKLs = getObjKLMyKLs();

  for (let r = klMyKLs.dataFirst_R; r <= klMyKLs.dataLast_R; r++) { // Adjusted to reference from klKits object
    let myKitNameID = klMyKLs.arr[r][klMyKLs.nameID_C];
    const myKitID = gTestID('db', 'Kits', myKitNameID);
    if (myKitID) {
      myKitNameID = gGetVal('db', 'Kits', myKitID, 'Name_ID');
      const myNotes = gGetVal('db', 'Kits', myKitID, 'Notes');
      klMyKLs.arr[r][klMyKLs.nameID_C] = myKitNameID;
      klMyKLs.arr[r][klMyKLs.notes_C] = myNotes;
    } else {
      klMyKLs.arr[r][klMyKLs.nameID_C] = '';
      klMyKLs.arr[r][klMyKLs.notes_C] = '';
    }
  }

  gSaveArraySectionToSheet(klMyKLs.ref, klMyKLs.arr, klMyKLs.dataFirst_R, klMyKLs.dataLast_R, klMyKLs.nameID_C, klMyKLs.notes_C);

  // force reload of klMyKLs
  klMyKLs = getObjKLMyKLs(true);

  // **** Clear all but (Key, NameID, Base1, Base2) of AllAbilities, then ... Populate AllAbilities based on MyKLs and Rouge
  fKLKitsAndRougeToAllAbilities();

  // **** Build MyAbilities ****
  fKLBuildMyAbilities();

} // End fKLRefreshMyKLs




// fKLLoadDBKits //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Populates <KLAll> with all eligible KeyLines
function fKLLoadDBKits() {

  // Load DB <Kits> Info
  let dbKits = getObjDBKits();

  // Load KL <List> Info
  let klList = getObjKLList();
  let needMoreArrRows = 0;

  gFillArraySection(klList.arr, klList.dataFirst_R, klList.dataLast_R, klList.kits_C, klList.kits_C, '');

  let listR = klList.dataFirst_R; 
  for (let r = dbKits.dataFirst_R; r <= dbKits.dataLast_R; r++) {
    const dbKitNameId = dbKits.arr[r][dbKits.nameID_C];
    if (!dbKitNameId) break; // If dbKitNameID is empty then end of actual data in db <Kits>

    // Check if there is space in the array; if not, resize it
    if (listR >= klList.arr.length) {
      klList.arr.push(Array(klList.arr[0].length).fill(''));  // Assuming all columns are initially empty
      needMoreArrRows++;
    }

    klList.arr[listR][klList.kits_C] = dbKitNameId;
    listR++;
  }

  // If additional rows are needed, add them to the bottom of the sheet
  if (needMoreArrRows > 0) {
    klList.ref.insertRowsAfter(klList.ref.getMaxRows(), needMoreArrRows);
  }

  gSortArraySection(klList.arr, klList.dataFirst_R, listR - 1, klList.kits_C, klList.kits_C,klList.kits_C);
  gSaveArraySectionToSheet(klList.ref, klList.arr, klList.dataFirst_R, listR - 1, klList.kits_C, klList.kits_C);

  // force load the changed table
  getObjKLList(true);

} // End fKLLoadDBKits



// fKLKitsAndRougeToAllAbilities //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Clear all but (Key, NameID, Base1, Base2) of AllAbilities, then ... Populate AllAbilities based on MyKLs and Rouge
function fKLKitsAndRougeToAllAbilities() {

  let klAllAbil = getObjKLAllAbilities();
  gFillArraySection(klAllAbil.arr,klAllAbil.dataFirst_R,klAllAbil.dataLast_R,klAllAbil.source_C,klAllAbil.source_C,'');
  gFillArraySection(klAllAbil.arr,klAllAbil.dataFirst_R,klAllAbil.dataLast_R,klAllAbil.cantLearn_C,klAllAbil.learn_C,false);
  gFillArraySection(klAllAbil.arr,klAllAbil.dataFirst_R,klAllAbil.dataLast_R,klAllAbil.morphOther_C,klAllAbil.last_C,'');
  fKLKitsToAllAbilities();
  fKLRogueToAllAbilities();

  fKLCalcMorphs();

} // End fKLKitsAndRougeToAllAbilities




// fKLKitsToAllAbilities //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: For Each KL in <MyKLs> process the MorphString and apply it to <AllAbilities>
function fKLKitsToAllAbilities() {

  let klMyKLs = getObjKLMyKLs();

  // *** For Each KL in <MyKLs> 
  for (let r = klMyKLs.dataFirst_R; r <= klMyKLs.dataLast_R; r++) {
    const myKitNameID = klMyKLs.arr[r][klMyKLs.nameID_C];
    const myKitID = gTestID('db', 'Kits', myKitNameID);
    if (myKitID) {
      const myKitMorphStr = gGetVal('db', 'Kits', myKitID, 'MorphString');
      fKLKitMorphStrToAllAbilities(myKitNameID, myKitMorphStr);
    }
  }

  // Force load due to changes
  getObjKLAllAbilities(true);

} // End fKLKitsToAllAbilities



// fKLKitMorphStrToAllAbilities //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Apply a Kit's MorphString to <AllAbilities>
function fKLKitMorphStrToAllAbilities(myKitNameID, myKitMorphStr) {

  // Load KL <AllAbilities> Info
  let klAllAbilities = getObjKLAllAbilities();
  const klAllAbilArr = klAllAbilities.arr;

  const myKitName = myKitNameID.split('   ')[0];
  const abilMorphArr = myKitMorphStr.split('\\').slice(1); // Remove the first empty element
  for (const ability of abilMorphArr) {
    const [abilID, otherMorph, sk1Morph, sk2Morph] = ability.split('|');
    const allAbil_R = gKeyR('mykl', 'AllAbilities', abilID);

    klAllAbilArr[allAbil_R][klAllAbilities.source_C] += (klAllAbilArr[allAbil_R][klAllAbilities.source_C]) ? `,${myKitName}` : myKitName;
    klAllAbilArr[allAbil_R][klAllAbilities.morphOther_C] += (klAllAbilArr[allAbil_R][klAllAbilities.morphOther_C]) ? `,${otherMorph}` : otherMorph;
    klAllAbilArr[allAbil_R][klAllAbilities.sk1PLAGHE_C] += (klAllAbilArr[allAbil_R][klAllAbilities.sk1PLAGHE_C]) ? `,${sk1Morph}` : sk1Morph;
    klAllAbilArr[allAbil_R][klAllAbilities.sk2PLAGHE_C] += (klAllAbilArr[allAbil_R][klAllAbilities.sk2PLAGHE_C]) ? `,${sk2Morph}` : sk2Morph;
  }

  gSaveArraySectionToSheet(klAllAbilities.ref,klAllAbilArr,klAllAbilities.dataFirst_R,klAllAbilities.dataLast_R,klAllAbilities.first_C,klAllAbilities.last_C);
  // Force reload
  klAllAbilities = getObjKLAllAbilities(true);

} // End fKLKitMorphStrToAllAbilities


// fKLRogueToAllAbilities //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Apply <RogueAbilities> to <AllAbilities>
function fKLRogueToAllAbilities() {

  let klRogueAbilities = getObjKLRogueAbilities();
  
  // ***** Save (arr and table) Rogue abilities as Learnable and Known if appropriate
  for (let r = klRogueAbilities.dataFirst_R; r <= klRogueAbilities.dataLast_R; r++) {
    let klRogueID = klRogueAbilities.arr[r][klRogueAbilities.id_C];
    klRogueID = gTestID('mykl','AllAbilities',klRogueID)
    if (klRogueID) {
      let source = gGetVal('mykl','AllAbilities',klRogueID,'Source');
      source += (source) ? `,Rouge` : 'Rouge';
      gSaveVal('mykl','AllAbilities',klRogueID,'Source',source);

      const isKnown = klRogueAbilities.arr[r][klRogueAbilities.known_C];
      let morphOther = gGetVal('mykl','AllAbilities',klRogueID,'MorphOther');
      if (isKnown) {
        morphOther += (morphOther) ? `,KNOWN` : 'KNOWN';
      } else {
        morphOther += (morphOther) ? `,LEARN` : 'LEARN';
      }
      gSaveVal('mykl','AllAbilities',klRogueID,'MorphOther',morphOther);
    }
  }

  // Force load due to changes
  getObjKLAllAbilities(true);

} // END fKLRogueToAllAbilities



// fKLCalcMorphs //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: In kl <AllAbilities>, applies all Morph Effects
function fKLCalcMorphs() {

  let klAllAbil = getObjKLAllAbilities();

  for (let r = klAllAbil.dataFirst_R; r <= klAllAbil.dataLast_R; r++) {

    // Convert values to string before splitting to handle non-string types like numbers
    const morphOtherArr = String(klAllAbil.arr[r][klAllAbil.morphOther_C]).split(',');
    const sk1PLAGHEArr = String(klAllAbil.arr[r][klAllAbil.sk1PLAGHE_C]).split(',');
    const sk2PLAGHEArr = String(klAllAbil.arr[r][klAllAbil.sk2PLAGHE_C]).split(',');

    const cantLearn = morphOtherArr.includes('~');
    const known = !cantLearn && morphOtherArr.includes('KNOWN');
    const learn = !cantLearn && (known || morphOtherArr.includes('LEARN'));

    klAllAbil.arr[r][klAllAbil.cantLearn_C] = cantLearn;
    klAllAbil.arr[r][klAllAbil.known_C] = known;
    klAllAbil.arr[r][klAllAbil.learn_C] = learn;

    // Calculate the FinalSk1 and FinalSk1 bases
    const base1 = klAllAbil.arr[r][klAllAbil.base1_C];
    const base2 = klAllAbil.arr[r][klAllAbil.base2_C];
    klAllAbil.arr[r][klAllAbil.finalSk1_C] = fKLcalcFinalSk(base1, sk1PLAGHEArr);
    klAllAbil.arr[r][klAllAbil.finalSk2_C] = fKLcalcFinalSk(base2, sk2PLAGHEArr);
  }

  gSaveArraySectionToSheet(klAllAbil.ref, klAllAbil.arr, klAllAbil.dataFirst_R, klAllAbil.dataLast_R, klAllAbil.cantLearn_C, klAllAbil.learn_C);
  gSaveArraySectionToSheet(klAllAbil.ref, klAllAbil.arr, klAllAbil.dataFirst_R, klAllAbil.dataLast_R, klAllAbil.finalSk1_C, klAllAbil.finalSk2_C);

  // Force reload
  klAllAbil = getObjKLAllAbilities(true);
} // END fKLCalcMorphs




// fKLcalcFinalSk //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Helper Function to calculate final skill value and if there is no base then just return ''
// Purpose: returns base * best ratingMult in morphArr + all ratingPlus in morphArr divided by (1,2,4,8...)
// Purpose: Minimum skill is 2
function fKLcalcFinalSk(base, morphArr) {

  if (!base || (morphArr.includes('~'))) return '';

  // Apply TPLAROYGBIVSUE 
  const ratingArr = morphArr.filter(item => ['T', 'P', 'L', 'A', 'R', 'O', 'Y', 'G', 'B', 'I', 'V', 'S', 'U', 'E'].includes(item));

  if (ratingArr.length > 0) {
    const ratingRank = { 'E': 15, 'U': 14, 'S': 13, 'V': 12, 'I': 11, 'B': 10, 'G': 9, 'Y': 8, 'O': 7, 'R': 6, 'A': 5, 'L': 4, 'P': 3, 'T': 2, '~': 1 };
    ratingArr.sort((a, b) => ratingRank[b] - ratingRank[a]);
  }

  const ratingMultArr = { 'T': 0.5, 'P': 0.7, 'L': 0.9, 'A': 1, 'R': 1.1, 'O': 1.2, 'Y': 1.3, 'G': 1.4, 'B': 1.5, 'I': 1.9, 'V': 2.3, 'S': 3.2, 'U': 4.1, 'E': 5 };
  const ratingPlusArr = { 'T': -3, 'P': -2, 'L': -1, 'A': 0, 'R': 1, 'O': 2, 'Y': 3, 'G': 4, 'B': 5, 'I': 6, 'V': 7, 'S': 8, 'U': 9, 'E': 10 };

  const ratingMult = ratingArr.length === 0 ? 1 : ratingMultArr[ratingArr[0]];
  
  const ratingPlus = ratingArr.length === 0 ? 0 : ratingArr.reduce((acc, item, idx) => {
    return acc + ratingPlusArr[item] / (1 << idx); // (1 << idx) is equivalent to 2^idx
  }, 0);

  const finalBase = Math.max(2,Math.round(base * ratingMult + ratingPlus));

  return finalBase;
} // END fKLcalcFinalSk





////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  My Abilities  (end Old MyKLs (menu))
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////












// fKLBuildMyAbilities //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Populates kl <MyAbilities> based on MyKLs and RougeAbilities and keeps oldKLMyAbilities info
function fKLBuildMyAbilities() {

  // Force Load KL <MyAbilities> Info
  let klMyAbilities = getObjKLMyAbilities(true); 
  
  // *** Reminder to -> Update cs <Game> Abilities names to match cs <List> *** 
  const reminderCB = klMyAbilities.arr[klMyAbilities.reminderCB_R][klMyAbilities.reminderCB_C];
  if (reminderCB) SpreadsheetApp.getUi().alert(`Character Sheet Reminder\n\nReminder: To see these changes, you will need to refresh the <Game> table on your Character Sheet.`);

  // ****** Update <AllAbilities> with CS <Game> possessions ******
  fGetPossessionsFromCS();
  // Get charLvl and save it to <MyAbilities>
  const charLvl = gCharLvl();
  gSaveVal('mykl','MyAbilities','Level','Level',charLvl);
  gSaveVal('mykl','MyAbilities','Level','Level2',charLvl);


  // ****** Assign a full copy of klMyAbilities to oldMyAbilities ******
  const oldKLMyAbilities = JSON.parse(JSON.stringify(klMyAbilities));

  // ****** Clear <MyAbilities> table *****
  gFillArraySection(klMyAbilities.arr, klMyAbilities.dataFirst_R, klMyAbilities.dataLast_R, klMyAbilities.first_C, klMyAbilities.last_C, '');
  gFillArraySection(klMyAbilities.arr, klMyAbilities.dataFirst_R, klMyAbilities.dataLast_R, klMyAbilities.known_C, klMyAbilities.learn_C, false);
  gSaveArraySectionToSheet(klMyAbilities.ref,klMyAbilities.arr,klMyAbilities.dataFirst_R,klMyAbilities.dataLast_R,klMyAbilities.first_C,klMyAbilities.last_C);

  // ***** Populate kl <MyAbilities> table with Known and Learnable kl <AllAbilities> *****
  fKLPopulateMyAbilitiesWithAllAbilities(klMyAbilities);

  // ***** Sort and Save <MyAbilities> table *****
  gSortArraySection(klMyAbilities.arr, klMyAbilities.dataFirst_R, klMyAbilities.dataLast_R, klMyAbilities.first_C, klMyAbilities.last_C, klMyAbilities.nameID_C);
  gSaveArraySectionToSheet(klMyAbilities.ref, klMyAbilities.arr, klMyAbilities.dataFirst_R, klMyAbilities.dataLast_R, klMyAbilities.first_C, klMyAbilities.last_C);

  // ***** Restore data from oldklMyAbilities *****
  fKLRestoreOldData(klMyAbilities,oldKLMyAbilities);

  // ***** Apply Training and Ver *****
  fKLApplyTrainingAndVer(klMyAbilities);
  fKLCalcMinLvl(klMyAbilities); // Do again after Train & Ver

  // **** Calculate AP ****
  fKLCalcAP(klMyAbilities);

  // *** Delete empty rows of <MyAbilities> ***
  gDeleteEmptyCol0DataRows('mykl','MyAbilities');

  // *** Save new MyAbiliites to CS <List>
  fKLUpdateCSAbilList(klMyAbilities);

} // End fKLBuildMyAbilities





// fGetPossessionsFromCS //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Update <AllAbilities> nonRCpossessions known & Learned to false, unless in CS <Game> Abilities/Gear table
function fGetPossessionsFromCS() {

  const gameObj = getObjCSGame();
  const dbAbilObj = getObjDBAbilities();
  const allAbilObj = getObjKLAllAbilities();

  // Build an array of allNonRCpossessionIDArr in DB <Abilities>
  const allNonRCpossessionIDArr = [];
  for (let r = dbAbilObj.dataFirst_R; r <= dbAbilObj.dataLast_R; r++) {
    if (dbAbilObj.arr[r][dbAbilObj.isNonRCPossession_C]) {
      allNonRCpossessionIDArr.push(dbAbilObj.arr[r][dbAbilObj.id_C]);
    }
  }

  // For each ID in CS <Game> Possessions table, if found in allNonRCpossessionIDArr, add to myNonRCpossessionIDArr
  const myNonRCpossessionIDArr = [];
  for (let r = gameObj.abilTblStart_R; r <= gameObj.gearTblEnd_R; r++) {
    const possessionID = gGetIDFromString(gameObj.arr[r][gameObj.abilNameID_C]);
    if (allNonRCpossessionIDArr.includes(possessionID)) {
      myNonRCpossessionIDArr.push(possessionID);
    }
  }

  // Update <AllAbilities> nonRCpossessions known & Learned to false, unless in CS <Game> Possessions (then both true)
  for (let r = allAbilObj.dataFirst_R; r <= allAbilObj.dataLast_R; r++) {
    const abilID = allAbilObj.arr[r][allAbilObj.id_C];
    if (allNonRCpossessionIDArr.includes(abilID)) {
      const isKnown = myNonRCpossessionIDArr.includes(abilID);
      allAbilObj.arr[r][allAbilObj.known_C] = isKnown;
      allAbilObj.arr[r][allAbilObj.learn_C] = isKnown;
    }
  }

  gSaveArraySectionToSheet(allAbilObj.ref, allAbilObj.arr, allAbilObj.dataFirst_R, allAbilObj.dataLast_R, allAbilObj.known_C, allAbilObj.learn_C);
  getObjKLAllAbilities(true);

} // End fGetPossessionsFromCS






// fKLPopulateMyAbilitiesWithAllAbilities //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Populate kl <MyAbilities> table with kl <AllAbilities>, for all eligible abilities
function fKLPopulateMyAbilitiesWithAllAbilities(klMyAbilities) {

  // Load KL <AllAbilities> Info
  let klAllAbilities = getObjKLAllAbilities();
  let dbAbilities = getObjDBAbilities();
  let myAbil_R = klMyAbilities.dataFirst_R;
  let needMoreMyAbilRows = 0;
  for (let r = klAllAbilities.dataFirst_R; r <= klAllAbilities.dataLast_R; r++) {
    const klAllAbilityID = klAllAbilities.arr[r][klAllAbilities.id_C];
    const cantLearn = klAllAbilities.arr[r][klAllAbilities.cantLearn_C];
    const isKnown = klAllAbilities.arr[r][klAllAbilities.known_C];
    const isLearnable = klAllAbilities.arr[r][klAllAbilities.learn_C];

    if (klAllAbilityID && !cantLearn && (isKnown || isLearnable)) {

      // Check if there is space in the array; if not, resize it
      if (myAbil_R >= klMyAbilities.arr.length) {
          klMyAbilities.arr.push(Array(klMyAbilities.arr[0].length).fill(''));  // Assuming all columns are initially empty
          needMoreMyAbilRows++;
      }

      let db_R = gKeyR('db','Abilities',klAllAbilityID);
      // Populate data
      klMyAbilities.arr[myAbil_R][klMyAbilities.id_C] = klAllAbilityID;
      klMyAbilities.arr[myAbil_R][klMyAbilities.known_C] = klAllAbilities.arr[r][klAllAbilities.known_C];
      klMyAbilities.arr[myAbil_R][klMyAbilities.learn_C] = false;
      klMyAbilities.arr[myAbil_R][klMyAbilities.nameID_C] = klAllAbilities.arr[r][klAllAbilities.nameID_C];
      if (gTestID('db','Abilities',klAllAbilityID)) klMyAbilities.arr[myAbil_R][klMyAbilities.notes_C] = dbAbilities.arr[db_R][dbAbilities.notes_C];
      klMyAbilities.arr[myAbil_R][klMyAbilities.source_C] = klAllAbilities.arr[r][klAllAbilities.source_C];
      klMyAbilities.arr[myAbil_R][klMyAbilities.parentKit_C] = dbAbilities.arr[db_R][dbAbilities.parentKit_C];
      klMyAbilities.arr[myAbil_R][klMyAbilities.parentMinVer_C] = dbAbilities.arr[db_R][dbAbilities.parentMinVer_C];
      klMyAbilities.arr[myAbil_R][klMyAbilities.base1_C] = klAllAbilities.arr[r][klAllAbilities.finalSk1_C];
      klMyAbilities.arr[myAbil_R][klMyAbilities.base2_C] = klAllAbilities.arr[r][klAllAbilities.finalSk2_C];
      myAbil_R++;
    }
  }

  // If needed, Expand <MyAbilities> rows to fit the data and copy first data rows formatting to all new rows
  if (needMoreMyAbilRows) {
    const sheet = klMyAbilities.ref;
    sheet.insertRowsAfter(sheet.getLastRow(), needMoreMyAbilRows);  
    const sourceRange = sheet.getRange(klMyAbilities.dataFirst_R + 1, 1, 1, sheet.getMaxColumns()); 
    sourceRange.copyFormatToRange(sheet, 1, sheet.getMaxColumns(), sheet.getMaxRows() - needMoreMyAbilRows + 1, sheet.getMaxRows());
  }


  // Save the new klMyAbilities.arr data
  gSaveArraySectionToSheet(klMyAbilities.ref,klMyAbilities.arr,klMyAbilities.dataFirst_R,myAbil_R - 1,klMyAbilities.first_C,klMyAbilities.last_C);

  // Reload <MyAbilities> table and Object
  klMyAbilities = getObjKLMyAbilities(true);


} // END fKLPopulateMyAbilitiesWithAllAbilities






// fKLRestoreOldData //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Restore oldKLMyAbilities Learned,Ver, and Training info
function fKLRestoreOldData(klMyAbilities,oldKLMyAbilities) {

  for (let r = oldKLMyAbilities.dataFirst_R; r <= oldKLMyAbilities.dataLast_R; r++) {
    const oldAbilID = oldKLMyAbilities.arr[r][oldKLMyAbilities.id_C];
    const abilID = gTestID('mykl', 'MyAbilities', oldAbilID, 'ID'); 
    if (abilID) {
      const myAbil_R = gKeyR('mykl','MyAbilities',abilID);
      klMyAbilities.arr[myAbil_R][klMyAbilities.learn_C] = oldKLMyAbilities.arr[r][oldKLMyAbilities.learn_C];
      klMyAbilities.arr[myAbil_R][klMyAbilities.ver_C] = oldKLMyAbilities.arr[r][oldKLMyAbilities.ver_C];
      klMyAbilities.arr[myAbil_R][klMyAbilities.train_C] = oldKLMyAbilities.arr[r][oldKLMyAbilities.train_C];
    }
  }

  gSaveArraySectionToSheet(klMyAbilities.ref,klMyAbilities.arr,klMyAbilities.dataFirst_R,klMyAbilities.dataLast_R,klMyAbilities.first_C,klMyAbilities.last_C);
  // Force reload
  klMyAbilities = getObjKLMyAbilities(true);

} // END fKLRestoreOldData




// fKLApplyTrainingAndVer //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> On kl <MyAbilities> calculates training and applies it to TrainedSk1 and TrainedSk2
function fKLApplyTrainingAndVer(klMyAbilities) {

  for (let r = klMyAbilities.dataFirst_R; r <= klMyAbilities.dataLast_R; r++) {
    const abilID = klMyAbilities.arr[r][klMyAbilities.id_C];

    if (!abilID) continue;

    fKLApplyVer(klMyAbilities,r,abilID);

    const base1 = klMyAbilities.arr[r][klMyAbilities.base1_C];
    const base2 = klMyAbilities.arr[r][klMyAbilities.base2_C];
    const parentKit = klMyAbilities.arr[r][klMyAbilities.parentKit_C];
    const hasParentKit = parentKit !== '';
    const canTrain = (!hasParentKit && (base1 || base2));

    let trainNum = (hasParentKit) ? fGetParentKitTraining(klMyAbilities,parentKit) : klMyAbilities.arr[r][klMyAbilities.train_C] || 0;
 
    // // Verify trainNum is either '' or an integer > 0
    // if (!(trainNum === '' || (Number.isInteger(trainNum) && trainNum >= 0))) {
    //   const abilityName = gSimpleName(klMyAbilities.arr[r][klMyAbilities.nameID_C]);
    //   throw new Error(`In fApplyTraining for the ability ${abilityName}, the Training must be a non-negative integer not "${trainNum}".\n\nPlease fix this and try again.`);
    // }

    let trainedSk1 = '';
    let trainedSk2 = '';
    if (base1) {
      const sk1DueToCharLvl = fKLCalcLog('charLvl', base1);
      const sk1DueToTrain = (trainNum && trainNum !== 0) ? fKLCalcLog('training', base1,trainNum) : 0;
      trainedSk1 = base1 + sk1DueToCharLvl + sk1DueToTrain;
    }

    if (base2) {
      const sk2DueToCharLvl = fKLCalcLog('charLvl', base2);
      const sk2DueToTrain = (trainNum && trainNum !== 0) ? fKLCalcLog('training', base2,trainNum) : 0;
      trainedSk2 = base2 + sk2DueToCharLvl + sk2DueToTrain;
    } 

    if (!canTrain || trainNum === 0) klMyAbilities.arr[r][klMyAbilities.train_C] = '';
    klMyAbilities.arr[r][klMyAbilities.kitTrain_C] = (hasParentKit && trainNum !== 0) ? trainNum : '';

    klMyAbilities.arr[r][klMyAbilities.trainedSk1_C] = trainedSk1;
    klMyAbilities.arr[r][klMyAbilities.trainedSk2_C] = trainedSk2;

    // force Learned if Ver > 1 or training > 0
    let known = klMyAbilities.arr[r][klMyAbilities.known_C];
    let learned = klMyAbilities.arr[r][klMyAbilities.learn_C];

    let ver = parseInt(klMyAbilities.arr[r][klMyAbilities.ver_C]) || 1;
    const train = parseInt(klMyAbilities.arr[r][klMyAbilities.train_C]) || 0;

    // Set Learned
    klMyAbilities.arr[r][klMyAbilities.learn_C] = (known) ? true : learned;

    // Set Training
    klMyAbilities.arr[r][klMyAbilities.train_C] = (learned) ? klMyAbilities.arr[r][klMyAbilities.train_C] : '';

    // Set Ver
    klMyAbilities.arr[r][klMyAbilities.ver_C] = (learned) ? klMyAbilities.arr[r][klMyAbilities.ver_C] || 1 : '';
  }

  gSaveArraySectionToSheet(klMyAbilities.ref, klMyAbilities.arr, klMyAbilities.dataFirst_R, klMyAbilities.dataLast_R, klMyAbilities.first_C, klMyAbilities.last_C);
  // Force reload
  klMyAbilities = getObjKLMyAbilities(true);
} // END fKLApplyTrainingAndVer




// fKLApplyVer //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> On kl <MyAbilities> applies Ver to vMax, Act, Dur, Rng, Meta, Uses, Regain
// Purpose: Helper of fApplyTrainingAndVer
function fKLApplyVer(klMyAbilities, r, abilID) {
  let dbVer = getObjDBVersions();

  const ver = parseInt(klMyAbilities.arr[r][klMyAbilities.ver_C]) || 1;

  if (!Number.isInteger(ver) || ver < 0) throw new Error(`In fKLApplyVer an illegal Ver of "${ver}" was loaded from <MyAbilities>`);

  // Find Max Ver allowed
  let maxVer = 0;
  let testVer = 1;
  while (gKeyR('db', 'Versions', `${abilID}.v${testVer}`)) {
      maxVer = testVer++;
  }

  let newVer = (ver <= maxVer) ? ver : maxVer;
  let db_R = gKeyR('db','Versions',`${abilID}.v${newVer}`);

  // In case newVer started at 1, give .v1 a try so initial Act/Dur/Rng/Meta/uses/Regain can be seen. Do NOT reset newVer, however.
  if (newVer === 1) {
    dbVerKey = `${abilID}.v1`;
    db_R = gKeyR('db','Versions',dbVerKey);
  }

  // Case where there is no id.v# in db
  if (!db_R) {
    klMyAbilities.arr[r][klMyAbilities.ver_C] = 1;
    klMyAbilities.arr[r][klMyAbilities.maxVer_C] = 1;
    gFillArraySection(klMyAbilities.arr, r, r, klMyAbilities.act_C, klMyAbilities.meta_C, '~');
    gFillArraySection(klMyAbilities.arr, r, r, klMyAbilities.uses_C, klMyAbilities.regain_C, '');
  } else {
    klMyAbilities.arr[r][klMyAbilities.ver_C] = newVer;
    klMyAbilities.arr[r][klMyAbilities.maxVer_C] = maxVer;
    klMyAbilities.arr[r][klMyAbilities.act_C] = dbVer.arr[db_R][dbVer.act_C];
    klMyAbilities.arr[r][klMyAbilities.dur_C] = dbVer.arr[db_R][dbVer.dur_C];
    klMyAbilities.arr[r][klMyAbilities.rng_C] = dbVer.arr[db_R][dbVer.rng_C];
    klMyAbilities.arr[r][klMyAbilities.meta_C] = dbVer.arr[db_R][dbVer.meta_C];
    klMyAbilities.arr[r][klMyAbilities.uses_C] = dbVer.arr[db_R][dbVer.uses_C];
    klMyAbilities.arr[r][klMyAbilities.regain_C] = dbVer.arr[db_R][dbVer.regain_C];
  }

} // END fKLApplyVer




// fGetParentKitTraining //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Gets the trainingNum from the parent Kit
function fGetParentKitTraining(klMyAbilities,parentKit) {

  const parentKey = gGetIDFromString(parentKit);
  const parent_R = gKeyR('mykl','MyAbilities',parentKey)

  return klMyAbilities.arr[parent_R][klMyAbilities.train_C]

} // END fGetParentKitTraining




// fKLCalcMinLvl //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> On kl <MyAbilities> calculates Min Lvl needed to learn and for next Ver (note min level to Train = minLvlToLearn)
function fKLCalcMinLvl(klMyAbilities) {
  const charLvl = klMyAbilities.arr[klMyAbilities.charLevel_R][klMyAbilities.charLevel_C];

  for (let r = klMyAbilities.dataFirst_R; r <= klMyAbilities.dataLast_R; r++) {
    const abilID = klMyAbilities.arr[r][klMyAbilities.id_C];
    if (!abilID) continue;
  
    // Lesser constant calculations
    const isKnown = klMyAbilities.arr[r][klMyAbilities.known_C];
    const parentID = gGetIDFromString(klMyAbilities.arr[r][klMyAbilities.parentKit_C]);
    const hasParentKit = !!parentID;
    const isParentLearned = hasParentKit && fIsParentLearned(klMyAbilities, parentID);
    const maxVer = klMyAbilities.arr[r][klMyAbilities.maxVer_C] || 1; // Default to 1 if empty
    const currentVer = klMyAbilities.arr[r][klMyAbilities.ver_C] || 1; // Default to 1 if empty
    const myBestVer = Math.min(Math.trunc(Math.sqrt(charLvl)), maxVer);
    const myVer = Math.min(currentVer, myBestVer);
    const minlvlForNextMyVer = (myVer === maxVer) ? 'Maxed' : (myVer + 1) ** 2;
    const minKitVerToLearn = hasParentKit ? klMyAbilities.arr[r][klMyAbilities.parentMinVer_C] || 0 : 0;
    const parentKitVer = (hasParentKit) ? fGetParentKitVer(klMyAbilities, parentID) : '';
    const minKitLvlToLearn = minKitVerToLearn ? minKitVerToLearn ** 2 : 1;
    const trainNum = klMyAbilities.arr[r][klMyAbilities.train_C];
    let minLvlToLearn, isLearned, minLvlToVer

    // Set minLvlToLearn
    if (hasParentKit) {
      if (isParentLearned) {
          if (parentKitVer >= minKitVerToLearn) {
            minLvlToLearn = minKitLvlToLearn;
          } else minLvlToLearn = `Kit v${minKitVerToLearn}`;
      } else minLvlToLearn = 'LrnKit';
    } else minLvlToLearn = 1;

    // Set isLearned
    if (isKnown) {
        isLearned = (Number.isInteger(minLvlToLearn) && charLvl >= minLvlToLearn) ? true : false;
    } else if (Number.isInteger(minLvlToLearn) && charLvl >= minLvlToLearn) {
        isLearned = klMyAbilities.arr[r][klMyAbilities.learn_C];
    } else isLearned = false;

    
    // Set minLvlToVer
    if (isLearned) {
      if (hasParentKit) {
        minLvlToVer = isParentLearned ? ((minlvlForNextMyVer === 'Maxed') ? 'Maxed' : minKitLvlToLearn + minlvlForNextMyVer - 1): 'LrnKit';
      } else minLvlToVer = (minlvlForNextMyVer === 'Maxed') ? 'Maxed' : minlvlForNextMyVer;
    } else minLvlToVer = '<- Learn';

    // Set minLvlToLearn



    // Set all the ability row's data
    klMyAbilities.arr[r][klMyAbilities.minLearnLvl_C] = minLvlToLearn;
    klMyAbilities.arr[r][klMyAbilities.minVerLvl_C] = minLvlToVer;
    klMyAbilities.arr[r][klMyAbilities.learn_C] = isLearned;
    klMyAbilities.arr[r][klMyAbilities.train_C] = (hasParentKit) ? 'by Kit ->' : ((!isLearned) ? '' : trainNum);
    klMyAbilities.arr[r][klMyAbilities.ver_C] = (isLearned) ? myVer : 1;
  }

  gSaveArraySectionToSheet(klMyAbilities.ref, klMyAbilities.arr, klMyAbilities.dataFirst_R, klMyAbilities.dataLast_R, klMyAbilities.first_C, klMyAbilities.last_C);
  // Force reload
  klMyAbilities = getObjKLMyAbilities(true);

} // END fKLCalcMinLvl




// fIsParentLearned //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Returns true if parentKit is Known or Learned
function fIsParentLearned(klMyAbilities, parentID) {

const parent_R = gKeyR('mykl','MyAbilities',parentID,'ID');
const isLearned = klMyAbilities.arr[parent_R][klMyAbilities.learn_C];

return isLearned;

} // END fIsParentLearned





// fGetParentKitVer //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Returns the current version of the parent kit
function fGetParentKitVer(klMyAbilities, parentID) {

const parent_R = gKeyR('mykl','MyAbilities',parentID,'ID');
const parentKitVer = klMyAbilities.arr[parent_R][klMyAbilities.ver_C];

return parentKitVer;

} // END fGetParentKitVer






// fKLCalcAP //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Calculates all AP costs on kl <MyAbilities>
function fKLCalcAP(klMyAbilities) {
  
  let learnedCost = 0, verCost = 0, trainingCost = 0;

  for (let r = klMyAbilities.dataFirst_R; r <= klMyAbilities.dataLast_R; r++) {
    const known = klMyAbilities.arr[r][klMyAbilities.known_C];
    const learned = klMyAbilities.arr[r][klMyAbilities.learn_C];
    let ver = parseInt(klMyAbilities.arr[r][klMyAbilities.ver_C]) || 1;
    let trainNum = parseInt(klMyAbilities.arr[r][klMyAbilities.train_C]) || 0;

    learnedCost += (!known && learned) ? 5 : 0;
    verCost += fKLVerCost(ver);
    trainingCost += trainNum;

    // Per ability Calc Total abilAPCost 
    const totOfthisAbilAPCost = fCalcTotalOfThisAbilAPCost(known,learned,klMyAbilities.arr[r][klMyAbilities.train_C],klMyAbilities.arr[r][klMyAbilities.ver_C]);
    klMyAbilities.arr[r][klMyAbilities.abilAPCost_C] = (totOfthisAbilAPCost > 0) ? totOfthisAbilAPCost : '';
  }

  const gearAP = parseInt(gGetVal('mycs','Game','PossAPTot','PossGrandAPTot')) || 0;
  const levelAP = 10 + ( 10 * fKLGetCharLevel());
  const bonusAP = parseInt(klMyAbilities.arr[klMyAbilities.apData2_R][klMyAbilities.apCount_C]) || 0;
  const apTotal = levelAP + bonusAP;
  const apTotalSpent = gearAP + learnedCost + verCost + trainingCost;
  const remainingAP = apTotal - apTotalSpent;

  // Calc AP Count
  klMyAbilities.arr[klMyAbilities.apData1_R][klMyAbilities.apCount_C] = levelAP;
  klMyAbilities.arr[klMyAbilities.apData3_R][klMyAbilities.apCount_C] = apTotal;

  // Calculate AP Spent
  klMyAbilities.arr[klMyAbilities.gearAP_R][klMyAbilities.apSpent_C] = gearAP;
  klMyAbilities.arr[klMyAbilities.apData1_R][klMyAbilities.apSpent_C] = learnedCost;
  klMyAbilities.arr[klMyAbilities.apData2_R][klMyAbilities.apSpent_C] = verCost;
  klMyAbilities.arr[klMyAbilities.apData3_R][klMyAbilities.apSpent_C] = trainingCost;

  // Calculate AP Total
  klMyAbilities.arr[klMyAbilities.apData1_R][klMyAbilities.apTotal_C] = apTotal;
  klMyAbilities.arr[klMyAbilities.apData2_R][klMyAbilities.apTotal_C] = apTotalSpent;
  klMyAbilities.arr[klMyAbilities.apData3_R][klMyAbilities.apTotal_C] = remainingAP;

  // Update the spreadsheet with calculated values
  gSaveArraySectionToSheet(klMyAbilities.ref, klMyAbilities.arr, klMyAbilities.gearAP_R, klMyAbilities.apData3_R, klMyAbilities.apCount_C, klMyAbilities.apTotal_C);
  gSaveArraySectionToSheet(klMyAbilities.ref, klMyAbilities.arr, klMyAbilities.dataFirst_R, klMyAbilities.dataLast_R, klMyAbilities.abilAPCost_C, klMyAbilities.abilAPCost_C);

  // Alert the user if AP is overspent
  if (remainingAP < 0) {
    let ui = SpreadsheetApp.getUi(); // getUi() part of the Google Apps Script for spreadsheets
    ui.alert('WARNING', `You have spent ${-remainingAP} more AP than you actually have. Please correct and try again.`, ui.ButtonSet.OK);
  }

} // END fKLCalcAP




// fCalcTotalOfThisAbilAPCost //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Calculates the Total AP cost of one ability on kl <MyAbilities>
function fCalcTotalOfThisAbilAPCost(known, learned, trainingAP, ver) {
    let totAP = (!known && learned) ? 5 : 0;
    
    // Ensure trainingAP is treated as a number and valid integer
    const trainingCost = Number.isInteger(trainingAP) ? trainingAP : 0;
    
    // Add the total AP cost from training and version cost
    totAP += trainingCost + fKLVerCost(ver);

    return totAP;
} // END fCalcTotalOfThisAbilAPCost





// fKLVerCost //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Calculates the total verCost based on the value of ver
// Purpose: If ver < 2, verCost = 0, otherwise calculate the cumulative of Ver^2 + (Ver-1)^2, etc. down to Ver = 2
function fKLVerCost(ver) {
    if (ver < 2) {
        return 0;
    } else {
        // Calculate cumulative sum starting from 2 up to ver
        let total = (ver >= 2) ? 5 : 0; // if Ver = 1, then AP 0, if Ver = 2 then AP 5
        for (let i = 3; i <= ver; i++) {
            total += i**2;
        }
        return total;
    }
} // END fKLVerCost



// fKLUpdateCSAbilList //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Updates CS <List> abilities to match <MyAbilities> plus Artf: and Gr: 
function fKLUpdateCSAbilList() {

  let klMyAbil = getObjKLMyAbilities();
  let csList = getObjCSList();
  // Clear CS <List> section for Known & Learned MyAbilities
  gFillArraySection(csList.arr, csList.dataFirst_R, csList.dataLast_R, csList.abilityNameID_C, csList.abilityNameID_C, '');

  let csList_R = csList.dataFirst_R;
  const csLast_R = csList.dataLast_R;
  let needMoreCSRows = 0;

  // Populate csList with valid MyAbilities
  for (let r = klMyAbil.dataFirst_R; r <= klMyAbil.dataLast_R; r++) {
    const myAbilID = klMyAbil.arr[r][klMyAbil.id_C];
    const learned = klMyAbil.arr[r][klMyAbil.learn_C];

    if (myAbilID && learned) {
      if (csList_R > csLast_R) {
        needMoreCSRows++;
        csList.arr.push(new Array(csList.arr[0].length).fill(''));  // Ensuring all columns are initially empty
      }
      csList.arr[csList_R][csList.abilityNameID_C] = klMyAbil.arr[r][klMyAbil.nameID_C];
      csList_R++;
    }
  }

  // Popuate csList with all Artf: and Eq: gear from DB
  let dbGearObj = getObjDBGear();
  for (let r = dbGearObj.dataFirst_R; r <= dbGearObj.dataLast_R; r++) {
    const nameID = dbGearObj.arr[r][dbGearObj.nameID_C];
    const isArtfEq = nameID.startsWith('Artf:') || nameID.startsWith('Eq:');

    if (isArtfEq) {
      if (csList_R > csLast_R) {
        needMoreCSRows++;
        csList.arr.push(new Array(csList.arr[0].length).fill(''));  // Ensuring all columns are initially empty
      }
      csList.arr[csList_R][csList.abilityNameID_C] = nameID;
      csList_R++;
    }
  }



  // Add Sheet rows and copy formatting if needed
  const sheet = csList.ref;
  if (needMoreCSRows > 0) {
    sheet.insertRowsAfter(sheet.getLastRow(), needMoreCSRows);
    const sourceRange = sheet.getRange(csList.dataFirst_R + 1, 1, 1, sheet.getMaxColumns());
    sourceRange.copyFormatToRange(sheet, 1, sheet.getMaxColumns(), sheet.getMaxRows() - needMoreCSRows + 1, sheet.getMaxRows());
  }

  // Save the updated CS List back to the sheet
  gSaveArraySectionToSheet(csList.ref, csList.arr, csList.dataFirst_R, sheet.getMaxRows() - 1, csList.abilityNameID_C, csList.abilityNameID_C);

  // Reload CS List
  csList = getObjCSList(true);
} // END fKLUpdateCSAbilList





////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  All Abilities  (My Abilities)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////





// fKLLoadAbilitiesFromDB //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Populates kl <AllAbilities> with all db <AllAbilities> but keeps oldKLAllAbilities calculated info
function fKLLoadAbilitiesFromDB() {
  let dbAbilities = getObjDBAbilities();
  let klAllAbilities = getObjKLAllAbilities();

  // Assign a full copy of klAllAbilities to oldKLAllAbilities
  const oldKLAllAbilities = JSON.parse(JSON.stringify(klAllAbilities));

  // ***** Size and Clear <AllAbilities> table *****
  klAllAbilities = fKLEqualizeNumberOfDataRowsFromTo(dbAbilities, klAllAbilities);
  gFillArraySection(klAllAbilities.arr, klAllAbilities.dataFirst_R, klAllAbilities.dataLast_R, klAllAbilities.first_C, klAllAbilities.last_C, '');
  gFillArraySection(klAllAbilities.arr, klAllAbilities.dataFirst_R, klAllAbilities.dataLast_R, klAllAbilities.cantLearn_C, klAllAbilities.learn_C, false);
  gSaveArraySectionToSheet(klAllAbilities.ref, klAllAbilities.arr, klAllAbilities.dataFirst_R, klAllAbilities.dataLast_R, klAllAbilities.first_C, klAllAbilities.last_C);

  // ***** Populate kl <AllAbilities> table with db <Abilities> *****
  fKLPopulateKlAllAbilities(klAllAbilities, dbAbilities);


  // ***** Restore data from oldKLAllAbilities *****
  fKLRestoreOldKlAbilitiesData(klAllAbilities, oldKLAllAbilities);

  // ***** Sort and Save <AllAbilities> table *****
  gSortArraySection(klAllAbilities.arr, klAllAbilities.dataFirst_R, klAllAbilities.dataLast_R, klAllAbilities.first_C, klAllAbilities.last_C, klAllAbilities.nameID_C);
  gSaveArraySectionToSheet(klAllAbilities.ref, klAllAbilities.arr, klAllAbilities.dataFirst_R, klAllAbilities.dataLast_R, klAllAbilities.first_C, klAllAbilities.last_C);
} // End fKLLoadAbilitiesFromDB




// fKLEqualizeNumberOfDataRowsFromTo //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Makes sure klAllAbilities number of data rows matches the number of data rows in dbAbilities
function fKLEqualizeNumberOfDataRowsFromTo(dbAbilities, klAllAbilities) {
  const dbAbilitiesNumOfDataRows = dbAbilities.dataLast_R - dbAbilities.dataFirst_R + 1;
  const klAllAbilitiesNumOfDataRows = klAllAbilities.dataLast_R - klAllAbilities.dataFirst_R + 1;

  // Add rows if needed
  if (dbAbilitiesNumOfDataRows > klAllAbilitiesNumOfDataRows) {
    const addNumRows = dbAbilitiesNumOfDataRows - klAllAbilitiesNumOfDataRows;

    // Add rows to the bottom of the sheet referenced by klAllAbilities.ref
    klAllAbilities.ref.insertRowsAfter(klAllAbilities.dataLast_R, addNumRows);
    klAllAbilities = getObjKLAllAbilities(true);

    // Delete rows if needed
  } else if (dbAbilitiesNumOfDataRows < klAllAbilitiesNumOfDataRows) {
    const deleteNumRows = klAllAbilitiesNumOfDataRows - dbAbilitiesNumOfDataRows;

    // Delete rows from the bottom of the sheet referenced by klAllAbilities.ref
    klAllAbilities.ref.deleteRows(klAllAbilities.dataLast_R - deleteNumRows + 1, deleteNumRows);
    klAllAbilities = getObjKLAllAbilities(true);
  }

  return klAllAbilities;
} // End fKLEqualizeNumberOfDataRowsFromTo





// fKLPopulateKlAllAbilities //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Populate kl <AllAbilities> table with db <Abilities> values for ID, NameID, Base1 and Base2
function fKLPopulateKlAllAbilities(klAllAbilities, dbAbilities) {
  let klAbil_R = klAllAbilities.dataFirst_R;

  for (let r = dbAbilities.dataFirst_R; r <= dbAbilities.dataLast_R; r++) {
      const dbAbilityID = dbAbilities.arr[r][dbAbilities.id_C];

      klAllAbilities.arr[klAbil_R][klAllAbilities.id_C] = dbAbilityID;
      klAllAbilities.arr[klAbil_R][klAllAbilities.nameID_C] = dbAbilities.arr[r][dbAbilities.nameID_C];
      klAllAbilities.arr[klAbil_R][klAllAbilities.base1_C] = dbAbilities.arr[r][dbAbilities.base1_C];
      klAllAbilities.arr[klAbil_R][klAllAbilities.base2_C] = dbAbilities.arr[r][dbAbilities.base2_C];
      klAbil_R++;
  }

} // END fKLPopulateKlAllAbilities





// fKLRestoreOldKlAbilitiesData //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Restore data from oldKLAllAbilities to klAllAbilities based on matching IDs
function fKLRestoreOldKlAbilitiesData(klAllAbilities, oldKLAllAbilities) {
    for (let r = oldKLAllAbilities.dataFirst_R; r <= oldKLAllAbilities.dataLast_R; r++) {
        const oldAbilID = oldKLAllAbilities.arr[r][oldKLAllAbilities.id_C];
        // Find the index of oldAbilID in the new klAllAbilities array
        const foundIndex = klAllAbilities.arr.findIndex(row => row[klAllAbilities.id_C] === oldAbilID);
        if (foundIndex !== -1) {
            klAllAbilities.arr[foundIndex][klAllAbilities.source_C] = oldKLAllAbilities.arr[r][oldKLAllAbilities.source_C];
            klAllAbilities.arr[foundIndex][klAllAbilities.cantLearn_C] = oldKLAllAbilities.arr[r][oldKLAllAbilities.cantLearn_C];
            klAllAbilities.arr[foundIndex][klAllAbilities.known_C] = oldKLAllAbilities.arr[r][oldKLAllAbilities.known_C];
            klAllAbilities.arr[foundIndex][klAllAbilities.learn_C] = oldKLAllAbilities.arr[r][oldKLAllAbilities.learn_C];
            klAllAbilities.arr[foundIndex][klAllAbilities.morphOther_C] = oldKLAllAbilities.arr[r][oldKLAllAbilities.morphOther_C];
            klAllAbilities.arr[foundIndex][klAllAbilities.sk1PLAGHE_C] = oldKLAllAbilities.arr[r][oldKLAllAbilities.sk1PLAGHE_C];
            klAllAbilities.arr[foundIndex][klAllAbilities.sk2PLAGHE_C] = oldKLAllAbilities.arr[r][oldKLAllAbilities.sk2PLAGHE_C];
            klAllAbilities.arr[foundIndex][klAllAbilities.finalSk1_C] = oldKLAllAbilities.arr[r][oldKLAllAbilities.finalSk1_C];
            klAllAbilities.arr[foundIndex][klAllAbilities.finalSk2_C] = oldKLAllAbilities.arr[r][oldKLAllAbilities.finalSk2_C];
        }
    }
} // END fKLRestoreOldKlAbilitiesData





////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  RogueAbilities  (end All Abilities)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////






// fKLRefreshRogueAbil //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Updates and verifies selected Rogue Abilities and applies notes to kl <RogueAbilities>
function fKLRefreshRogueAbil() {

  // Load KL <RogueAbilities> Info
  let klRogueAbilities = getObjKLRogueAbilities();

  // Remove non-existing abilities, clean ID, nameID, and Notes of existing abilities
  for (let r = klRogueAbilities.dataFirst_R; r <= klRogueAbilities.dataLast_R; r++) {
    const nameID = klRogueAbilities.arr[r][klRogueAbilities.nameID_C];
    const id = gTestID('mykl','AllAbilities',nameID);
    if (id) {
      klRogueAbilities.arr[r][klRogueAbilities.id_C] = gGetVal('mykl','AllAbilities',id,'ID');
      klRogueAbilities.arr[r][klRogueAbilities.nameID_C] = gGetVal('mykl','AllAbilities',id,'Name_ID');
      klRogueAbilities.arr[r][klRogueAbilities.notes_C] = gGetVal('db','Abilities',id,'Notes');
    } else {
      gFillArraySection(klRogueAbilities.arr, r, r, klRogueAbilities.first_C, klRogueAbilities.last_C,'');
      klRogueAbilities.arr[r][klRogueAbilities.known_C] = false;
    }
  }

  // Check for duplicates
  const nameIDsSeen = {};
  let duplicateFound = false;
  for (let r = klRogueAbilities.dataFirst_R; r <= klRogueAbilities.dataLast_R; r++) {
    const nameID = klRogueAbilities.arr[r][klRogueAbilities.nameID_C];
    if (nameID === '') continue;
    if (nameIDsSeen[nameID]) {
      if (!duplicateFound) {
        duplicateFound = true;
        var ui = SpreadsheetApp.getUi(); // Gets the user interface object for a Google Sheet
        ui.alert('Duplicate Abilities Found.', 'All duplicates will be deleted.', ui.ButtonSet.OK);
      }
      gFillArraySection(klRogueAbilities.arr, r, r, klRogueAbilities.first_C, klRogueAbilities.last_C,'');
      klRogueAbilities.arr[r][klRogueAbilities.known_C] = false;
    }
    nameIDsSeen[nameID] = true;
  }

  // Sort and Save <RogueAbilities> table
  gSortArraySection(klRogueAbilities.arr, klRogueAbilities.dataFirst_R, klRogueAbilities.dataLast_R, klRogueAbilities.first_C, klRogueAbilities.last_C,klRogueAbilities.nameID_C);
  gSaveArraySectionToSheet(klRogueAbilities.ref, klRogueAbilities.arr, klRogueAbilities.dataFirst_R, klRogueAbilities.dataLast_R, klRogueAbilities.first_C, klRogueAbilities.last_C);

  // Force load the changed data
  klRogueAbilities = getObjKLRogueAbilities(true);

  // **** Clear all but (Key, NameID, Base1, Base2) of AllAbilities, then ... Populate AllAbilities based on MyKLs and Rouge
  fKLKitsAndRougeToAllAbilities();

  // **** Build MyAbilities ****
  fKLBuildMyAbilities();

} // END fKLRefreshRogueAbil


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  Math  (end RogueAbilities)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


// fKLGetCharLevel //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Gets the character level from CS and verifies it is a legitimate integer > 0
function fKLGetCharLevel() {
  let raceClass = getObjCSRaceClass();
  const level = raceClass.arr[raceClass.charLevel_R][raceClass.value_C];
  if (!Number.isInteger(level) || level <= 0) {
    throw new Error(`Your character's level on your <RaceClass> tab is "${level}", which is not a valid integer`);
  }

  return level;

} // End fKLGetCharLevel



// fKLCalcLog //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Calculates the logarithmic improvement value based on type and some logLevel (could be training level of charlevel)
// Purpose: a = rootnum of curve which is 0 for training and skill leveling
// Purpose: b = Steepness of curve
// Purpose: c = inverse growth (smaller the greater)
// Purpose: logLevel = Charlevel for Sks, or trainingLevel for training
// Purpose: d = early jump to logLevel
// Purpose: e = scale adjustment ("kind of" like an adj to base, but not entirely)
// Purpose: with a,b,c,d,e of 10,7.6,1.6,20,-49 the formula is 10+7.6*log1.6(level+20)-49  where log1.6 means log base 1.6
// Purpose: To create more or see the curve, use the Sheet Notes <Logs> tab
function fKLCalcLog(logType, base, logLevel = fKLGetCharLevel()) {
    // Verify base is an integer > -1 else throw an error

    if (!Number.isInteger(base) || logLevel < 0) {
        throw new Error(`Base should be a positive integer but fCalcLog was passed a level of "${base}".\n\nPlease fix this and try again.`);
    }

    // Verify logLevel is an integer > 0 else set to 0
    logLevel = (!Number.isInteger(logLevel) || logLevel <= 0) ? 0 : logLevel;

    const skLvlArr = [0, 7.6, 1.6, 20, -49];
    const trainingArr = [0, 44, 11, 30, -62];

    let a, b, c, d, e;
    switch (logType) {
        case 'charLvl': [a, b, c, d, e] = skLvlArr; break;
        case 'training': [a, b, c, d, e] = trainingArr; break;
        default: 
          throw new Error(`fCalcLog was passed an unknown logType of "${logType}"`);
    }
    let result = a + b * (Math.log(logLevel + d) / Math.log(c)) + e;
    // Normalize to PlAGHE so that base of 2 is mult of about .3, base of 10 is mult of 1, base of 60 is mult of 1.78
    const newMult = Math.log2(base) / Math.log2(10);
    return Math.round(result * newMult);
} // End fKLCalcLog






////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  DESIGNER  (end Math)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////















