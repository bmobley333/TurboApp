// KL

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  Menu (end initialize)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



// fKLCreateMenu //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Create MetaScape menu
function fKLCreateMenu() {
  SpreadsheetApp.getUi()
  .createMenu('*** MyKLs')
    .addItem('Calc KLs & AP', 'fKLMenuCalcKLAndAP')
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
          .addItem('Validate AP Costs', 'fKLMenuSetKLAPCosts')
          .addItem('Set Notes', 'fKLMenuSetKLNotes')
          .addItem('Hide Other RC Tabs', 'fKLMenuHideOtherRCTabs')
          .addItem('Un-Hide All RC Tabs', 'fKLMenuUn_HideAllRCTabs')
        )
      .addToUi();
  } 
} // End fKLCreateMenu



// Menu Functions //////////////////////////////////////////////////////////////////////////////////////////////////
// Kit Menu
function fKLMenuCalcKLAndAP() {fKLRunMenuOrButton('CalcKLAndAP');}
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
      case 'CalcKLAndAP': fKLCalcKLAndAP(); break;
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
 * Assumptions: The KL sheets are correctly formatted.
 * Notes: This function orchestrates the entire AP calculation and sheet update process.
 * @returns {void}
 */
function fKLCalcKLAndAP() {

    // Get the character's active RC tab name, hiding other RC tabs in the process.
    const myRCTabName = fKLGetRCAndHideOtherRCTabs();
    const myUsedRCTabs = ['All', myRCTabName];

    // Verify RC Checkboxes and set KL Card Colors.
    const tierString = gGetVal('mykl', 'All', 'Tier', 'Tier');
    const tierMatch = String(tierString).match(/\d+/);
    const tierNum = tierMatch ? parseInt(tierMatch[0], 10) : 0;
    fKLVerifyRCCheckedBoxesSetColors(myUsedRCTabs,tierNum);

    // Calculate AP Spent from all checked abilities.
    const apSpent = fKLAPSpentCalc(myUsedRCTabs);

    // Build and save final header info based on the calculated spent AP.
    const headerInfo = fKLBuildHeaderObject(apSpent);
    fKLSaveRCHeaderInfo(myUsedRCTabs, headerInfo);
    fKLAlertIfOverspentAP(headerInfo);

    // Build array of known KLs and their best buff/version numbers.
    const knownKLs = fKLBuildKnownKLs(myUsedRCTabs);
    const extractedKLs = fKLExtractKLsFromKLGroups(knownKLs);

    // Build the KnownAbilities sheet from the extracted KLs.
    fKLBuildKnownAbilitiesSheet(extractedKLs);

    // Copy Known Ability list to CS <List>
    fKLCopyKnownAbilitiesToCS();

    SpreadsheetApp.getUi().alert('Character Sheet Reminder', 'Reminder: To see these changes, you will need to refresh the <Game> table on your Character Sheet.', SpreadsheetApp.getUi().ButtonSet.OK);

} // End fKLCalcKLAndAP




/**
 * Purpose: Calculates the total spent AP from a given list of KL sheets by summing costs from checked abilities.
 * Assumptions: The KL sheets are correctly formatted.
 * Notes: This function iterates through the specified sheets to calculate a single AP total.
 * @param {string[]} myRCTabName - An array of KL sheet names to process.
 * @returns {number} The total AP spent across all specified sheets.
 */
function fKLAPSpentCalc(myRCTabName) {
    let totalAPSpent = 0;

    // Iterate through each provided sheet name.
    myRCTabName.forEach(tabName => {
        const currentTab = getObjKL_RCTab(tabName, true);
        const dataFirstTF_C = 1;
        const dataLastTF_C = currentTab.arr[0].length - 1;

        for (let r = currentTab.dataFirst_R; r <= currentTab.dataLast_R; r++) {
            for (let c = dataFirstTF_C; c < dataLastTF_C; c++) {
                // If a checkbox in the current cell is checked, process the adjacent card.
                if (currentTab.arr[r][c] === true) {
                    const cardText = currentTab.arr[r][c + 1];
                    if (!cardText) continue;

                    const klCard = fGetKLCardObj(cardText, tabName, r, c + 1);

                    // Add the card's AP cost to the total unless it's free.
                    if (!klCard.isFree) {
                        totalAPSpent += klCard.apCost;
                    }
                }
            }
        }
    });

    return totalAPSpent;
} // End fKLAPSpentCalc




/**
 * Purpose: Calculates all header AP, Level, and Tier values and assembles them into a single object.
 * Assumptions: The `apSpent` value is a valid number.
 * Notes: This function centralizes all header calculations.
 * @param {number} apSpent - The total calculated AP spent.
 * @returns {object} An object containing all calculated header values.
 */
function fKLBuildHeaderObject(apSpent) {
    // First, calculate all the necessary values in sequence.
    const myLevel = gCharLvl();
    const levelAP = 10 * myLevel + 10;
    const tierString = gGetVal('mykl', 'All', 'Tier', 'Tier');
    const tierMatch = String(tierString).match(/\d+/);
    const tierNum = tierMatch ? parseInt(tierMatch[0], 10) : 0;
    const rcName_ID = gGetVal('mykl', 'All', 'RC', 'RC');
    const bnsAP = gGetVal('mykl', 'All', 'BnsAP', 'APVal');
    const totalAP = levelAP + bnsAP;
    const apRemaining = totalAP - apSpent;

    // Then, assemble the final object using the calculated constants.
    const header = {
        myLevel,
        tierString,
        tierNum,
        rcName_ID,
        bnsAP,
        totalAP,
        apSpent,
        apRemaining,
    };

    return header;
} // End fKLBuildHeaderObject




/**
 * Purpose: Saves the calculated header AP, Level, and Tier info to a given list of RC sheets.
 * Assumptions: The header object 'h' has been pre-calculated by fKLBuildHeaderObject().
 * Notes: This function updates the header section of multiple sheets with consistent data.
 * @param {string[]} myUsedRCTabs - An array of KL sheet names to update.
 * @param {object} h - The header object containing the values to be saved.
 * @returns {void}
 */
function fKLSaveRCHeaderInfo(myUsedRCTabs, h) {
    // For each provided sheet name, set all header values from the header object, then save the sheet.
    myUsedRCTabs.forEach(tabName => {
        gSetVal('mykl', tabName, 'MyLvl', 'MyLvl', h.myLevel);
        gSetVal('mykl', tabName, 'Tier', 'Tier', h.tierString);
        gSetVal('mykl', tabName, 'RC', 'RC', h.rcName_ID);
        gSetVal('mykl', tabName, 'BnsAP', 'APVal', h.bnsAP);
        gSetVal('mykl', tabName, 'TotalAP', 'APVal', h.totalAP);
        gSetVal('mykl', tabName, 'SpentAP', 'APVal', h.apSpent);
        gSetVal('mykl', tabName, 'RemainingAP', 'APVal', h.apRemaining);

        gSaveSheet('mykl', tabName);
    });
} // End fKLSaveRCHeaderInfo




/**
 * Purpose: Verifies all checkboxes in the data rows of specified KL sheets, unchecking any that are invalid, checking any that are 'Free' and valid, and setting cell colors to indicate ability status.
 * Assumptions: The KL sheets are correctly formatted.
 * Notes: Enforces the rule that a higher-tier ability cannot be selected if the tier directly below it is not selected.
 * @param {string[]} myUsedRCTabs - An array of KL sheet names to process.
 * @param {number} myTierNum - The character's current tier number.
 * @returns {void}
 */
function fKLVerifyRCCheckedBoxesSetColors(myUsedRCTabs, myTierNum) {
    const lightRed = '#fc8279';
    const lightGreen = '#a6f04d';
    const lightYellow = '#fce803';
    const lighterYellow = '#ede477';

    // Iterate through each provided sheet name.
    myUsedRCTabs.forEach(tabName => {
        const currentTab = getObjKL_RCTab(tabName, true);
        const numRows = currentTab.arr.length;
        const numCols = currentTab.arr[0].length;

        // Read all existing colors from the sheet first to preserve all original formatting.
        const colorArr = currentTab.ref.getRange(1, 1, numRows, numCols).getBackgrounds();
        const lastCol = numCols - 2; // Loop until the second to last column to safely access c+1

        for (let r = currentTab.dataFirst_R; r <= currentTab.dataLast_R; r++) {
            for (let c = 1; c <= lastCol; c++) {
                // If there is a checkbox in the current cell, verify it and set colors.
                if (currentTab.arr[r][c] === true || currentTab.arr[r][c] === false) {
                    const cardText = currentTab.arr[r][c + 1];
                    if (!cardText) continue;

                    const klCard = fGetKLCardObj(cardText, tabName, r, c + 1);

                    // Perform boundary-safe dependency checks.
                    const isDependentAboveUn_Checked = (r > currentTab.dataFirst_R) && (currentTab.arr[r - 1][c] === false);

                    // Verify and set the checkbox state based on tier and dependency rules.
                    if (klCard.tier > myTierNum || isDependentAboveUn_Checked) {
                        currentTab.arr[r][c] = false;
                    } else if (klCard.isFree) {
                        currentTab.arr[r][c] = true;
                    }

                    // Determine and set the background color based on the ability's final state.
                    if (klCard.tier > myTierNum) {
                        colorArr[r][c + 1] = lightRed; // Illegal ability (too high tier)
                    } else if (currentTab.arr[r][c] === true) {
                        colorArr[r][c + 1] = lightGreen; // Legal and selected ability
                    } else {
                        colorArr[r][c + 1] = (isDependentAboveUn_Checked) ? lighterYellow : lightYellow; // Legal but not selected ability
                    }
                }
            }
        }

        // Save the updated values and the new colors in two separate, fast operations.
        gSaveSheet('mykl', tabName);
        currentTab.ref.getRange(1, 1, numRows, numCols).setBackgrounds(colorArr);
    });
} // End fKLVerifyRCCheckedBoxesSetColors




/**
 * Purpose: Alerts the user if they have overspent their AP total.
 * Assumptions: The input object 'h' contains an `apRemaining` number property.
 * Notes: This provides a non-interrupting warning to the user.
 * @param {object} h - A pre-calculated header object containing all AP values.
 * @returns {void}
 */
function fKLAlertIfOverspentAP(h) {
    if (h.apRemaining < 0) {
        const errorString = `You have overspent your AP by ${-h.apRemaining}.`;
        SpreadsheetApp.getUi().alert('AP Warning', errorString, SpreadsheetApp.getUi().ButtonSet.OK);
    }
} // End fKLAlertIfOverspentAP




/**
 * Purpose: Builds a master array of unique, known KeyLine abilities from specified RC tabs, consolidating to the highest version and buff number for each.
 * Assumptions: Assumes fGetKLCardObj and getObjKL_KLTab functions exist and work as expected.
 * Notes: A KeyLine card represents a specific ability or trait.
 * @param {string[]} myUsedRCTabs - An array of KeyLine RC sheet names (e.g., 'CIV', 'HBE') to process.
 * @returns {object[]} An array of simplified objects, each containing an `id`, the maximum `buffNum`, and the maximum `verNum`.
 */
function fKLBuildKnownKLs(myUsedRCTabs) {
    const knownKLs = [];

    // Loop through each used RC tab.
    myUsedRCTabs.forEach(tabName => {
        const currentTab = getObjKL_RCTab(tabName, true);
        const lastCol = currentTab.arr[0].length - 2; // Loop until the second to last column to safely access c+1

        // Iterate through the data rows and columns to find checked boxes.
        for (let r = currentTab.dataFirst_R; r <= currentTab.dataLast_R; r++) {
            for (let c = 1; c <= lastCol; c++) {

                // If a checkbox in the current cell is checked (TRUE).
                if (currentTab.arr[r][c] === true) {
                    const cardText = currentTab.arr[r][c + 1];
                    if (!cardText) continue;

                    // Create a card object from the cell text.
                    const klCard = fGetKLCardObj(cardText, tabName, r, c + 1);
                    const existingKL = knownKLs.find(kl => kl.id === klCard.id);

                    if (existingKL) {
                        // If it exists, update the appropriate buff or version number to the highest value found.
                        if (klCard.buffVerType === 'b') {
                            existingKL.bestBuff = Math.max(existingKL.bestBuff, klCard.buffVerNum);
                        } else if (klCard.buffVerType === 'v') {
                            existingKL.bestVer = Math.max(existingKL.bestVer, klCard.buffVerNum);
                        }
                    } else {
                        // If it's a new ability, create the simplified object using ternary operators and add it.
                        const newKL = {
                            id: klCard.id,
                            bestBuff: (klCard.buffVerType === 'b') ? klCard.buffVerNum : 0,
                            bestVer: (klCard.buffVerType === 'v') ? klCard.buffVerNum : 0,
                        };
                        knownKLs.push(newKL);
                    }
                }
            }
        }
    });

    return knownKLs;
} // End fKLBuildKnownKLs




/**
 * Purpose: Expands any KeyLine Groups from a list of known KLs and consolidates the result with the original individual KLs, ensuring each unique KL ID is represented once with its highest found buff and version.
 * Assumptions: The input array 'knownKLs' contains objects with id, bestBuff, and bestVer properties.
 * Notes: This function flattens a list that may contain high-level groups, producing a final, definitive list of all individual abilities and their most powerful discovered stats.
 * @param {object[]} knownKLs - An array of simplified, known KL objects.
 * @returns {object[]} A new, consolidated array of unique KL objects.
 */
function fKLExtractKLsFromKLGroups(knownKLs) {
    const { flatList, parentIDs } = fKLExpandKLGroups(knownKLs);
    const consolidatedList = fKLFlattenDuplicates(flatList);

    // Filter out the parent group IDs to create the final list.
    const finalList = consolidatedList.filter(kl => !parentIDs.includes(kl.id));

    return finalList;
} // End fKLExtractKLsFromKLGroups




/**
 * Purpose: Recursively expands any KeyLine Groups to create a "flat" list of all individual KLs.
 * Assumptions: A KL Group's KLList can contain IDs of other groups.
 * Notes: Uses an iterative approach with a stack to prevent deep recursion errors and a Set to avoid infinite loops from circular dependencies.
 * @param {object[]} knownKLs - An array of simplified, known KL objects.
 * @returns {object} An object containing the `flatList` of all individual KLs and the `parentIDs` of all groups that were expanded.
 */
function fKLExpandKLGroups(knownKLs) {
    const finalList = [];
    const groupsToProcess = [];

    // 1. Separate initial KLs into individuals and top-level groups.
    for (const kl of knownKLs) {
        if (gTestID('db', 'KeyLines', kl.id)) {
            groupsToProcess.push(kl); // It's a group, add to the processing stack.
        } else {
            finalList.push(kl); // It's an individual, add directly to the final list.
        }
    }

    const processedGroupIDs = new Set(); // Tracks all groups that have been expanded.

    // 2. Iteratively process the stack of groups until it's empty.
    while (groupsToProcess.length > 0) {
        const currentGroup = groupsToProcess.pop();

        // Avoid infinite loops from circular dependencies (e.g., Group A contains B, B contains A).
        if (processedGroupIDs.has(currentGroup.id)) {
            continue;
        }
        processedGroupIDs.add(currentGroup.id);

        const childIdListString = gGetVal('db', 'KeyLines', currentGroup.id, 'KLList');
        if (!childIdListString) continue;

        const childIdArray = childIdListString.split(',');

        // 3. Process each child ID from the current group.
        for (const childId of childIdArray) {
            const trimmedId = childId.trim();
            if (gTestID('db', 'KeyLines', trimmedId)) {
                // The child is another group, add it to the stack to be processed.
                // It inherits the buff/ver from its immediate parent.
                groupsToProcess.push({
                    id: trimmedId,
                    bestBuff: currentGroup.bestBuff,
                    bestVer: currentGroup.bestVer,
                });
            } else {
                // The child is an individual KL, add it to our final flat list.
                finalList.push({
                    id: trimmedId,
                    bestBuff: currentGroup.bestBuff,
                    bestVer: currentGroup.bestVer,
                });
            }
        }
    }

    return { flatList: finalList, parentIDs: Array.from(processedGroupIDs) };
} // End fKLExpandKLGroups




/**
 * Purpose: Consolidates a list of KLs to ensure each unique ID is represented only once, with the highest buff and version numbers - no duplicates of any kind!
 * Assumptions: The input array may contain KL objects with duplicate IDs.
 * Notes: Uses a Map to efficiently handle the consolidation. This single step resolves all potential duplicates, regardless of whether they were in the original list, created by expanding multiple KL Groups, or came from duplicate KL Groups themselves.
 * @param {object[]} allIndividualKLs - A "flat" list of KL objects, possibly with duplicates.
 * @returns {object[]} A new array of unique, consolidated KL objects.
 */
function fKLFlattenDuplicates(allIndividualKLs) {
    const consolidatedKLs = new Map();

    // Consolidate the flat list to find the max buff/ver for each unique ID.
    for (const kl of allIndividualKLs) {
        const existingKL = consolidatedKLs.get(kl.id);
        if (existingKL) {
            existingKL.bestBuff = Math.max(existingKL.bestBuff, kl.bestBuff);
            existingKL.bestVer = Math.max(existingKL.bestVer, kl.bestVer);
        } else {
            consolidatedKLs.set(kl.id, { ...kl });
        }
    }

    return Array.from(consolidatedKLs.values());
} // End fKLFlattenDuplicates



/**
 * Purpose: Resizes and populates the 'KnownAbilities' sheet with data, including parent kit info, then removes abilities irrelevant to the current RaceClass.
 * Assumptions: The input array has already been fully expanded and consolidated. The 'Abilities' and 'Elements' sheets in the 'db' are correctly formatted.
 * Notes: This function overwrites existing data on the 'KnownAbilities' sheet. The final step removes abilities where both skill slots are non-applicable ('~') for the current RaceClass.
 * @param {object[]} extractedKLs - The final array of unique, individual KL objects.
 * @returns {void}
 */
function fKLBuildKnownAbilitiesSheet(extractedKLs) {
    let abil = getObjKnownAbilities(true);
    const numAbilities = extractedKLs.length;

    // Adjust the number of rows on the sheet to exactly fit the new data.
    const newRowCount = abil.dataFirst_R + numAbilities;
    const currentRowCount = abil.ref.getMaxRows();

    if (newRowCount > currentRowCount) {
        abil.ref.insertRowsAfter(currentRowCount, newRowCount - currentRowCount);
    } else if (newRowCount < currentRowCount) {
        abil.ref.deleteRows(newRowCount + 1, currentRowCount - newRowCount);
    }

    // Reload the sheet object to get a correctly sized array, then clear the data portion.
    abil = getObjKnownAbilities(true);
    if (abil.dataLast_R >= abil.dataFirst_R) {
        gFillArraySection(abil.arr, abil.dataFirst_R, abil.dataLast_R, 0, abil.arr[0].length - 1, '');
    }

    // Populate the KnownAbilities array with the extracted KLs.
    const rcID = gGetIDFromString(gGetVal('mykl', 'All', 'RC', 'RC'));
    const rcID_C = gHeaderC('db', 'Abilities', rcID);

    for (let i = 0; i < numAbilities; i++) {
        const r = abil.dataFirst_R + i;
        const kl = extractedKLs[i];
        abil.arr[r][abil.id_C] = kl.id;
        abil.arr[r][abil.nameID_C] = gGetVal('db', 'Elements', kl.id, 'Name_ID');
        abil.arr[r][abil.ver_C] = kl.bestVer;
        abil.arr[r][abil.buff_C] = kl.bestBuff;
        const [parentKitName,parentKitID] = fKLGetParentKitNameAndID(kl.id);
        abil.arr[r][abil.kitID_C] = parentKitID;
        abil.arr[r][abil.parentKit_C] = parentKitName;
        abil.arr[r][abil.notes_C] = gGetVal('db', 'Elements', kl.id, 'Notes');

        if (gTestID('db', 'Abilities', kl.id)) {
            abil.arr[r][abil.sk1Typ_C] = gGetVal('db', 'Abilities', kl.id, 'SkTyp1');
            abil.arr[r][abil.sk2Typ_C] = gGetVal('db', 'Abilities', kl.id, 'SkTyp2');
            abil.arr[r][abil.base1_C] = gGetVal('db', 'Abilities', kl.id, 'Base1');
            abil.arr[r][abil.base2_C] = gGetVal('db', 'Abilities', kl.id, 'Base2');
            // Assign the RC specific PlAGHE if it exists, else the default PlAGHE if it exists, else '~'
            abil.arr[r][abil.sk1PLAGHE_C] = gGetVal('db', 'Abilities', kl.id, rcID_C) ||gGetVal('db', 'Abilities', kl.id, 'DefaultPLAGHESk1') || '~';
            abil.arr[r][abil.sk2PLAGHE_C] = gGetVal('db', 'Abilities', kl.id, rcID_C + 1) || gGetVal('db', 'Abilities', kl.id, 'DefaultPLAGHESk2') || '~';

            // Fills in Act, Dur, Rng, Meta, Uses, Regain to KL 'KnownAbilities' from DB 'Versions'
            fKLFillInDBVersionsStats(abil, r);
        }
    }
        
    // Save and refresh sheet then calculate KitBuffs for kit feats and then final skills  
    gSaveSheet('mykl', 'knownabilities');
    abil = getObjKnownAbilities(true);
    fKLCalcKitBuffsAndFinalSkills(abil);
    gSaveSheet('mykl', 'knownabilities');
    
    // Get a fresh reference to the data just saved to the sheet.
    const finalAbil = getObjKnownAbilities(true);
    const sheetValues = finalAbil.arr;

    // Loop backwards from the last data row to the first.
    for (let i = finalAbil.dataLast_R; i >= finalAbil.dataFirst_R; i--) {
        const row = sheetValues[i];
        if (row[finalAbil.sk1PLAGHE_C] === '~' && row[finalAbil.sk2PLAGHE_C] === '~') {
            // Delete the corresponding row from the sheet (i + 1 converts 0-based index to 1-based row number).
            finalAbil.ref.deleteRow(i + 1);
        }
    }

    // Final load to capture the final KnownAbilities structure and data after the deletions
    getObjKnownAbilities(true);

} // End fKLBuildKnownAbilitiesSheet





/**
 * Purpose: Retrieves the full Name_ID and the 6-character ID of an ability's parent kit.
 * Assumptions: The 'Abilities' sheet in the 'db' spreadsheet is correctly formatted with a 'ParentKit' column.
 * Notes: Performs multiple validations to ensure valid data is returned.
 * @param {string} abilID - The 6-character ID of the child ability to check.
 * @returns {string[]} An array containing two strings: [parentKitName_ID, parentKitID]. Returns ['', ''] if no valid parent is found.
 */
function fKLGetParentKitNameAndID(abilID) {
    if (gTestID('db', 'Abilities', abilID)) {
        const parentKitName_ID = gGetVal('db', 'Abilities', abilID, 'ParentKit');
        const kitID = gGetIDFromString(parentKitName_ID);
        if (gTestID('db', 'Abilities', kitID)) {
            return [parentKitName_ID,kitID];
        }
    }
    return ['',''];
} // End fKLGetParentKitNameAndID





/**
 * Purpose: First populates the 'KitBuff' column for all abilities, then calculates the 'FinalSk1' and 'FinalSk2' values for every ability.
 * Assumptions: This function is called after the 'KnownAbilities' sheet has been populated and all 'Buff' values for parent kits are present.
 * Notes: This function uses a two-pass approach. The first pass gathers all kit buff dependencies. The second pass calculates the final skills, ensuring all prerequisite data is available. It modifies the 'KnownAbilities' array in memory and does not save the changes to the sheet.
 * @param {object} abil - The sheet object for 'KnownAbilities', typically from getObjKnownAbilities().
 * @returns {void}
 */
function fKLCalcKitBuffsAndFinalSkills(abil) {
    // First Pass: Populate all KitBuff values. This ensures that when the second pass runs,
    // the buff value for any parent kit is already available in the array, regardless of row order.
    for (let r = abil.dataFirst_R; r <= abil.dataLast_R; r++) {
        const abilRow = abil.arr[r];
        const parentKitID = abilRow[abil.kitID_C];

        if (parentKitID) {
            // If a parent kit exists, look up its 'Buff' value from within the same sheet and assign it.
            abilRow[abil.kitBuff_C] = gGetVal('mykl', 'KnownAbilities', parentKitID, 'Buff');
        }
    }

    // Second Pass: Calculate the Final Skills for every ability.
    // Now that all KitBuffs are populated, this calculation will be correct.
    for (let r = abil.dataFirst_R; r <= abil.dataLast_R; r++) {
        fKLCalcFinalSkills(abil, r);
    }

} // End fKLCalcKitBuffsAndFinalSkills


/**
 * Purpose: Calculates the final skill values for a known ability based on its bases, PLG rating, version, and buff numbers.
 * Assumptions: This function modifies the provided ability object's array directly by reference.
 * Notes: The final skill is a weighted combination of the calculated base, PLG rating, version bonus, and buff bonus.
 * @param {object} abil - The entire KnownAbilities sheet object from getObjKnownAbilities().
 * @param {number} r - The 0-indexed row of the ability to calculate.
 * @returns {void}
 */
function fKLCalcFinalSkills(abil, r) {
    const row = abil.arr[r];
    const level = gCharLvl();

    const initVer = row[abil.ver_C] || 0;
    const ver = (initVer >= 1) ? initVer - 1 : initVer; // Version 1 provides a 0 bonus.
    const buff = Math.max(row[abil.buff_C],row[abil.kitBuff_C]) || 0;
    const base1 = row[abil.base1_C];
    const base2 = row[abil.base2_C];
    const sk1PLG = row[abil.sk1PLAGHE_C];
    const sk2PLG = row[abil.sk2PLAGHE_C];
    const plgMap = new Map([['T', 2], ['P', 5], ['L', 8], ['A', 10], ['R', 12], ['O', 14], ['Y', 16], ['G', 18], ['B', 20], ['I', 25], ['V', 30], ['S', 40], ['U', 50], ['E', 60]]);

    // Calculate new Bases after PLG Map and base(10) adjustments.
    let plgBase1 = '~';
    if (!isNaN(base1) && base1 > 0 && plgMap.has(sk1PLG)) {
        plgBase1 = Math.round(plgMap.get(sk1PLG) * base1 / 10);
    }
    let plgBase2 = '~';
    if (!isNaN(base2) && base2 > 0 && plgMap.has(sk2PLG)) {
        plgBase2 = Math.round(plgMap.get(sk2PLG) * base2 / 10);
    }

    // Calculate Ver and Buff effects for FinalSks.
    row[abil.finalSk1_C] = '';
    if (plgBase1 !== '~') {
        const combine1 = [plgBase1, ver * 3, level/3, buff * 5];
        combine1.sort((a, b) => b - a);
        row[abil.finalSk1_C] = Math.round(combine1[0] + combine1[1] / 2 + combine1[2] / 4 + combine1[3]/8);
    }

    row[abil.finalSk2_C] = '';
    if (plgBase2 !== '~') {
        const combine2 = [plgBase2, ver * 3, buff * 5];
        combine2.sort((a, b) => b - a);
        row[abil.finalSk2_C] = Math.round(combine2[0] + combine2[1] / 2 + combine2[2] / 4);
    }
} // End fKLCalcFinalSkills




/**
 * Purpose: Fills in the Act, Dur, Rng, Meta, Uses, and Regain stats for an ability based on its version number.
 * Assumptions: The ability's version is available. It will find the highest valid version stats from the DB that is less than or equal to the ability's current version.
 * Notes: If no valid version is found in the database, the stats fields will not be populated.
 * @param {object} abil - The entire MyAbilities sheet object from getObjKLMyAbilities.
 * @param {number} r - The 0-indexed row of the ability to update in the abil.arr.
 * @returns {void}
 */
function fKLFillInDBVersionsStats(abil, r) {
    const dbVer = getObjDBVersions();
    const abilRow = abil.arr[r];

    const abilID = abilRow[abil.id_C];
    let tryVerNum = abilRow[abil.ver_C] || 1; // Default to 1 if no version is set
    let tryVerID = `${abilID}.v${tryVerNum}`;

    // Decrement the version number until a valid entry is found in DB 'Versions'
    while (tryVerNum > 0 && !gKeyR('db', 'Versions', tryVerID)) { // Note can't use gTestID as this is a verID not an ID
        tryVerNum--;
        if (tryVerNum > 0) {
            tryVerID = `${abilID}.v${tryVerNum}`;
        }
    }

    // If a valid version was found (tryVerNum > 0), populate the stats
    if (tryVerNum > 0) {
        const db_R = gKeyR('db', 'Versions', tryVerID);
        abilRow[abil.act_C] = dbVer.arr[db_R][dbVer.act_C];
        abilRow[abil.dur_C] = dbVer.arr[db_R][dbVer.dur_C];
        abilRow[abil.rng_C] = dbVer.arr[db_R][dbVer.rng_C];
        abilRow[abil.meta_C] = dbVer.arr[db_R][dbVer.meta_C];
        abilRow[abil.uses_C] = dbVer.arr[db_R][dbVer.uses_C];
        abilRow[abil.regain_C] = dbVer.arr[db_R][dbVer.regain_C];
    }

} // End fKLFillInDBVersionsStats




/**
 * Purpose: Copies all ability and gear Name_IDs from the KL 'KnownAbilities' and DB 'Gear' sheets to the CS 'List' sheet, resizing the destination sheet if necessary.
 * Assumptions: The 'KnownAbilities' sheet on 'mykl', 'Gear' on 'db', and 'List' sheet on 'mycs' exist and are properly formatted.
 * Notes: This function will overwrite the existing ability list on the Character Sheet.
 * @returns {void}
 */
function fKLCopyKnownAbilitiesToCS() {
    const kl = getObjKnownAbilities(true);
    let cs = getObjCSList(true);
    const gear = getObjDBGear(true);

    // Extract the list of gear Name_IDs from the DB 'Gear' sheet, filtering out all armor and weapons (as these will be in listOfAbilName_ID).
    const listOfGearName_ID = gear.arr
        .slice(gear.dataFirst_R, gear.dataLast_R + 1)
        .map(row => row[gear.nameID_C])
        .filter(nameID => nameID && !nameID.startsWith('Armor:') && !nameID.startsWith('Wpn:'));

    // Extract the list of ability Name_IDs from the KL 'KnownAbilities' sheet.
    const listOfAbilName_ID = kl.arr.slice(kl.dataFirst_R, kl.dataLast_R + 1).map(row => row[kl.nameID_C]);

    // Combine gear and abilities, remove duplicates and blanks, then sort alphabetically.
    const combinedList = [...listOfGearName_ID, ...listOfAbilName_ID];
    const listOfAllName_ID = [...new Set(combinedList)].filter(Boolean).sort();

    // Determine if the CS List sheet needs more rows to accommodate all items.
    const numTotalItems = listOfAllName_ID.length;
    const numCSAbils = (cs.dataLast_R - cs.dataFirst_R + 1);

    // Clear the existing ability list on the CS List sheet.
    gFillArraySection(cs.arr, cs.dataFirst_R, cs.dataLast_R, cs.abilityNameID_C, cs.abilityNameID_C, '');

    if (numTotalItems > numCSAbils) {
        const rowsToAdd = numTotalItems - numCSAbils;
        cs.ref.insertRowsAfter(cs.ref.getMaxRows(), rowsToAdd);
        cs = getObjCSList(true); // Recache the sheet object after resizing.
    }

    // Populate the CS List array with the sorted list of all items.
    for (let i = 0; i < listOfAllName_ID.length; i++) {
        const targetRow = cs.dataFirst_R + i;
        if (targetRow <= cs.dataLast_R) {
            cs.arr[targetRow][cs.abilityNameID_C] = listOfAllName_ID[i];
        }
    }

    // Save the entire updated array back to the CS 'List' sheet.
    gSaveSheet('mycs', 'list');

} // End fKLCopyKnownAbilitiesToCS





/**
 * Updates element names in specified 'KL' sheets.
 * It iterates through each sheet, finds rows marked with 'true' in the first column,
 * looks up the element ID from the second column in a central database,
 * and writes the corresponding element name back into that cell.
 */
function fKLUpdateKLElementNames() {

  g.klRCSheetNames.forEach(tabName => {
    // Get the object for the current tab, forcing a reload to ensure data is current.
    const currentTab = getObjKL_RCTab(tabName, true);

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
 * Assumptions: cardText must be a string with exactly two newline characters. The AP cost is either '$Free' or '$' followed by a number.
 * Notes: This function is the central parser for KL card data.
 * @param {string} cardText - A formatted string containing the card's name, cost, tier, ID, and buff info.
 * @param {string} tabName - The name of the sheet where the cardText is located.
 * @param {number} r - The 0-indexed row of the cell.
 * @param {number} c - The 0-indexed column of the cell.
 * @returns {object} A structured object with properties for name, isFree, apCost, tier, id, buffVerType, and buffVerNum.
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

    // Parse AP Cost and isFree status (from the second line)
    const apMatch = costTierLine.match(/\$(Free|\d+)/i);
    if (!apMatch) throw new Error(`In fGetKLCardObj, the cost/tier line from sheet "${tabName}" at cell ${r},${c} ("${costTierLine}") has an invalid or missing AP cost block (e.g., $5, $Free).`);

    const apValue = apMatch[1];
    if (apValue.toUpperCase() === 'FREE') {
        cardObj.isFree = true;
        cardObj.apCost = 0;
    } else {
        cardObj.isFree = false;
        cardObj.apCost = parseInt(apValue, 10);
        if (isNaN(cardObj.apCost) || cardObj.apCost < 0) {
            throw new Error(`In fGetKLCardObj, the cost/tier line from sheet "${tabName}" at cell ${r},${c} ("${costTierLine}") has an apCost that is not a positive integer.`);
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
    const buffVerMatch = idVerLine.match(/\.([bv])([0-9])$/);
    if (!buffVerMatch) throw new Error(`In fGetKLCardObj, the ID/version line from sheet "${tabName}" at cell ${r},${c} ("${idVerLine}") has an invalid or missing buff/version suffix (e.g., .b1, .v9).`);

    [, cardObj.buffVerType, cardObj.buffVerNum] = buffVerMatch;
    cardObj.buffVerNum = parseInt(cardObj.buffVerNum, 10);

    return cardObj;
} // End fGetKLCardObj




/**
 * Purpose: Updates the name of a KL Card object and rebuilds the source string in the provided array.
 * Assumptions: The input objects are valid.
 * Notes: This function modifies the `currentTab.arr` directly.
 * @param {object} currentTab - The cached sheet object containing the .arr to be modified.
 * @param {object} klCard - The parsed card object to be updated.
 * @param {string} elementName - The new name for the card.
 * @param {number} r - The 0-indexed row in the array to update.
 * @param {number} c - The 0-indexed column in the array to update.
 * @returns {void}
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
 * Assumptions: The klCard object has all the necessary properties (name, isFree, apCost, etc.).
 * Notes: This is the reverse of `fGetKLCardObj`.
 * @param {object} klCard - The object representing the card's data.
 * @returns {string} A formatted, three-line string representing the card.
 */
function fKLCardObjToStr(klCard) {
    // Conditionally format the AP cost string based on the isFree property.
    const apCostString = klCard.isFree ? 'Free' : klCard.apCost;

    return `${klCard.name}\n$${apCostString} T${klCard.tier}\n🔑${klCard.id}.${klCard.buffVerType}${klCard.buffVerNum}`;
} // End fKLCardObjToStr




/**
 * Purpose: A wrapper function that unchecks all 'true' checkboxes across all standard KL RC sheets.
 * Assumptions: The global `g.klRCSheetNames` is populated.
 * Notes: This is a convenience function for a common operation.
 * @returns {void}
 */
function fKLClearAllCheckBoxes() {

    // Call the core function with the global list of all RC sheet names.
    fKLClearAllCheckBoxesFrom(g.klRCSheetNames);

} // End fKLClearAllCheckBoxes



/**
 * Purpose: Iterates through a provided list of KL RC sheets and unchecks any cell containing a boolean 'true' value.
 * Assumptions: The sheet names provided are valid.
 * Notes: This is the core logic function; it is called by other checkbox-clearing functions.
 * @param {(string|string[])} sheetNameList - A single sheet name or an array of sheet names to process.
 * @returns {void}
 */
function fKLClearAllCheckBoxesFrom(sheetNameList) {

    // Ensure the input is an array so .forEach can be used reliably.
    const sheetNames = Array.isArray(sheetNameList) ? sheetNameList : [sheetNameList];

    // For each specified sheet, load it, change all 'true' values to 'false', and save it.
    sheetNames.forEach(tabName => {
        const currentTab = getObjKL_RCTab(tabName, true);
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
 * Purpose: Recalculates and sets the AP Cost for all abilities in the KL RC-style sheets based on their tier progression.
 * Assumptions: The klRCSheetNames array is globally available at g.klRCSheetNames. The fGetKLCardObj and fKLCardObjToStr functions have been updated for the new AP cost system.
 * Notes: This function overwrites the apCost in the card text with a dynamically calculated value.
 * @returns {void}
 */
function fKLSetKLAPCosts() {

    const firstTierMap = {};
    const buffAPCost = [2, 4, 8, 16, 32];
    const verAPCost = [5, 5, 9, 16, 25];

    // For each specified sheet, load it, recalculate AP costs, and save it.
    g.klRCSheetNames.forEach(tabName => {
        const currentTab = getObjKL_RCTab(tabName, true);
        const lastCol = currentTab.arr[0].length - 2; // Loop until the second to last column to safely access c+1

        for (let r = currentTab.dataFirst_R; r <= currentTab.dataLast_R; r++) {

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

                    // Skip cost calculation for 'Free' or $0 abilities.
                    if (klCard.isFree || klCard.apCost === 0) {
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
                    
                    // Rebuild the card string with the new cost and save it to the array.
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
 * @returns {void}
 */
function fKLSetKLNotes() {

    const firstInstanceMap = {};

    g.klRCSheetNames.forEach(tabName => {
        const currentTab = getObjKL_RCTab(tabName, true);
        const numRows = currentTab.arr.length;
        const numCols = currentTab.arr[0].length;

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

        // Save the entire notes array, now containing both old header notes and new ability notes, in a single operation.
        currentTab.ref.getRange(1, 1, numRows, numCols).setNotes(noteArr);
    });

} // End fKLSetKLNotes



/**
 * Purpose: Hides all RC-related sheets in the KeyLine except for the <All> sheet and the one currently selected on the Character Sheet.
 * Assumptions: The g.klRCSheetNames and g.matchingKLRCIDs global arrays are parallel and correctly populated.
 * Notes: This function will now skip clearing and hiding sheets that are already hidden.
 * @returns {string} The name of the character's active RC tab.
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
 * Assumptions: The `g.klRCSheetNames` global array is populated.
 * Notes: This is a utility function for showing all possible RC sheets.
 * @returns {void}
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
////////////////////                                    (end KLs)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

