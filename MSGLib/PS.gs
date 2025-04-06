// PS

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  Menu (end initialize)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// fPSCreateMenu //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Create MetaScape menu
function fPSCreateMenu() {
  SpreadsheetApp.getUi()
    .createMenu("*** PartyLog")
    .addItem("Clear Logs", "fPSMenuClearPartyLogs")
    .addToUi();
  SpreadsheetApp.getUi()
    .createMenu("System ***")
    .addItem("1 - Authorize Script", "fPSMenuAuthorize")
    .addSeparator()
    .addItem("Refresh Menus", "fPSMenuRefreshAll")
    .addToUi();
  if (gGetVal("ps", "Data", "designer", "Val") === true) {
    SpreadsheetApp.getUi()
      .createMenu("DESIGNER")
      .addItem("Hide All", "fPSMenuHideAll")
      .addItem("Un-Hide All", "fPSMenuUn_HideAll")
      .addToUi();
  }
} // End fPSCreateMenu

// Menu Functions //////////////////////////////////////////////////////////////////////////////////////////////////
// PartyLog Menu
function fPSMenuClearPartyLogs() {
  fPSRunMenuOrButton("ClearPartyLogs");
}
// System Menu
function fPSMenuAuthorize() {
  fPSRunMenuOrButton("Authorize");
}
function fPSMenuRefreshAll() {
  fPSRunMenuOrButton("RefreshMenu");
}
// Designer Menu
function fPSMenuHideAll() {
  fPSRunMenuOrButton("HideAll");
}
function fPSMenuUn_HideAll() {
  fPSRunMenuOrButton("Un_HideAll");
}
// End Menu Functions

// Image Button Functions //////////////////////////////////////////////////////////////////////////////////////////////////
// function calc() {fRunMenuOrButton('Calc AP & KL');}
// End Button Functions

// fPSRunMenuOrButton //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> To run all menu & button choices inside a try-catch-error
function fPSRunMenuOrButton(menuChoice) {
  try {
    switch (menuChoice) {
      // PartyLog Menu
      case "ClearPartyLogs":
        fPSClearPartyLogs();
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
        fPSCreateMenu();
        break;
      // Designer Menu
      case "HideAll":
        gHideAll("ps");
        break;
      case "Un_HideAll":
        gUn_HideAll("ps");
        break;
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert(error); // NOTE: an error of End or end will simply end the program.
  }
} // End fPSRunMenuOrButton

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  PartyLog Menu  (end Data)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// fPSClearPartyLogs //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Clear's all PartyLogs
function fPSClearPartyLogs(ss) {
  // Set <PartyLog> data
  const partyLogRef = gSheetRef("ps", "PartyLog");
  const partyLogArr = gArr("ps", "PartyLog");
  const partyLogFirst_R = gKeyR("ps", "PartyLog", "URL");
  const partyLogLast_R = gKeyR("ps", "PartyLog", "Log");
  const partyLogFirst_C = gHeaderC("ps", "PartyLog", "Slot1");
  const partyLogLast_C = gHeaderC("ps", "PartyLog", "Slot9");

  gFillArraySection(
    partyLogArr,
    partyLogFirst_R,
    partyLogLast_R,
    partyLogFirst_C,
    partyLogLast_C,
    ""
  );
  gSaveArraySectionToSheet(
    partyLogRef,
    partyLogArr,
    partyLogFirst_R,
    partyLogLast_R,
    partyLogFirst_C,
    partyLogLast_C
  );
} // End fPSClearPartyLogs

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  DESIGNER  (end PartyLog Menu)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
