// Sort Headings

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  Menu ()
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// // onInstall //////////////////////////////////////////////////////////////////////////////////////////////////
// // onInstall trigger
// function onInstall() {
//     onOpen();
// } // End onInstall

// // onOpen //////////////////////////////////////////////////////////////////////////////////////////////////
// // onOpen trigger
// function onOpen() {
//     fDOCCreateMenu();
// } // End onOpen

// fDOCCreateMenu //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose: Add options to the Extensions Menu
function fDOCCreateMenu() {
  const ui = DocumentApp.getUi();
  ui.createAddonMenu() // This creates an "Add-ons" menu
    .addSubMenu(
      ui
        .createMenu("Sort Headings")
        .addItem("Sort Headings A to Z", "fDOCMenuSortHeadingsAtoZ")
        .addItem("Sort Headings Z to A", "fDOCMenuSortHeadingsZtoA")
    )
    .addToUi();
} // End fDOCCreateMenu

// Menu Functions //////////////////////////////////////////////////////////////////////////////////////////////////
// Headings Menu
function fDOCMenuSortHeadingsAtoZ() {
  fDOCRunMenuOrButton("SortHeadingsAtoZ");
}
function fDOCMenuSortHeadingsZtoA() {
  fDOCRunMenuOrButton("SortHeadingsZtoA");
}
// End Menu Functions

// fDOCRunMenuOrButton //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> To run all menu & button choices inside a try-catch-error
function fDOCRunMenuOrButton(menuChoice) {
  try {
    switch (menuChoice) {
      // Headings Menu
      case "SortHeadingsAtoZ":
        fDOCSortOnHeader("AtoZ");
        break;
      case "SortHeadingsZtoA":
        fDOCSortOnHeader("ZtoA");
        break;
    }
  } catch (error) {
    DocumentApp.getUi().alert(error); // NOTE: an error of End or end will simply end the program.
  }
} // End fDOCRunMenuOrButton

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  Headings  (end Data)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// fDOCSortOnHeader //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Sorts on the header level at cursor or selection between headers of one type closer to HEADER1
function fDOCSortOnHeader(sortOrder) {
  const doc = DocumentApp.getActiveDocument();
  let body = doc.getBody();
  let { headerAboveIndex, headerBelowIndex, targetHeaderLevel } =
    fDOCFindHeaderLevelAndSurroundingLargerHeaderIndex();

  const headers = fDOCSaveHeadersAndBodyStartEndIndex(
    body,
    headerAboveIndex,
    headerBelowIndex,
    targetHeaderLevel
  );
  if (sortOrder === "AtoZ") {
    headers.sort((a, b) => a.headerText.localeCompare(b.headerText));
  } else {
    headers.sort((a, b) => b.headerText.localeCompare(a.headerText));
  }

  // Insert the new sorted headings (and heading's body) below the last of the existing selected headings
  fDocInstertSortedHeadersAndBody(body, headers, headerBelowIndex);

  // deletes all original headings and body from headerAboveIndex to headerBelowIndex
  fDocRemoveOriginalSections(body, headerAboveIndex, headerBelowIndex);
} // END fDOCSortOnHeader

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////                                  Helper Functions  (end Headings)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// fDOCFindHeaderLevelAndSurroundingLargerHeaderIndex //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Finds the paragraph index of the header above and below the cursor (or selected text) that is one header closer to Heading1 from the current header
// Purpose -> Also returns the targetHeaderLevel as a number 1 to 6
function fDOCFindHeaderLevelAndSurroundingLargerHeaderIndex() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  let cursor = doc.getCursor();
  let selection = doc.getSelection();
  let element, cursorIndex;

  if (cursor) {
    element = cursor.getElement();
  } else if (selection) {
    const elements = selection.getRangeElements();
    if (elements.length > 0) {
      element = elements[0].getElement();
    } else {
      throw new Error("Please click on a heading.");
    }
  } else {
    throw new Error("Please click on a heading.");
  }

  // Find the index of the element's parent in the body
  while (
    element.getParent() &&
    element.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION
  ) {
    element = element.getParent();
  }
  cursorIndex = body.getChildIndex(element);

  let headerAboveIndex = cursorIndex;
  let headerBelowIndex = body.getNumChildren() - 1;

  let targetHeaderLevel = 7;
  const headingMap = {
    [DocumentApp.ParagraphHeading.HEADING1]: 1,
    [DocumentApp.ParagraphHeading.HEADING2]: 2,
    [DocumentApp.ParagraphHeading.HEADING3]: 3,
    [DocumentApp.ParagraphHeading.HEADING4]: 4,
    [DocumentApp.ParagraphHeading.HEADING5]: 5,
    [DocumentApp.ParagraphHeading.HEADING6]: 6,
  };

  // Find the header type at the cursor or selection
  if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
    const paragraph = element.asParagraph();
    const style = paragraph.getHeading();
    targetHeaderLevel = headingMap[style] || 7;
  }

  if (targetHeaderLevel === 7) throw new Error(`Please click on a heading.`);

  // Find first header closer to HEADING1 above the cursor or selection
  for (let i = cursorIndex; i >= 0; i--) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const paragraph = child.asParagraph();
      const style = paragraph.getHeading();
      const headingLevel = headingMap[style] || 7;
      if (headingLevel >= targetHeaderLevel) {
        if (headingLevel !== 7 && headingLevel === targetHeaderLevel)
          headerAboveIndex = i;
      } else {
        break;
      }
    }
  }

  // Find first header closer to HEADING1 below the cursor or selection
  for (let i = cursorIndex + 1; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const paragraph = child.asParagraph();
      const style = paragraph.getHeading();
      const headingLevel = headingMap[style] || 7;
      if (headingLevel < targetHeaderLevel) {
        headerBelowIndex = i;
        break;
      }
    }
  }

  return { headerAboveIndex, headerBelowIndex, targetHeaderLevel };
} // END fDOCFindHeaderLevelAndSurroundingLargerHeaderIndex

// fDOCGetElementsBetweenIndexes //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Grabs all elements (text, headers, formatting, charts, images, pictures, line breaks, etc.) between element index a and b and stores them in an array
function fDOCGetElementsBetweenIndexes(body, a, b) {
  b--;
  const elements = [];

  for (let i = a; i <= b; i++) {
    const element = body.getChild(i).copy();
    elements.push(element);
  }

  return elements;
} // END fDOCGetElementsBetweenIndexes

// fDocInstertSortedHeadersAndBody //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> starting at body index of insertAt, insert all headers in the order found in headers array, and copies all elements from body between headers[] object names of headerIndex + 1 and lastIndex
function fDocInstertSortedHeadersAndBody(body, headers, insertAt) {
  headers.forEach((header) => {
    // Insert the header
    const newHeader = body.insertParagraph(insertAt, header.headerText);
    newHeader.setHeading(
      body.getChild(header.headerIndex).asParagraph().getHeading()
    );
    insertAt++;

    // Insert all elements between headerIndex + 1 and lastIndex
    for (let i = header.headerIndex + 1; i <= header.lastIndex; i++) {
      const element = body.getChild(i).copy(); // Copy the element
      switch (element.getType()) {
        case DocumentApp.ElementType.PARAGRAPH:
          body.insertParagraph(insertAt, element);
          break;
        case DocumentApp.ElementType.TABLE:
          body.insertTable(insertAt, element);
          break;
        case DocumentApp.ElementType.LIST_ITEM:
          body.insertListItem(insertAt, element);
          break;
        case DocumentApp.ElementType.INLINE_IMAGE:
          body.insertImage(insertAt, element);
          break;
        case DocumentApp.ElementType.HORIZONTAL_RULE:
          body.insertHorizontalRule(insertAt);
          break;
        case DocumentApp.ElementType.FOOTNOTE:
          body.insertFootnote(insertAt, element);
          break;
        // Add more cases as needed for other element types
        default:
          throw new Error(`Unsupported element type: ${element.getType()}`);
      }
      insertAt++;
    }
  });
} // END fDocInstertSortedHeadersAndBody

// fDOCSaveHeadersAndBodyStartEndIndex //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Extracts headers within a specified range, and returns an array of objects {headerText, headerIndex, lastIndex}
function fDOCSaveHeadersAndBodyStartEndIndex(
  body,
  headerAboveIndex,
  headerBelowIndex,
  targetHeaderLevel
) {
  const mapHeadingTextToLevel = fDOCgetHeadingTextToLevelMap();

  const headers = [];
  let currentHeader = null;
  for (let i = headerAboveIndex; i < headerBelowIndex; i++) {
    const element = body.getChild(i);
    const elementType = element.getType();

    if (elementType === DocumentApp.ElementType.PARAGRAPH) {
      const paragraph = element.asParagraph();
      const styleLevel = mapHeadingTextToLevel[paragraph.getHeading()];

      if (styleLevel === targetHeaderLevel) {
        if (currentHeader)
          headers.push({
            headerText: currentHeader.header.getText(),
            headerIndex: currentHeader.index,
            lastIndex: i - 1,
          }); // Save the previous header and its indexes
        currentHeader = { header: paragraph.copy(), body: [], index: i };
      } else if (currentHeader) {
        currentHeader.body.push(element.copy());
      }
    } else if (currentHeader) {
      currentHeader.body.push(element.copy());
    }
  }
  if (currentHeader)
    headers.push({
      headerText: currentHeader.header.getText(),
      headerIndex: currentHeader.index,
      lastIndex: headerBelowIndex - 1,
    }); // Save the last header and its indexes

  return headers;
} // END fDOCSaveHeadersAndBodyStartEndIndex

function fDOCgetHeadingTextToLevelMap() {
  return {
    [DocumentApp.ParagraphHeading.HEADING1]: 1,
    [DocumentApp.ParagraphHeading.HEADING2]: 2,
    [DocumentApp.ParagraphHeading.HEADING3]: 3,
    [DocumentApp.ParagraphHeading.HEADING4]: 4,
    [DocumentApp.ParagraphHeading.HEADING5]: 5,
    [DocumentApp.ParagraphHeading.HEADING6]: 6,
  };
}

// fDocRemoveOriginalSections //////////////////////////////////////////////////////////////////////////////////////////////////
// Purpose -> Deletes all body elements from index headerAboveIndex for another (headerBelowIndex - headerAboveIndex)
function fDocRemoveOriginalSections(body, headerAboveIndex, headerBelowIndex) {
  const numElementsToRemove = headerBelowIndex - headerAboveIndex;
  for (let i = 0; i < numElementsToRemove; i++) {
    body.removeChild(body.getChild(headerAboveIndex));
  }
} // END fDocRemoveOriginalSections
