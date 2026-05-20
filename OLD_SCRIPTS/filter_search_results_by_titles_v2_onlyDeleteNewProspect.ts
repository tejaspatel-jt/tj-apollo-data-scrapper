/**
 * VB Script equivalent:
 * 
 * This script will read the list of titles from the "titles" tab and then loop through the "Contacts" tab 
 * to delete any rows where the "title" column matches any of the titles in the list.
 * Additionally, it will only delete rows where the "source_type" column has the value "NEW_PROSPECT".
 * 
 */
function main(workbook: ExcelScript.Workbook) {
  const contactsSheet = workbook.getWorksheet("Contacts");
  const titlesSheet = workbook.getWorksheet("titles");

  // Load titles from "titles" sheet into a Set
  const titlesData = titlesSheet.getUsedRange().getValues();
  const titlesToDelete = new Set<string>();

  for (let i = 1; i < titlesData.length; i++) {
    const title = String(titlesData[i][0]).trim().toLowerCase();
    if (title !== "") {
      titlesToDelete.add(title);
    }
  }

  // Read Contacts data
  const contactsData = contactsSheet.getUsedRange().getValues();
  const headerRow = contactsData[0];

  let titleColIndex = -1;
  let sourceTypeColIndex = -1;

  // Find required columns
  for (let col = 0; col < headerRow.length; col++) {
    const header = String(headerRow[col]).trim().toLowerCase();

    if (header === "title") {
      titleColIndex = col;
    }

    if (header === "source_type") {
      sourceTypeColIndex = col;
    }
  }

  if (titleColIndex === -1) {
    console.log("ERROR: 'title' column not found in Contacts tab!");
    return;
  }

  if (sourceTypeColIndex === -1) {
    console.log("ERROR: 'source_type' column not found in Contacts tab!");
    return;
  }

  // Delete rows bottom-up where:
  // 1. title is in titles sheet
  // 2. source_type = NEW_PROSPECT
  for (let i = contactsData.length - 1; i >= 1; i--) {
    const cellTitle = String(contactsData[i][titleColIndex]).trim().toLowerCase();
    const sourceType = String(contactsData[i][sourceTypeColIndex]).trim().toUpperCase();

    if (titlesToDelete.has(cellTitle) && sourceType === "NEW_PROSPECT") {
      contactsSheet.getRange(`${i + 1}:${i + 1}`).delete(ExcelScript.DeleteShiftDirection.up);
    }
  }

  console.log("Done! Matching rows with source_type = NEW_PROSPECT have been deleted.");
}