/**
 * VB Script equivalent:
 * 
 * This script will read the list of titles from the "titles" tab and then loop through the "Contacts" tab 
 * to delete any rows where the "title" column matches any of the titles in the list.
 * 
 */
function main(workbook: ExcelScript.Workbook) {

  const contactsSheet = workbook.getWorksheet("Contacts");
  const titlesSheet = workbook.getWorksheet("titles");

  // --- Load titles from Titles tab into a Set (for fast lookup) ---
  const titlesData = titlesSheet.getUsedRange().getValues();
  const titlesToDelete = new Set<string>();

  for (let i = 1; i < titlesData.length; i++) { // i=1 to skip header row
    const title = String(titlesData[i][0]).trim().toLowerCase();
    if (title !== "") {
      titlesToDelete.add(title);
    }
  }

  // --- Find the "title" column index in Contacts tab ---
  const contactsData = contactsSheet.getUsedRange().getValues();
  const headerRow = contactsData[0];

  let titleColIndex = -1;
  for (let col = 0; col < headerRow.length; col++) {
    if (String(headerRow[col]).trim().toLowerCase() === "title") {
      titleColIndex = col;
      break;
    }
  }

  if (titleColIndex === -1) {
    console.log("ERROR: 'title' column not found in Contacts tab!");
    return;
  }

  // --- Loop bottom-up and delete matching rows ---
  const lastRow = contactsData.length;

  for (let i = lastRow - 1; i >= 1; i--) { // i>=1 to skip header row
    const cellTitle = String(contactsData[i][titleColIndex]).trim().toLowerCase();
    if (titlesToDelete.has(cellTitle)) {
      contactsSheet.getRange(`${i + 1}:${i + 1}`).delete(ExcelScript.DeleteShiftDirection.up);
    }
  }

  console.log("Done! Matching rows have been deleted.");
}