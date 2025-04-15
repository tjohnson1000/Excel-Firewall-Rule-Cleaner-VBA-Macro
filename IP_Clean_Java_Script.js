function main(workbook: ExcelScript.Workbook) {
  // Get the active worksheet
  const sheet = workbook.getActiveWorksheet();

  // Remove unnecessary columns
  const columnsToDelete = ["Application", "Comment", "Name", "Destination", "Service", "Source"];
  const table = sheet.getTables()[0]; // Assumes data is in a table
  columnsToDelete.forEach(columnName => {
    const column = table.getColumnByName(columnName);
    if (column) {
      column.delete();
    }
  });

  // Rename columns
  const renameMap = {
    "Sources": "Source",
    "Destinations": "Destination",
    "Services": "Service",
    "Network Applications": "Application"
  };
  Object.keys(renameMap).forEach(oldName => {
    const column = table.getColumnByName(oldName);
    if (column) {
      column.setName(renameMap[oldName]);
    }
  });

  // Clean Destination column
  const destinationColumn = table.getColumnByName("Destination");
  if (destinationColumn) {
    const rowCount = destinationColumn.getRange().getRowCount();
    for (let i = 0; i < rowCount; i++) {
      const cellValue = destinationColumn.getRange().getCell(i, 0).getValue() as string;
      destinationColumn.getRange().getCell(i, 0).setValue(cleanDestination(cellValue));
    }
  }
}

// Helper function to clean Destination entries
function cleanDestination(inputText: string): string {
  const ipPattern = /\b(?:\d{1,3}\.){3}\d{1,3}\b/;
  const fqdnPattern = /\b[\w-]+(\.[\w-]+)+\.(com|net)\b/;
  
  // Split by delimiters (comma or semicolon)
  const parts = inputText.split(/[,;]/).map(part => part.trim());
  
  // Filter valid entries
  const validEntries = parts.filter(part => ipPattern.test(part) || fqdnPattern.test(part));
  
  // Join valid entries with commas
  return validEntries.join(", ");
}
