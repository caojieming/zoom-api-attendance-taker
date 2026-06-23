// sorts the subsheets of a google sheet in reverse alphabetical order: Z-A, then 9-0
// this keeps "Base" at the front while sorting all the meetings from most recent to oldest
function sortRecords() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();

    sheets
        .sort((a, b) => b.getName().localeCompare(a.getName()))
        .forEach((sheet, i) => {
            ss.setActiveSheet(sheet);
            ss.moveActiveSheet(i + 1); // positions are 1-based
        });
}
