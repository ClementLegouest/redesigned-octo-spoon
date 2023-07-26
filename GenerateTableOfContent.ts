const SUMMARY_SHEET_NAME: string = "summary";
const SHEET_TO_AVOID_IN_SUMMARY = ["summary"];
const CELL_TO_START_SUMMARY = [0, 0];

function main(workbook: ExcelScript.Workbook) {

    let summarySheet: ExcelScript.Worksheet = workbook.getWorksheet(SUMMARY_SHEET_NAME);
    let currentCell = summarySheet.getCell(CELL_TO_START_SUMMARY[0], CELL_TO_START_SUMMARY[1]);

    workbook.getWorksheets().forEach( (worksheet) => {
        if(!SHEET_TO_AVOID_IN_SUMMARY.includes(worksheet.getName())) {
            currentCell.setValue(worksheet.getName());
            currentCell = currentCell.getRowsBelow();
        }
    });
}
