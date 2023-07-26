/*
**
**  Author :  Clément Legouest
**  Date :    26/07/2023
**
**  This script will generate a list of all the sheets in the file in the SUMMARY_SHEET_NAME starting at the CELL_TO_START_SUMMARY
**  Every sheet designated in the SHEET_TO_AVOID_IN_SUMMARY will be avoided
**  // TODO : Add a link to the sheet
**
*/

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
