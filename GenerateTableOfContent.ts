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
            currentCell.setValue(hyperLinkToTab(worksheet.getName()));
            /*
            * At the moment, when inserting the formated link as the value of the cell, this is not really working and it shows the value #NAME?
            * I have to manually update the cell doing <F2> then <Enter> to make the link works so this is not very handy
            * I am actually looking for a solution to make this working
            */
            currentCell = currentCell.getRowsBelow();
        }
    });
}

/**
 * This will return a formated link to a tab
 */
function hyperLinkToTab(tabName: string) {
    return "=LIEN_HYPERTEXTE(\"#${tabName}!A1\", \"${tabName}\")";
}
