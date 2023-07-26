/*
**
**  Author :  Clément Legouest
**  Date :    26/07/2023
**
**  This script will generate a list of all the sheets in the file in the SUMMARY_SHEET_NAME starting at the CELL_TO_START_SUMMARY
**  Every sheet designated in the SHEET_TO_AVOID_IN_SUMMARY will be avoided
**  Every sheet listed in the summary will have a home link to summary
**
*/

const SUMMARY_SHEET_NAME: string = "summary";
const SHEET_TO_AVOID_IN_SUMMARY = ["summary"];
const CELL_TO_START_SUMMARY = [0, 0];
const CELL_LINK_TO_SUMMARY = [0, 1]
const HOME_ICON = "🏠"
const BLUE = "0000FF";
const FONT_SIZE = 24;

function main(workbook: ExcelScript.Workbook) {

  let summarySheet: ExcelScript.Worksheet = workbook.getWorksheet(SUMMARY_SHEET_NAME);
  let currentCell = summarySheet.getCell(CELL_TO_START_SUMMARY[0], CELL_TO_START_SUMMARY[1]);

  workbook.getWorksheets().forEach((worksheet) => {
    if (!SHEET_TO_AVOID_IN_SUMMARY.includes(worksheet.getName())) {

      let homeLinkCell = worksheet.getCell(CELL_LINK_TO_SUMMARY[0], CELL_LINK_TO_SUMMARY[1]);
      homeLinkCell.setFormulaLocal(hyperLinkToTab(SUMMARY_SHEET_NAME, HOME_ICON));
      const homeLinkFont = homeLinkCell.getFormat().getFont();
      homeLinkFont.setSize(FONT_SIZE);

      currentCell.setFormulaLocal(hyperLinkToTab(worksheet.getName()));
      const cellFont = currentCell.getFormat().getFont();
      cellFont.setSize(FONT_SIZE);
      cellFont.setColor(BLUE);

      currentCell = currentCell.getRowsBelow();
    }
  });
}

/**
 * This will return a formated link to a tab
 */
function hyperLinkToTab(tabName: string, fancyName: string = tabName) {
  return `=LIEN_HYPERTEXTE(\"#${tabName}!A1\", \"${fancyName}\")`;
}
