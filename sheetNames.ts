/*
**
**  Author :  Clément Legouest
**  Date :    26/07/2023
**
**  This script will iterate the sheets of your document and fill the cells given in CELLS_TO_FILL with the name of the sheet
**  except for the sheets given in SHEETNAMES_TO_AVOID
**
*/

const CELLS_TO_FILL: Array<string> = ["D5", "A7"];
const SHEETNAMES_TO_AVOID: Array<string> = ["summary"];

function main(workbook: ExcelScript.Workbook) {
  workbook.getWorksheets().forEach((workSheet) => {
    if (!SHEETNAMES_TO_AVOID.includes(workSheet.getName())) {
      CELLS_TO_FILL.forEach((cellName: string) => {
        workSheet
          .getRange(cellName)
          .setValue(workSheet.getName());
      });
    }
  });
}
