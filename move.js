function move() {
  const spreadsheet = SpreadsheetApp.getActive().getActiveSheet();
  const itemList = ['WAREHOUSE ITEMS','ARCHIVED ITEMS','OFFICE ORDER ITEMS','SMALL WARES ITEMS','SHAMROCK ITEMS','MISC ITEMS'];

  let sheetName = spreadsheet.getSheetName();
  let location = sheetName.match(/^\w+/)[0]; //not currently being used
  let moveCOLUMN = 'by'; //this is the column where you much write the destination tab for the row
  let sortCOLUMN = 'b'; //each row in this column has a numerical value that is sorted at the end of this function
  let headerRows = 5;  //this value should be the amount of header rows +1, this is the first row that is not a header

  let length = spreadsheet.getLastRow();
  for (let i = 0; i < length; i++){
    if(spreadsheet.getRange(moveCOLUMN+(headerRows+i)).getValue() != "") {

    let row = spreadsheet.getRange(moveCOLUMN+(headerRows+i)).getRow();
    let newRow = SpreadsheetApp.getActive().getSheetByName(spreadsheet.getRange(moveCOLUMN+(5+i)).getValue()+" ITEMS").getLastRow()+1;
    let currentRowSheet = sheetName;
    let newRowSheet = spreadsheet.getRange(moveCOLUMN+(headerRows+i)).getValue()+" ITEMS";

    SpreadsheetApp.getActive().getActiveSheet().getRange(sortCOLUMN+(headerRows+i)).setValue(99999);

    SpreadsheetApp.getActive().getSheetByName(currentRowSheet).getRange(row+":"+row).
    copyTo(SpreadsheetApp.getActive().getSheetByName(newRowSheet).getRange(newRow+":"+newRow), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    
    SpreadsheetApp.getActive().getSheetByName(currentRowSheet).getRange(row+":"+row).clear({contentsOnly: true, skipFilteredRows: true});
    };
    
    spreadsheet.getRange('5:360').sort({column: 2, ascending: true});

  };

  for (let s = 0; s < itemList.length; s++){
  console.log(itemList[s].toString());
  SpreadsheetApp.getActive().getSheetByName(itemList[s]).getRange('5:360').sort({column: 2, ascending: true});
  SpreadsheetApp.getActive().getSheetByName(itemList[s]).getRange('BY5:BY360').clear({contentsOnly: true, skipFilteredRows: true});
  };

};
