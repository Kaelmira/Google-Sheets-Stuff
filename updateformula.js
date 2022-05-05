function updateorderfrom() {

  for (let i = 1; i <14; i++){

  var spreadsheet = SpreadsheetApp.getActive().getSheetByName('WK'+i);
  spreadsheet.getRange('H13').setFormula('=FILTER(indirect(H5),indirect(B9)=G7,indirect(B6)<>"x",indirect(B5)<>"",datevalue(F3)>indirect(B7),datevalue(F3)<=indirect(B8))');
  spreadsheet.getRange('H14:H63').clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('H13').copyTo(spreadsheet.getRange('V13'), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  spreadsheet.getRange('V14:V44').clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('H13').copyTo(spreadsheet.getRange('AJ13'), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  spreadsheet.getRange('AJ14:AJ174').clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('H13').copyTo(spreadsheet.getRange('AX13'), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  spreadsheet.getRange('AX14:AX87').clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('H13').copyTo(spreadsheet.getRange('BL13'), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  spreadsheet.getRange('BL14:BL63').clear({contentsOnly: true, skipFilteredRows: true});
  };
};
