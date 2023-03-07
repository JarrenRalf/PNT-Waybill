//function onEdit(e) {
//  try { moveRow(e) } catch (error) { Browser.msgBox(error) }}
//function moveRow(e) {
//
//   var colStart = e.range.columnStart;  // Only look at a single cell edit
//   var active = e.source.getActiveSheet();
//   var name = active.getName();
//    
//   if ( name == "Consignee" && colStart == 6 ) {        // Change to your "From" sheet and Column reference
//   var value = e.value; 
//      
////   if ( value == "fillWaybill") { 
////   var spreadsheet = SpreadsheetApp.getActive();
////    spreadsheet.getCurrentCell().offset(0, -1).activate();
////    spreadsheet.getActiveRange().copyTo(spreadsheet.getRange('WAYBILL!B14'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
////    spreadsheet.getCurrentCell().offset(0, -4, 1, 4).activate();
////    spreadsheet.getActiveRange().copyTo(spreadsheet.getRange('WAYBILL!J7'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, true);
////    spreadsheet.getCurrentCell().offset(0, 5).activate();
////    spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
////     spreadsheet.getRange('WAYBILL!F14').activate();
// }
//}
    
function onEdit(e){   
//  else 
  var colStart = e.range.columnStart;  // Only look at a single cell edit
  var active = e.source.getActiveSheet();
  var name = active.getName();
  
   if ( name == "WAYBILL" && colStart == 16 ) {        // Change to your "From" sheet and Column reference
   var value = e.value; 
      
   if ( value == "saveAddress") {   
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Consignee'), true);
    spreadsheet.getRange('A1').activate();
    spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
    spreadsheet.getCurrentCell().offset(1, 0).activate();
    spreadsheet.getRange('WAYBILL!J7:J10').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, true);
    spreadsheet.getCurrentCell().offset(0, 4).activate();
    spreadsheet.getRange('WAYBILL!B14').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getRange('A:F').activate().sort([{column: 1, ascending: true}, {column: 2, ascending: true}]);
    spreadsheet.getRange('A2').activate();
    spreadsheet.getRange('WAYBILL!P6').clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('WAYBILL!B14').activate();
  }
 }
  else if ( e.range.getA1Notation() === 'J6') {   
    var spreadsheet = SpreadsheetApp.getActive();
    var wb = spreadsheet.getRange('WAYBILL!I8').getValue()+1
    spreadsheet.getRange('WAYBILL!I8').setValue(wb)
    
    if(!(spreadsheet.getRange('J7').getFormula().charAt(0)  == '=' &&
         spreadsheet.getRange('J8').getFormula().charAt(0)  == '=' &&
         spreadsheet.getRange('J9').getFormula().charAt(0)  == '=' &&
         spreadsheet.getRange('J10').getFormula().charAt(0) == '='))
    {
      active.getRange('J7').setFormula('=if($J$6="","x",QUERY(CONSIGNEE,"select A where G like \'%"&$J$6&"%\'" ,))');
      active.getRange('J8').setFormula('=if($J$6="","x",QUERY(CONSIGNEE,"select B where G like \'%"&$J$6&"%\'" ,))');
      active.getRange('J9').setFormula('=if($J$6="","x",QUERY(CONSIGNEE,"select C where G like \'%"&$J$6&"%\'" ,))');
      active.getRange('J10').setFormula('=if($J$6="","x",QUERY(CONSIGNEE,"select D where G like \'%"&$J$6&"%\'" ,))');
      active.getRange('B14').setFormula('=if($J$6="","x",QUERY(CONSIGNEE,"select E where G like \'%"&$J$6&"%\'" ,))');
    }
  }
  else if ( e.range.getA1Notation() === 'J8') {   
    var spreadsheet = SpreadsheetApp.getActive();
    var wb = spreadsheet.getRange('WAYBILL!I8').getValue()+1
    spreadsheet.getRange('WAYBILL!I8').setValue(wb)
    
}  }

function saveLabelAddress() {
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Consignee'), true);
    spreadsheet.getRange('A1').activate();
    spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
    spreadsheet.getCurrentCell().offset(1, 0).activate();
    spreadsheet.getRange('Labels!B9:B12').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, true);
//    spreadsheet.getCurrentCell().offset(0, 4).activate();
//    spreadsheet.getRange('WAYBILL!B14').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getRange('A:F').activate().sort([{column: 1, ascending: true}, {column: 2, ascending: true}]);
    spreadsheet.getRange('A2').activate();
//    spreadsheet.getRange('WAYBILL!O6').clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('Labels!B9').activate();


}
function resetPackingS() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B18:B36').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('d18:d36').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('B19').setFormula('=WAYBILL!$B18');
  spreadsheet.getRange('B20').setFormula('=WAYBILL!$B20');
  spreadsheet.getRange('B21').setFormula('=WAYBILL!$B22');
  spreadsheet.getRange('B22').setFormula('=WAYBILL!$B24');
  spreadsheet.getRange('B23').setFormula('=WAYBILL!$B26');
  spreadsheet.getRange('B24').setFormula('=WAYBILL!$B28');
  spreadsheet.getRange('D19').setFormula('=WAYBILL!$D18');
  spreadsheet.getRange('D20').setFormula('=WAYBILL!$D20');
  spreadsheet.getRange('D21').setFormula('=WAYBILL!$D22');
  spreadsheet.getRange('D22').setFormula('=WAYBILL!$D24');
  spreadsheet.getRange('D23').setFormula('=WAYBILL!$D26');
  spreadsheet.getRange('D24').setFormula('=WAYBILL!$D28');
  spreadsheet.getRange('B19').activate();
};

function resetLabel() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B9').setFormula('=WAYBILL!$J7');
  spreadsheet.getRange('B10').setFormula('WAYBILL!$J8');
  spreadsheet.getRange('B11').setFormula('WAYBILL!$J9');
  spreadsheet.getRange('B12').setFormula('WAYBILL!$J10');
  spreadsheet.getRange('B13').setFormula('=CONCATENATE("",WAYBILL!$K14)');
  spreadsheet.getRange('B9').activate();
};

function printPage() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Print Label').activate(), true);
//  spreadsheet.getRange('E1:E34').activate();
};

function UntitledMacro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B19').activate();
  spreadsheet.getCurrentCell().setFormula('=WAYBILL!B18');
  spreadsheet.getRange('D19:P19').activate();
  spreadsheet.getCurrentCell().setFormula('=WAYBILL!D18');
  spreadsheet.getRange('B20').activate();
  spreadsheet.getCurrentCell().setFormula('=WAYBILL!B20');
  spreadsheet.getRange('D19:P19').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('WAYBILL'), true);
  spreadsheet.getRange('D20:I21').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Consignee'), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Packing Slip'), true);
  spreadsheet.getRange('D20:P20').activate();
  spreadsheet.getCurrentCell().setFormula('=WAYBILL!D22');
};

function kjhjkkjk() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getCurrentCell().offset(-1, -2, 20, 15).activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
};