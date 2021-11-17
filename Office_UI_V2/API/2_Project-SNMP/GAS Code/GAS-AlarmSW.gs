/*主程式的飲用修改GOOGLE表單的AlarmSW 指定"O2"*/

function ON() {
  var url = 'https://docs.google.com/spreadsheets/d/1FApfprfwR77wCJ03ZegJkKIeCJfojUn-YmyuShOnFDU/edit#gid=2048880426';
  var name = 'DB'
  var SpreadSheet = SpreadsheetApp.openByUrl(url);
  var SheetName = SpreadSheet.getSheetByName(name);

  SheetName.getRange('O2').setValue("ON");    
}

function OFF() {
  var url = 'https://docs.google.com/spreadsheets/d/1FApfprfwR77wCJ03ZegJkKIeCJfojUn-YmyuShOnFDU/edit#gid=2048880426';
  var name = 'DB'
  var SpreadSheet = SpreadsheetApp.openByUrl(url);
  var SheetName = SpreadSheet.getSheetByName(name);

  SheetName.getRange('O2').setValue("OFF");    
}