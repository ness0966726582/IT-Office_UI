/*用途異常告警判斷
功能0 更新表格CELL 目的為了刷新公式
功能1 提取GOOGLE表單內的CELL "P2"
功能2 異常告警(PS.需開啟定時啟動功能)
*/

function myFunction() { 
  
  var url = 'https://docs.google.com/spreadsheets/d/1FApfprfwR77wCJ03ZegJkKIeCJfojUn-YmyuShOnFDU/edit#gid=2048880426';
  var name = 'judgement'
  var SpreadSheet = SpreadsheetApp.openByUrl(url);
  var SheetName = SpreadSheet.getSheetByName(name);

//目的是更新表格刷新公式
  SheetName.getRange('I2').setValue("自動刷新"); 
  var name = 'DB'
  var O2 = SpreadSheet.getSheetByName(name).getRange(2,15).getValue()//取得O2數值
  var P2 = SpreadSheet.getSheetByName(name).getRange(2,16).getValue()//取得P2數值>0為異常筆數!! 正常為0
  var message = "TESTING....";
  var token = "0aLQmkLv3xpQtV5lj9vmX5Fi4Y6oLqzn0LauhGqXndU";  //Line Notify 的權杖
  
  if (P2>0 && O2 === "ON"){
    message =P2+"筆數錯誤"
    sendLineNotify(message,token);
    return;
  }
}

function sendLineNotify(message, token){
  var options =
   {
     "method"  : "post",
     "payload" : {"message" : message},
     "headers" : {"Authorization" : "Bearer " + token}
   };
   UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
}