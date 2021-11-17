/*用途: 抓取GoogleSeet表單的內容作檢視欄位比對---->抓取符合的行  

另外搭配兩支副程式
0.Google Sheet ------->https://docs.google.com/spreadsheets/d/1FApfprfwR77wCJ03ZegJkKIeCJfojUn-YmyuShOnFDU/edit#gid=2048880426
1.AlarmSW.gs--------->切換告警"P2"功能ON/OFF修改表單
2.notify-bot.gs--------->比對表單"O2"(PS.需要開啟時間啟動)

*/

var CHANNEL_ACCESS_TOKEN = "fdsuKXu03UXy7Ui8PEmm5APv+B5xGe2jZRtS+rLG+HyUaA3MwQbP33cW0fbYki4e8dfU4UAH+PvGe5NZxDO/1l2JQFn3aWAXnfqm95MRiPF5rU8FSF8XJWFqukmvXn3jMHxpJdF4x8t/yxg92I1ltQdB04t89/1O/w1cDnyilFU=";
var spreadSheetId = "1FApfprfwR77wCJ03ZegJkKIeCJfojUn-YmyuShOnFDU";    //試算表 ID
var sheetName = ["DB"];    //工作表名稱
var searchColumn = [13];    //搜尋第幾欄的資料
var whiteList = ["U49b744c8412a42b81c73387b58e68b7a"];  //白名單（允許取得資料的使用者ID
var whiteListMode = 0;    //白名單模式，1 表示需要設定白名單才能進行查詢，0 表示任何人加了好友都可以查詢。
var fuzzySearch = 1;    //模糊搜尋模式，1 表示表格內部分字串相同即取出資料，0 表示儲存格內容完全相同才取出資料。
var spreadSheet = SpreadsheetApp.openById(spreadSheetId);
var url = 'https://docs.google.com/spreadsheets/d/1FApfprfwR77wCJ03ZegJkKIeCJfojUn-YmyuShOnFDU/edit#gid=2048880426';
var name = 'DB'
var SpreadSheet = SpreadsheetApp.openByUrl(url);

function doPost(e) {
  var userData = JSON.parse(e.postData.contents);
  var allowed = whiteListMode;
  var clientID = userData.events[0].source.userId;
  var replyToken = userData.events[0].replyToken;
  
  if ((userData.events[0].type == "follow" || userData.events[0].message.text.toLowerCase() == "findmyid") && whiteListMode === 0) {
    var replyMessage = [{type:"text", text:"您的使用者 ID 是「" + clientID + "」，請將此 ID 告知此官方帳號的擁有者加入白名單後才能開始查詢資料。"}];
    sendReplyMessage(CHANNEL_ACCESS_TOKEN, replyToken, replyMessage);
    return;
  }
  //Alarm切換on/off 
  if ((userData.events[0].type == "follow" || userData.events[0].message.text.toLowerCase() == "on") && whiteListMode === 0) {
    var replyMessage = [{type:"text", text:"Alarm status「ON啟動異常告警」。"}];
    sendReplyMessage(CHANNEL_ACCESS_TOKEN, replyToken, replyMessage);
    ON();
    return;
  }
  if ((userData.events[0].type == "follow" || userData.events[0].message.text.toLowerCase() == "off") && whiteListMode === 0) {
    var replyMessage = [{type:"text", text:"Alarm status「OFF關閉異常告警」。"}];
    sendReplyMessage(CHANNEL_ACCESS_TOKEN, replyToken, replyMessage);
    OFF();
    return;
  }
  if (userData.events[0].type != "message") {return;}
  if (userData.events[0].message.type != "text") {return;}
  // 檢查是否是允許的用者提出搜尋需求
  for (var i = 0; i < whiteList.length; i++) {
    if (whiteList[i] == clientID) {
      allowed = 0;
      break;
    }
  }
  if (allowed === 1) {
    var replyMessage = [{type:"text", text:"您的使用者 ID 是「" + clientID + "」，請將此 ID 告知此官方帳號的擁有者加入白名單後才能開始查詢資料。"}];
    sendReplyMessage(CHANNEL_ACCESS_TOKEN, replyToken, replyMessage);
    return;
  }
  var replyMessage = [];
  var replyToken = userData.events[0].replyToken;
  var searchContent = userData.events[0].message.text;
  for (var i = 0; i < sheetName.length; i++) {
    var searchResult = [];
    var sheet = spreadSheet.getSheetByName(sheetName[i]);
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    var sheetData = sheet.getSheetValues(1, 1, lastRow, lastColumn);
    for (var j = 0; j < searchColumn.length; j++){
      var searchTemp = sheetData.filter(function(item, index, array){
        if (fuzzySearch == 0) {return item[searchColumn[j] - 1].toString().toLowerCase() === searchContent.toLowerCase();}
        else {return item[searchColumn[j] - 1].toString().toLowerCase().indexOf(searchContent.toLowerCase()) != -1 ;}
      });
      searchResult = searchResult.concat(searchTemp);
    }
    if (searchResult.length > 0) {
      var replyContent = "";
      searchResult = uniqueArrayElement(searchResult);
      replyContent = "在「" + sheetName[i] + "」中搜尋到 " + searchResult.length + " 筆資料";
      for (var k = 0; k < searchResult.length; k++) {
        replyContent += "\n\n" + sheetData[0][0] + "：" + searchResult[k][0];
        for (var l = 1; l < lastColumn; l++) {
          replyContent += "\n" + sheetData[0][l] + "：" + searchResult[k][l];
        }
      }
      replyMessage.push({type:"text", text:replyContent});
    }
    if (replyMessage.length == 5) {break;}
  }
  //if (replyMessage.length == 0) {replyMessage.push({type:"text", text:"查詢不到「" + searchContent + "」的資料"});}
  sendReplyMessage(CHANNEL_ACCESS_TOKEN, replyToken, replyMessage);
}
//移除陣列中重複的元素
function uniqueArrayElement(arrayData) {
  var result = arrayData.filter(function(element, index, arr){
    return arr.indexOf(element) === index;
  });
  return result;
}
//回送 Line Bot 訊息給使用者
function sendReplyMessage(CHANNEL_ACCESS_TOKEN, replyToken, replyMessage) {
  var url = "https://api.line.me/v2/bot/message/reply";
  UrlFetchApp.fetch(url, {
    "headers": {
      "Content-Type": "application/json; charset=UTF-8",
      "Authorization": "Bearer " + CHANNEL_ACCESS_TOKEN,
    },
    "method": "post",
    "payload": JSON.stringify({
      "replyToken": replyToken,
      "messages": replyMessage,
    }),
  });
}