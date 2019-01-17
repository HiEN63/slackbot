function Main() {
  //スプレッドシートを定義する
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
 
  //シートEventのコピーを作成
  Copysheet(spreadSheet);
  
  var searchUrl = "https://connpass.com/api/v1/event/?";
  //ここに検索条件を付けておくと下のfor文でsearchUrlに連結してくれる
  var searchOption = ["keyword=IT","keyword=大阪","count=100","order=2"] ;
  for(var i = 0;i < searchOption.length; i++){
  
    if(i > 0){
      searchUrl = Concat(searchUrl,"&");
    }
    searchUrl = Concat(searchUrl,searchOption[i]);
  }
  
  try{
     var allEvent = JSON.parse(UrlFetchApp.fetch(searchUrl,
     {
       "muteHttpExceptions" :true,
       "timeout": "240000",
       "useragent" :"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36", 
     }));
  }catch(e){
    Logger.log("connpassAPI is Timeout");
    var delSheet = spreadSheet.getSheetByName("Event のコピー");
    spreadSheet.deleteSheet(delSheet);
    return 1;
  }
  
  //取得したイベント情報をスプレッドシートに書き込む関数　return 0
  Writesheet(allEvent,spreadSheet);
  
 
  var postEvent = [];
  
  //保持しているイベント情報を投稿するかしないか判定する関数 return Slackに投稿する必要のあるイベント情報
  postEvent = checkPostEvent(spreadSheet);
   
  //checkEventの内容をSlackAPI呼び出し関数
  Slackpost(postEvent,'https://hooks.slack.com/services/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');
  
  
  //シートEventのコピーを削除
  var delSheet = spreadSheet.getSheetByName("Event のコピー");
  spreadSheet.deleteSheet(delSheet);
  
}

//------------------------------------------------------------------------------------^-^--------------------------------------

//Eventsシートのコピーを生成する関数
function Copysheet(spreadSheet){

  var objDestSpreadsheet = SpreadsheetApp.openById("1N3on6_duegS3JGhlHs1Y_Mc4PWsUokv6Cu4e0Ivuaig");
  var objSheet = spreadSheet.getSheetByName("Event");
  objSheet.copyTo(objDestSpreadsheet); 
  
}
//文字列連結用の関数　複数回呼び出す
function Concat(FrontLen,BackLen){
  var result = [];
  
  for(var i = 0;i < BackLen.length; i++){
  
    for(var j = 0;j < BackLen[i].length; j++){
    
      FrontLen += BackLen[i][j];
    }
  }
  result.push(FrontLen);
  
  return result;
}

//シートに書き込む関数
function Writesheet(allEvent,spreadSheet){
  
  var wSheet = spreadSheet.getSheetByName("Event");
  var cSheet = spreadSheet.getSheetByName("Event のコピー");
 
  var checkEventid = cSheet.getRange("C2:C101").getValues();
  
  //コピー先のシートのフラグ欄を削除
  var delFlagRange = cSheet.getRange("E2:E101");
  delFlagRange.clearContent();

    wSheet.getRange(1,1).setValue("イベント名");
    wSheet.getRange(1,2).setValue("URL");
    wSheet.getRange(1,3).setValue("イベントID");
    wSheet.getRange(1,4).setValue("開催日時");
    wSheet.getRange(1,5).setValue("検出フラグ");
    
  for(var i = 0; i < 100; i++){

    var newOrOld = 0;
    for (var j = 0;j < 100; j++){
  
      if (allEvent["events"][i]["event_id"] == checkEventid[j]){
         wSheet.getRange(i +2,5).setValue("既出");
         newOrOld = 1;
         break;
      }
    }
      if (newOrOld == 0){
         wSheet.getRange(i +2,1).setValue(allEvent["events"][i]["title"]);
         wSheet.getRange(i +2,2).setValue(allEvent["events"][i]["event_url"]);
         wSheet.getRange(i +2,3).setValue(allEvent["events"][i]["event_id"]);
         wSheet.getRange(i +2,4).setValue(allEvent["events"][i]["started_at"]);
         wSheet.getRange(i +2,5).setValue("新規");
//         wSheet.getRange(i +2,6).setValue("");
    }
  }
  return 0;
}

function checkPostEvent(spreadSheet) {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var alreadyIdSheet = spreadSheet.getSheetByName("AlreadyPostEvent");
  var eventSheet = spreadSheet.getSheetByName("Event");
  
  //二次元配列　n番目のURLは、[n][0] n番目のIDは[n][1]　という表現をする
  var ckPostData = eventSheet.getRange("B2:C101").getValues();
  //Logger.log(ckPostData);

  var alPostId = alreadyIdSheet.getRange("A:A").getValues();
  var alPostRow =alPostId.filter(String).length;
  var postEventUrl = [];
  var postEventId = [];
  
  for (var i = 0; i < 100; i++){
  
    for (var j = 0; j <= alPostRow; j++){
    
      var ckId = ckPostData[i][1].toString();
      var alId = alPostId[j].toString();
      if (ckId == alId){
        break;
      }
      if (j  == alPostRow){
        postEventUrl.push(ckPostData[i][0]);
        postEventId.push(ckPostData[i][1]);
      }
    }
  }
  
  for (var k = 1;k < postEventId.length + 1; k++){
    alreadyIdSheet.getRange(alPostRow + k ,1).setValue(postEventId[k-1]);
  }
  
  //Logger.log(postEventUrl);
 
  return postEventUrl;
}


function Slackpost(postEvent,hookPoint) {
  var cnt = postEvent.length;
  for ( var i = 0; i < cnt; i++ ){
    var payload = {
      "text": postEvent[i], 
      "icon_emoji": ':connpass:',
      "username": 'イベント告知BOT',
      "unfurl_links" : true,
      
    }
    var options = {
      "method" : "POST",
      "payload" : JSON.stringify(payload),
      "headers": {
      "Content-type": "application/json",
      }
    }
    UrlFetchApp.fetch(hookPoint, options);
  }
}