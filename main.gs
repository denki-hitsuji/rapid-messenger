function onInstall(e){ 
//  onOpen(e);
}

function deleteStartRow(){
PropertiesService.getDocumentProperties().deleteProperty("startRow");
}

function onOpen(e) {  
  Logger.log('AuthMode: ' + e.authMode);
  var lang = Session.getActiveUserLocale();
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  if(e && e.authMode == 'NONE'){
    var startLabel = lang === 'ja' ? '使用開始' : 'start';
    menu.addItem(startLabel, 'askEnabled');
  } else {
    if( lang === 'ja')
    {
    menu.addItem('URL短縮', 'generateShortUrls')
      .addItem('配信サンプル生成', 'createFiles')
      .addItem('メール送信', 'sendMails')
      .addItem('結果をクリア', 'clearUrls')
      .addItem('新規キャンペーン', 'newCampaign')    
      .addItem('設定', 'showDialog');
//    var userProps = PropertiesService.getUserProperties();
//    var setDefault = userProps.getProperty("willSetDefault");
//    if(setDefault == 1){
      menu.addItem('初期値設定', 'defineDefaultProperties');
      menu.addItem('置き換え文字列一覧', 'showKeywords');
    }
    else{
    menu.addItem('shorten URL', 'generateShortUrls')
      .addItem('generate doc', 'createFiles')
      .addItem('send emails', 'sendMails')
      .addItem('clear result', 'clearUrls')
      .addItem('new campaign', 'newCampaign')    
      .addItem('config', 'showDialog');
//    var userProps = PropertiesService.getUserProperties();
//    var setDefault = userProps.getProperty("willSetDefault");
//    if(setDefault == 1){
      menu.addItem('set default', 'defineDefaultProperties');
      menu.addItem('show placeholders', 'showKeywords_en');
    }
//      menu.addItem('新規プロジェクト', 'defineDefaultProperties');
//    }
    //setDefaultIfBlank();
  };
  menu.addToUi();

};

function showKeywords(){
  var html = HtmlService.createTemplateFromFile('keywords.html').evaluate()
      .setWidth(450)
      .setHeight(300);
  SpreadsheetApp.getUi()
      .showModalDialog(html, '置き換え文字列一覧');
}


function showKeywords_en(){
  var html = HtmlService.createTemplateFromFile('keywords_en.html').evaluate()
      .setWidth(450)
      .setHeight(300);
  SpreadsheetApp.getUi() 
      .showModalDialog(html, 'Placeholder List');}

function setDefaultIfBlank(){
  var isDocPartnerList = SpreadsheetApp.getActive().getName().match(/^3./);
  if(!isDocPartnerList){
    return;
  }
  
  var docProps = PropertiesService.getDocumentProperties();
  var token = docProps.getProperty("bitly_token");
  var isBlank = token == null || token.length === 0; 
  if(isBlank){
    defineDefaultProperties();
  }
}

function log_WillSetDefault(){
  var userProps = PropertiesService.getUserProperties();
  var val = userProps.getProperty("willSetDefault");
  Logger.log("willSetDefault:" + val);
}

function set_WillSetDefault_True(){
  var userProps = PropertiesService.getUserProperties();
//  userProps.deleteProperty("willSetDefault");
  userProps.setProperty("willSetDefault", 1);
}

function defineDefaultProperties(){
  var props = PropertiesService.getScriptProperties();
  logProperties();
  setDefaultProperty("bitly_token", props.getProperty("bitly_token"));
  setDefaultProperty("originUrlCol", "Z");
  setDefaultProperty("newUrlCol", "AA");
  setDefaultProperty("newIdCol", "AH");
  setDefaultProperty("nicknameCol", "D");
  setDefaultProperty("mailAddressCol", "E");
  setDefaultProperty("mailTemplate", getTemplateId(/^2./));
  setDefaultProperty("templateDocId", getTemplateId(/^1./));
}

function getTemplateIdTest(){
  Logger.log('templateDocId: ' + getTemplateId(/^1./));
  Logger.log("mailTemplate: " + getTemplateId(/^2./));
}

function getTemplateId(title){
  
  var ss = SpreadsheetApp.getActive();
  var ssid =ss.getId();
  Logger.log("active spreadsheet id : " + ssid);
  var ssFile = DriveApp.getFileById(ssid);
  if(ssFile == null) 
    return null;
   
  Logger.log("file got ");
  var parents = ssFile.getParents();
  Logger.log("parents hasNext? : " + parents.hasNext());

  if(!parents.hasNext()) return null;
  
  var folder = parents.next();
  Logger.log("parent folder id : " + ss.getId());
  var files = folder.getFiles();
  while(files.hasNext()){
    var file = files.next();
    var fileName = file.getName();
    if(fileName.match(title)){      
      Logger.log("file found name : " +fileName);
      return file.getId();
    }
  }
  
  return null;
}

function setDefaultProperty(key, defaultValue){
  var props = PropertiesService.getDocumentProperties();
//  if(props.
  var prop = props.getProperty(key);
  Logger.log("property got " + key + ":" + prop);
  if(prop == null || prop === "undefined" || prop.length === 0){
    props.setProperty(key, defaultValue);
    Logger.log("property set " + key + ":" + defaultValue);
  }
}


function askEnabled(){
  var lang = Session.getActiveUserLocale();
  var title = 'Your Script\'s Title';
  var msg = lang === 'ja' ? '瞬速メッセンジャーが有効になりました。ブラウザを更新してください。' : 'Rapid Messenger has been enabled.';
  var ui = SpreadsheetApp.getUi();
  ui.alert(title, msg, ui.ButtonSet.OK);
};

function clearUrls(){
  var mySheet=SpreadsheetApp.getActiveSheet(); //シートを取得

//  Browser.msgBox ("endRow OK");
  var newUrlCol = getNewUrlCol(); //24;
  var docIdCol = getNewIdCol();
  var lastRow =mySheet.getDataRange().getLastRow(); //シートの使用範囲のうち最終行を取得
//  Browser.msgBox ("newUrlCol OK " + newUrlCol + " " + docIdCol);
  

  for(var i=2;i<=lastRow;i++){
    var id = getRange(mySheet, i,docIdCol).getValue(); 
    if( id.length === 0 ) continue;
    var removingDoc = DriveApp.getFileById(id);

    if(removingDoc != null || !removingDoc.isTrashed()) {
      removingDoc.setTrashed(true);
    }
//    var newDocument = DriveApp.removeFile(id);
  }
//  Browser.msgBox ("removeFile OK");

  getRange(mySheet, 2, newUrlCol).offset(0,0,lastRow - 1, 0).clearContent();
  getRange(mySheet, 2, docIdCol).offset(0,0,lastRow - 1, 0).clearContent();

}

function showDialog() {
  var html = HtmlService.createTemplateFromFile('setting.html').evaluate()
      .setWidth(400)
      .setHeight(400);
  SpreadsheetApp.getUi() 
      .showModalDialog(html, '設定');
}

function saveSettings(e){
  var props = PropertiesService.getDocumentProperties();
  props.setProperty("bitly_token", e.bitly_token);
  props.setProperty("mailTemplate", e.mailTemplate);
  props.setProperty("templateDocId", e.templateDocId);
  props.setProperty("newUrlCol", e.newUrlCol);
  props.setProperty("newIdCol", e.newIdCol);
  props.setProperty("originUrlCol", e.originUrlCol);
  props.setProperty("nicknameCol", e.nicknameCol);
  props.setProperty("mailAddressCol", e.mailAddressCol);
  var userProps = PropertiesService.getUserProperties();
  userProps.setProperty("userName", e.userName);

  logProperties();
}

function test(){
  var psid = "";
  postMessage(psid) ;
}

//function addProperty(){
//  var key = "PAGE_ACCESS_TOKEN";
//  var value = "EAACOrw8r3ykBAKK2NHp54f92auUkZBZALiR5HaUmmnVACiX8l8eV3AnhlUETb1naLjy87ZBGXjaBOVcQPW8WZC2y8duAXdt76eBZCMWeMZAksg5vTuY1WRKqdxCxZBgfkOoQSNhCCOisjGP0uMvu4AS8KYnSiAscwz3Hd1Nw8YoDU4jy8xJ2f1N";
//  PropertiesService.getScriptProperties().setProperty(key, value);
//}

function createFiles(){  
  var newFolder = createNewFolder();
  
  var mySheet=SpreadsheetApp.getActiveSheet(); //シートを取得
  var newUrlCol = getNewUrlCol(); 
  var shortUrl = getRange(mySheet,2,newUrlCol).getValue();　// 短縮URL

  if(shortUrl == "undefined"){
    clearUrls();
  }
  if(shortUrl.length == 0){
    generateShortUrls();
  }
  generateFiles(newFolder);
}



function isAllLetter(inputtxt)
  {
   var letters = /^[A-Za-z]+$/;
   if(inputtxt.match(letters))
     {
      return true;
     }
   else
     {
     return false;
     }
  }

function isNumber(x){ 
//  Browser.msgBox ("tyoe:" + typeof(x) );

  if( typeof(x) != 'number' && typeof(x) != 'string' )
  return false;
  else 
    return (x == parseFloat(x) && isFinite(x));
}

function getRange(sheet, row, col){
  if(col == null){throw new RangeError("アドレスが不正です:" + col + row )};
//  Browser.msgBox ("getRange start");
  if(isNumber(col)){
    Logger.log(row + "," + col);
//    Browser.msgBox ("getRange isNumber=true OK");
    return sheet.getRange(row, col);
  }
//  Browser.msgBox ("getRange isNumber=false OK");
  
  if(isAllLetter(col)) {
    var a1 = col + row;
    Logger.log(a1);
//    Browser.msgBox ("getRange isAllLetter=true OK");    
    return sheet.getRange(a1);
  }
  else {throw new RangeError("アドレスが不正です:" + col + row ) };
}

function setPropertyAsTest(){
  PropertiesService.getDocumentProperties().setProperty("test1", "foo");
  PropertiesService.getDocumentProperties().setProperty("test2", "bar");
  PropertiesService.getDocumentProperties().setProperty("test3", "baz");

  logProperties();
}

function logProperties(){
  var props = PropertiesService.getDocumentProperties();
  var keys = props.getKeys();
  Logger.log("keys:" + keys.join(","));
  keys.forEach(function(key) {
    Logger.log(key + ": " + props.getProperty(key));
  });

}

function mailAddressTest(){
  var address = "-";
  Logger.log(isValidMailAddress(address));
  
  var address2 = "test@example.com";
  Logger.log(isValidMailAddress(address2));
}

function isValidMailAddress(mailAddress){
  var regex = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
  var hit = mailAddress.match(regex);
  return (hit != null);
}

function sendMails(){
  /* スプレッドシートのシートを取得と準備 */
  var mySheet=SpreadsheetApp.getActiveSheet(); //シートを取得
  var endRow=mySheet.getDataRange().getLastRow(); //シートの使用範囲のうち最終行を取得
  var userName = PropertiesService.getUserProperties().getProperty("userName");
  
  /* メールテンプレートは独立した文書 */
  // 紹介依頼メールテンプレート
  logProperties();
  var mailTemplateId= PropertiesService.getDocumentProperties().getProperty("mailTemplate");
  Logger.log("mailTmpl:" + mailTemplateId);
  var mailTemplate = DocumentApp.openById(mailTemplateId);
  var title = getMailTemplateTitle(mailTemplate);

  for(var i=2;i<=endRow;i++){
    var personName = getRange(mySheet, i,NAME_COL).getValue().replace(" ", "");
    if (personName.length == 0){
      break;
    }

    var docIdCol = getNewIdCol();
    var nicknameCol = getNicknameCol();
    var mailAddressCol = getMailAddressCol();
    
    var documentId = getRange(mySheet, i,docIdCol).getValue(); // ドキュメントID 
    var fileUrl = DriveApp.getFileById(documentId).getUrl();
    var nickname = getRange(mySheet,i,nicknameCol).getValue();　//メール内呼称
    var emailAddress = getRange(mySheet,i,mailAddressCol).getValue();　
    if (documentId.length == 0 || nickname.length == 0 ||  emailAddress.length == 0){
      continue;
    }
    
    if(!isMailAddressValid(emailAddress)){
      continue;
    }

    var newBody = "";

    var body = getMailTemplateBody(mailTemplate);
    // 新しい本文を生成 (ここで置換を全部やる)
    var lang = Session.getActiveUserLocale();
    if(lang　=== 'ja'){
      newBody=body
      .replaceText("{お名前}",nickname)
      .replaceText("{ドキュメント}",fileUrl);
    } else {
      newBody=body
      .replaceText("{nickname}",nickname)
      .replaceText("{docUrl}",fileUrl);
    }
    // リンクを編集
    var mailBody = replaceLink(newBody, fileUrl).getText();
    
    var to = emailAddress;
    var subject = "";
    
    if(lang　=== 'ja'){
      subject = title
      .replace(/{お名前}/,nickname)
      .replace(/{ドキュメント}/,fileUrl);    
    } else {
      subject = title
      .replace(/{nickname}/,nickname)
      .replace(/{docUrl}/,fileUrl);          
    }    

    // 生成した本文をメールで送信  
    GmailApp.sendEmail(
      to,
      subject,
      mailBody,
      {
      name: userName
    }); //MailAppではfromが設定できないとのこと
    Logger.log(mailSent + to); //ドキュメントの内容をログに表示
//     Logger.log(newBody.getText());
  }

  var mailSent = lang　=== 'ja' ? "メールを送信しました。" : "email(s) sent.";
  SpreadsheetApp.getUi().alert(mailSent);

}

function isMailAddressValid(mailAddress)
{
   var mailFormat = /^[a-zA-Z0-9.!#$%&'*+\/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*$/;
   return mailAddress.match(mailFormat);
}

function replaceLink(body, url){
    var urlLink = null;
    
    while (urlLink = body.findText(url, urlLink)){
      var originUrl = urlLink.getElement().asText().getLinkUrl();
//      if(originUrl == null && urlLink.isPartial()){        
//        continue;
//      }
      Logger.log("リンクを設定します:" & urlLink.getElement().asText());
      
      urlLink.getElement().asText().setLinkUrl(url);
    }
    return body;
}

function generateShortUrls(){
  Logger.log("スプレッドシート内のURLを短縮します。");
  setDefaultIfBlank();  

  /* スプレッドシートのシートを取得と準備 */
  var mySheet=SpreadsheetApp.getActiveSheet(); //シートを取得
  var endRow = mySheet.getDataRange().getLastRow(); //シートの使用範囲のうち最終行を取得
  for(var i = 2 ; i <= endRow; i++ ){
    var personName = getRange(mySheet,i,NAME_COL).getValue().replace(" ", "");
    if (personName.length == 0){
      Logger.log("終了します。");
      break;
    }
  
    var newUrlCol = getNewUrlCol();
    var originUrlCol = getOriginUrlCol();
    
    var newUrl = getRange(mySheet,i,newUrlCol).getValue();
    if( newUrl != ""){
      continue;
    }
    
    var originUrl = getRange(mySheet,i,originUrlCol).getValue();　// 元のURL
    Logger.log(originUrl);
    //Browser.msgBox("originUrl =" + originUrl);
    var shortUrl =  shorten(originUrl);
    getRange(mySheet,i,newUrlCol).setValue(shortUrl);
  }
}

function shorten(originUrl){
   //var url = UrlShortener.Url.insert({longUrl: originUrl});
   //return url.id; 
  
   var token = PropertiesService.getDocumentProperties().getProperty("bitly_token");
   var url = "https://api-ssl.bitly.com/v3/shorten?access_token=" + token + "&longUrl=" + encodeURIComponent(originUrl);
   var responseApi = UrlFetchApp.fetch(url);
   var responseJson = JSON.parse(responseApi.getContentText());
   return responseJson["data"]["url"];
}

function letterToColumn(letter)
{
  if(isNumber(letter)) return letter;
  
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
  {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

function generateFiles(folder){
  //◆開始時刻を取得
  var startTime = new Date();
  Logger.log("startTime:" + startTime);
  var props = PropertiesService.getDocumentProperties();

  //◆何行目まで処理したかを保存するときに使用するkey
  var resumeSheetKey = 'resumeSheet';
  var activeSsheetName = SpreadsheetApp.getActiveSheet().getSheetName();
 // Browser.msgBox(sheetName);
 // return;
  props.setProperty(resumeSheetKey, activeSsheetName);
  var startRowKey = "startRow";
  var triggerKey = "trigger";
  var startRow = parseInt(props.getProperty(startRowKey));
  if(!startRow){
    //初めて実行する場合はこっち。!startRow　は、startRowが0（空）の時。↓で初期値（始める行数）を設定
    startRow = 2;
  }
  
  /* スプレッドシートのシートを取得と準備 */
  // 中断に備えて、処理中のシート名を控えておく
  var sheetName = props.getProperty(resumeSheetKey);
  var mySheet=SpreadsheetApp.getActive().getSheetByName(sheetName); 
  var endRow=mySheet.getDataRange().getLastRow(); //シートの使用範囲のうち最終行を取得
  var docIdCol = getNewIdCol();

  /* テンプレートは独立した文書で、ひとつだけ */
  var strDocUrl= props.getProperty("templateDocId"); //ドキュメントのURL
  var templateFile = DriveApp.getFileById(strDocUrl); 
    
  /* シートの全ての行について姓名を差し込みファイルを生成*/
  for(var i=startRow;i<=endRow;i++){
    
     //◆開始時刻（startTime）とここの処理時点の時間を比較する 
    var passedHalfMinutes = parseInt((new Date() - startTime) / (1000 * 30)); 
    if(8 <= passedHalfMinutes){
      //4分経過していたら処理を中断
      Logger.log("Aborting on: Row " + i);

      //何行まで処理したかを保存　参考:http://tbpgr.hatenablog.com/entry/2016/12/20/233349
      //props.setProperty(startRowKey, i);
      //トリガーを発行
      setTrigger(triggerKey, "createFiles");
      return;
    }
    
    var personName = getRange(mySheet,i,NAME_COL).getValue().replace(" ", "");
    
    // 名前が途切れたところで処理終了
    if (personName.length == 0){
      deleteStartRow();
      break;
    }

    var newUrlCol = getNewUrlCol(); 
    var shortUrl = getRange(mySheet,i,newUrlCol).getValue();　// 短縮URL
    // 短縮URLがない場合は処理終了
    if (shortUrl.length == 0 || shortUrl == "undefined"){
      deleteStartRow();
      break;
    }

    var existingDocument = getRange(mySheet,i,docIdCol).getValue();
    // 行をスキップする基準は、1.すでにドキュメントが生成済み 2.短縮URLが生成されてない
    if(0 == shortUrl.length){
      // Browser.msgBox(i + "行目はスキップ");
      continue;
    }
    else if(0 < existingDocument.length)
    {
      //  Browser.msgBox(i + "行目は生成済み");
      continue;
    }
    
    Logger.log(i + "行目:" + shortUrl + " " + shortUrl.length );
 
    
    var number = "000" + getRange(mySheet,i,1).getValue();
    number = number.substring(number.length - 3);    
    
    var lang = Session.getActiveUserLocale();
    var fileName = number + "_" + personName + lang　=== 'ja' ? "さん" : "";

    var newFile = templateFile.makeCopy(fileName, folder);
    var docIdCell = getRange(mySheet,i,docIdCol);

    docIdCell.setValue('=HYPERLINK("' + newFile.getUrl() + '","' + newFile.getId() + '")' ); // 新しいドキュメントIDを控えておく

    var newDocument=DocumentApp.openById(newFile.getId()); //ドキュメントをIDで取得
    var body = newDocument.getBody();
    
    var nicknameCol = getNicknameCol();
    var partnerName =  getRange(mySheet,i,NAME_COL).getValue();　// 紹介者名
    var nickname = getRange(mySheet,i,nicknameCol).getValue();　//メール内呼称
//    var strMessage =mySheet.getRange(i,3).getValue();　//メッセージ 

    // 新しい本文を生成 (ここで置換を全部やる)
    var newBody = "";
    var lang = Session.getActiveUserLocale();
    if(lang　=== 'ja'){
      newBody=body
      .replaceText("{紹介者}",partnerName)
      .replaceText("{紹介者呼称}",nickname)
      .replaceText("{短縮URL}",shortUrl);
    } else {
      newBody=body
      .replaceText("{partnerName}",partnerName)
      .replaceText("{nickname}",nickname)
      .replaceText("{shortUrl}",shortUrl);
    }

    // リンクを編集
    replaceLink(newBody, shortUrl);
    
    Logger.log("書き込みました：" + newBody); //ドキュメントの内容をログに表示
 
  }


}

//◆指定したkeyに保存されているトリガーIDを使って、トリガーを削除する
function deleteTrigger(triggerKey) {
var triggerId = PropertiesService.getDocumentProperties().getProperty(triggerKey);
 
if(!triggerId) return;
 
ScriptApp.getProjectTriggers().filter(function(trigger){
return trigger.getUniqueId() == triggerId;
})
.forEach(function(trigger) {
ScriptApp.deleteTrigger(trigger);
});
PropertiesService.getDocumentProperties().deleteProperty(triggerKey);
  Logger.log("Trigger deleted: " + triggerId);

}
 
//◆トリガーを発行。トリガーを発行した箇所から
function setTrigger(triggerKey, funcName){
 
  //保存しているトリガーがあったら削除
  deleteTrigger(triggerKey);
  var dt = new Date();
  
  //再実行
  dt.setTime(dt.getTime() + (1000 * 40));
//  dt.setMinutes(dt.getMinutes() + 1);
  var triggerId = ScriptApp.newTrigger(funcName).timeBased().at(dt).create().getUniqueId();
  
  Logger.log("Trigger set: " + triggerId);
  
  //あとでトリガーを削除するためにトリガーIDを保存しておく
  PropertiesService.getDocumentProperties().setProperty(triggerKey, triggerId);
}

// 水平線の手前まで、ドキュメントの内容を取得します。
function getTemplateSection(document){
    var searchTypeParagraph = DocumentApp.ElementType.PARAGRAPH;
    var searchTypeHR = DocumentApp.ElementType.HORIZONTAL_RULE;
    var body = document.getBody();
    var firstHR = body.findElement(searchTypeHR);
    var templateBody = "";
    var templateParagraph = null;
    var theHr = firstHR.getElement();
    Logger.log(theHr.getParent());
    
    // 水平線(HR)の前までがテンプレート。これを取得して値を差し込み、新しい本文を作る
    while (templateParagraph = body.findElement(searchTypeParagraph, templateParagraph)){
      var theParagraph = templateParagraph.getElement().asParagraph();
      Logger.log("既存のテンプレート：" + theParagraph.getText());
//      Logger.log("水平線前の段落" + theHr.getParent().getText());

      if(body.getChildIndex( theHr.getParent()) <body.getChildIndex(theParagraph)){
        Logger.log("水平線が見つかりました。" );
        break;
      }
      templateBody += theParagraph.getText() + "\n"; //最初の段落の内容を取得
    }
    Logger.log("テンプレート全体： " + templateBody); 
    return templateBody;
}

function createNewFolder(){
  // 出力先のフォルダを生成
  Logger.log("出力先のフォルダを生成");
  
  var sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  Logger.log("スプレッドシートID：" + sheetId);
  var file = DriveApp.getFileById(sheetId);
  var thisFolder  = file.getParents().next();
  var today = new Date(); 
  var dateString = "";
  dateString += today.getFullYear() + "-";
  dateString += (today.getMonth() + 1) + "-";
  dateString += today.getDate();
  
  while(thisFolder.getFoldersByName(dateString).hasNext()){
    var child = thisFolder.getFoldersByName(dateString).next();
    return child;
  }
  
  var newFolder = thisFolder.createFolder(dateString);
  newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);

  Logger.log("生成しました：" + dateString);
  
  return newFolder;
}

// 過去の出力結果を削除します。
function removeOldOutput(document){

  

}