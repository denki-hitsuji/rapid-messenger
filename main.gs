var _ = underscoreGS;

function onInstall(e){ 
//  onOpen(e);
}

function deleteStartRow(){
PropertiesService.getDocumentProperties().deleteProperty("startRow");
}

function onOpen(e) {  
  Logger.log('AuthMode: ' + e.authMode);
  var lang = Session.getActiveUserLocale();
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createAddonMenu();
  if(e && e.authMode == 'NONE'){
    var startLabel = lang === 'ja' ? '使用開始' : 'start';
    menu.addItem(startLabel, 'askEnabled');
  } else {
    if( lang === 'ja')
    {
    menu.addItem('URL短縮', 'generateShortUrls')
      .addItem('配信サンプル生成', 'createFiles')
      .addItem('メール送信', 'sendMails')
      .addSeparator()
      .addSubMenu(
        ui.createMenu("便利機能")
        .addItem('画像アップロード', 'openSidebar')
        .addItem('結果をクリア', 'clearUrls')
        .addItem('新規キャンペーン', 'newCampaign')    
        .addItem('置き換え文字列一覧', 'showKeywords_jp')
       )
      .addSubMenu(
        ui.createMenu("設定")
        .addItem('設定画面', 'showDialog')
        .addItem('初期値を設定', 'defineDefaultProperties')
       );
    }
    else{
    menu.addItem('shorten URL', 'generateShortUrls')
      .addItem('generate doc', 'createFiles')
      .addItem('send emails', 'sendMails')
      .addSeparator()
      .addSubMenu(
        ui.createMenu("utilties")
        .addItem('upload image', 'openSidebar')
        .addItem('clear result', 'clearUrls')
        .addItem('new campaign', 'newCampaign')    
        .addItem('show placeholders', 'showKeywords_en')
       )
      .addSubMenu(
        ui.createMenu("config")
        .addItem('preferences', 'showDialog')
        .addItem('set default', 'defineDefaultProperties')
       );
    
    }
//    var userProps = PropertiesService.getUserProperties();
//    var setDefault = userProps.getProperty("willSetDefault");
//    if(setDefault == 1){
      //setDefaultIfBlank();
//    }
  };
  menu.addToUi();

};

function showKeywords(htmlFile, menuTitle, colWord ){
  var replacers = getReplacingWords();
  var replacingObj = _._map(replacers, function(rep, key){return "<label>" + rep + ' (' + key + colWord + ')</label> <input type="text" class="float-right"' 
    + 'value="{' + rep + '}" />'}, null);
  var replacingTexts = replacingObj.toString().replace(/,/g,'<br/>');
  Logger.log(replacingTexts);
  var body =  HtmlService.createTemplateFromFile(htmlFile).evaluate().getContent().replace('<hr/>', replacingTexts + '<hr/>');
  var html = HtmlService.createHtmlOutput(body)
      .setWidth(550)
      .setHeight(550);
  SpreadsheetApp.getUi()
      .showModalDialog(html, menuTitle);
}

function showKeywords_jp(){
  showKeywords('keywords.html', '置き換え文字列一覧', '列' );
}

function showKeywords_en(){
  showKeywords('keywords_en.html', 'Placeholder List', 'column' );
}


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
  var newUrlCol = getNewUrlCol(); //AA;
  var docIdCol = getNewIdCol(); //AH
  var lastRow = getLastRowNumber(NAME_COL); 
//  Browser.msgBox ("newUrlCol OK " + newUrlCol + " " + docIdCol);
//  Logger.log("lastRow:" + lastRow);

  // ヘッダー行を除いて２行目からデータを取得
  for(var i=2;i<lastRow;i++){
    var id = getRange(mySheet, i,docIdCol).getValue(); 
    if( id.length === 0 ) continue;
    var removingDoc = DriveApp.getFileById(id);

    if(removingDoc != null || !removingDoc.isTrashed()) {
      removingDoc.setTrashed(true);
    }
//    var newDocument = DriveApp.removeFile(id);
  }
//  Browser.msgBox ("removeFile OK");

  getRange(mySheet, 2, newUrlCol).offset(0,0,lastRow - 1).clearContent();
  getRange(mySheet, 2, docIdCol).offset(0,0,lastRow - 1).clearContent();

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

function getLastRowNumber(col){
　var mySheet=SpreadsheetApp.getActiveSheet(); //シートを取得
  var last_row = mySheet.getLastRow();
  const sheet = SpreadsheetApp.getActiveSheet(); 
  const columnVals = getRange(mySheet, 1, col).offset(0,0, last_row).getValues();
  const lastRow = columnVals.filter(String).length;  //空白を除き、配列の数を取得

  return lastRow;  
}

function createFiles(){  
  var newFolder = createNewFolder();
  
  // 短縮URLが、最終行まで入っているか判定し、なければ再生成（追記したときにこうなる）
  var mySheet=SpreadsheetApp.getActiveSheet(); //シートを取得
  var newUrlCol = getNewUrlCol(); 
  var nameLastRow = getLastRowNumber(NAME_COL);
  var lastShortUrl = getRange(mySheet, nameLastRow ,newUrlCol).getValue();　// 短縮URL

  if(lastShortUrl == "undefined"){
    clearUrls();
  }
  if(lastShortUrl.length == 0){
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
  var ui = SpreadsheetApp.getUi();
  
  const mailAddressCol = getMailAddressCol();
  const howManyReceivers = getLastRowNumber(mailAddressCol) -1; 
  var button = ui.alert('送信数:' + howManyReceivers + '通\nメールを送信しますか？', ui.ButtonSet.OK_CANCEL);
  if( button != ui.Button.OK) return;

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
      .replaceText("{ドキュメント}",fileUrl)
    } else {
      newBody=body
      .replaceText("{nickname}",nickname)
      .replaceText("{docUrl}",fileUrl);
    }
    // リンクを編集
    var mailBody = replaceLinks(newBody, fileUrl).getText();
    
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

function replaceLinks(body, url){
    var urlLink = null;
    
    while (urlLink = body.findText(url, urlLink)){
      link(urlLink, url);
    }
    return body;
}

function link(urlLink, url)
{
  urlLink.getElement().asText().setLinkUrl(null);
  Logger.log("リンクを設定します:" + urlLink.getElement().asText().getText());
  Logger.log("IsPartial? :" + urlLink.isPartial());
  const startOffset = urlLink.getStartOffset();
  const endOffsetInclusive = urlLink.getEndOffsetInclusive();
  const length = urlLink.getElement().asText().getText().length;
  Logger.log("start:" + startOffset + " end:" + endOffsetInclusive + " length:" + length);
  
  urlLink.getElement().asText().setLinkUrl(startOffset, endOffsetInclusive, url);
}

function generateShortUrls(){
  Logger.log("スプレッドシート内のURLを短縮します。");
  setDefaultIfBlank();  

  /* スプレッドシートのシートを取得と準備 */
  var mySheet=SpreadsheetApp.getActiveSheet(); //シートを取得
  var endRow = getLastRowNumber( NAME_COL); // 名前の最終行を取得
  for(var i = 2 ; i <= endRow; i++ ){
    var personName = getRange(mySheet,i,NAME_COL).getValue().replace(" ", "");
    if (personName.length == 0){
      Logger.log("終了します。");
      break;
    }
  
    var newUrlCol = getNewUrlCol();
    var originUrlCol = getOriginUrlCol();
    
    var newUrl = getRange(mySheet,i,newUrlCol).getValue();
    if(newUrl === "undefined"){
      getRange(mySheet,i,newUrlCol).setValue("");
    } else if( newUrl != ""){
      continue;
    }
    
    var originUrl = getRange(mySheet,i,originUrlCol).getValue();　// 元のURL
    Logger.log(originUrl);
    //Browser.msgBox("originUrl =" + originUrl);
    var shortUrl =  shorten(originUrl).replace('http://','https://');
    getRange(mySheet,i,newUrlCol).setValue(shortUrl);
  }
}

function shorten(originUrl){  
   var token = PropertiesService.getDocumentProperties().getProperty("bitly_token");
   var url = "https://api-ssl.bitly.com/v3/shorten?access_token=" + token + "&longUrl=" + encodeURIComponent(originUrl.trim());
   console.log(url);
   var responseApi = UrlFetchApp.fetch(url);
   var responseJson = JSON.parse(responseApi.getContentText());
   console.log("URL shortened:" + responseJson["data"]["url"]);
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

function resizeImage(image){
  var width = image.getWidth();
  var newW = width;
  var height = image.getHeight();
  var newH = height;
  var ratio = width/height
  
  if(width>290){
    newW = 290;
    newH = parseInt(newW/ratio);
  }  
  image.setWidth(newW).setHeight(newH);

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