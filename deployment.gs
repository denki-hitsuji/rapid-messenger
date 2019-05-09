function test(){
  var userName = "山崎昌宏";
  var emailAddress = "myamazaki08@gmail.com";  
    startCopy(userName, emailAddress);

}
/* 
・使い方
getFeatureName に名前をつけて startCopy 実行します。

・成果物
フォーマットフォルダ内をコピーした新規フォルダを作成し、すべてのURLを記載したテキストを出力します。
ログにテキストのURLを出力します。
*/
function doPost(e){
  var params = JSON.parse(e.postData.getDataAsString());  // ※
  Logger.log("doPost started" + params);
  
  if(params.formType === "fileUploader"){
    Logger.log("file upload");
    uploadImage() ;
    return;
  }
  
  var userName = params.name;
  var emailAddress = params.email;  
  lpUrl = params.lp;

  if(userName.length == 0 || emailAddress.length == 0){
    return HtmlService.createHtmlOutput("<p>failed</p>");
  }
  
  startCopy(userName, emailAddress);
  return HtmlService.createHtmlOutput("<p>succeeded</p>");
  
}

function doGet(e){
  // デプロイメント
  var userName =  e.parameter.name? e.parameter.name : "";
  var emailAddress =  e.parameter.email? e.parameter.email: "";  
  lpUrl = e.parameter.lp;
  
  if(userName.length == 0 || emailAddress.length == 0){
    return HtmlService.createHtmlOutput("<p>failed</p>");
  }
  
  startCopy(userName, emailAddress);
    return HtmlService.createHtmlOutput("<p>succeeded</p>");

}

var lpUrl = "";
function getLpUrl(){
  return lpUrl;
}

function getAppName(){
  return "瞬速メッセンジャー";
}

//この関数を実行する
function startCopy(userName, email) {
  var parentFolder = DriveApp.getFolderById(getFolderId("Parent"));
  var sourceFolder = DriveApp.getFolderById(getFolderId("Source"));
  var createdFolder = parentFolder.createFolder(getAppName() + "_" + userName + "さん専用フォルダ");
  var testFolder = createdFolder.createFolder("テスト用");
  var campaign1Folder = createdFolder.createFolder("キャンペーン1");
  var campaign2Folder = createdFolder.createFolder("キャンペーン2");
  var campaign3Folder = createdFolder.createFolder("キャンペーン3");

  shareFolder(createdFolder, email);

   var requests = [testFolder, campaign1Folder, campaign2Folder, campaign3Folder]
   .map (function (folder) {
     // フォルダの中身をコピー
     folderCopy(sourceFolder, folder);    
     // URL一覧のテキスト作成
     createFileDescribedAllURL(folder);
     // ファイル名に施策名をつける
     setFeatureNameToFiles(folder); 
     modifyTemplates (folder, userName);
    });

}

// フォルダーのIDを取得する　定数管理用
function getFolderId(type){
  switch(type){
    case "Source":
      // フォーマットとなるディレクトリのID
      return "1HX5oj2KJJNzb7LWLo3rZ9JuGJiDh0woI"
    case "Parent":
      // 作成したフォルダを置くディレクトリのID
      return "18CD369zGw-1cl1rydAw7q2evRyxCijAt";
  }
}

function shareFolder(folder, sharedAccount){
  
  var base_folder_id = folder; // 検索対象とするフォルダのID

  folder.addEditor(sharedAccount);
}

// sourceFolder 内のファイルを createdFolder にコピーする
function folderCopy(sourceFolder, createdFolder) {
  var sourceFiles = sourceFolder.getFiles();
  while(sourceFiles.hasNext()) {
    var sourceFile = sourceFiles.next();
    sourceFile.makeCopy(sourceFile.getName(), createdFolder);
  }
}

function getParentFolder(){
  var sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  Logger.log("スプレッドシートID：" + sheetId);
  var file = DriveApp.getFileById(sheetId);
  var thisFolder  = file.getParents().next();
  if(!thisFolder.getParents().hasNext()) return null; 
  
  var parentFolder = thisFolder.getParents().next();
  return parentFolder;
}

function newCampaign(){
  // 新しいキャンペーンフォルダを作る
  var parentFolder = getParentFolder();
  var newFolderName = "新しいキャンペーン";
  var newFolder = parentFolder.createFolder(newFolderName);
  newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);

  
  // ソースファイル一式をコピーする
  var sourceFolder = DriveApp.getFolderById(getFolderId("Source"));
  folderCopy(sourceFolder, newFolder);

  // このファイルから、リストをコピーする
  
  var createdFiles = newFolder.getFiles();  
  while(createdFiles.hasNext()){
    var file = createdFiles.next();
    var fileName = file.getName();
    if(-1 < fileName.search("3.") )
    {
      // 応援者リストなので、リストをコピーする
      copyList(file);
    }
  }
  
}

// このフォルダの値を転記します。
function copyList(destFile){
  var thisSheet = SpreadsheetApp.getActiveSheet();
  
  var listCells = "C2:F200";
  var remarksCells = "AI2:AI200";
  var destSheet = SpreadsheetApp.open(destFile).getSheetByName("リスト");
  var listValues = thisSheet.getRange(listCells).getValues();  
  destSheet.getRange(listCells).setValues(listValues);
  
  var remarks = thisSheet.getRange(remarksCells).getValues();  
  destSheet.getRange(remarksCells).setValues(remarks);
  
}


// フォルダ内にあるファイルのURLを列挙したテキストファイルを作成する
function createFileDescribedAllURL(createdFolder){
  var createdFiles = createdFolder.getFiles();
  var text = "";
  while(createdFiles.hasNext()){
    var file = createdFiles.next();
    var fileName = file.getName();
    text += "・" + fileName + "\n";
    text += file.getUrl() + "\n";
    text += "\n";
  }
  if(text == null){
    return;
  }
  var textFile = createdFolder.createFile("url一覧", text, MimeType.PLAIN_TEXT);
  Logger.log(textFile.getUrl());　// ログにテキストのURLを出力する
}
// フォルダ内にあるファイルに、ユーザーのお名前を入れる
function modifyTemplates (createdFolder, name){
  var createdFiles = createdFolder.getFiles();
  
  while(createdFiles.hasNext()){
    var file = createdFiles.next();
    var fileName = file.getName();
    if(-1 < fileName.search("1.") )
    {
      // 投稿文ファイルなので、名前を挿入
      replaceName(file, name);
      continue;
    }
    if(-1 < fileName.search("2.") )
    {
      // 依頼文ファイルなので、Idを控える
      replaceName(file, name);
      continue;
    }
    if(-1 < fileName.search("3.") )
    {
      // 応援者リストなので、LPのリンクを挿入する
      var lpLink = getLpUrl();
      insertLink(file, lpLink);
      continue;
    }
  }
}

function replaceName(file, name){
  var template = DocumentApp.openById(file.getId());
  var body = template.getBody();
  body.replaceText("〇〇",name);
}

function insertLink(file, lpUrl){
  var spreadsheet = SpreadsheetApp.open(file);
  var cell = spreadsheet.getSheetByName("リスト").getRange("Y2");
  cell.setValue(lpUrl);
}

// フォルダ内にあるファイルの頭に「施策名_」を追加する
function setFeatureNameToFiles(createdFolder){
  var createdFiles = createdFolder.getFiles();
  while(createdFiles.hasNext()){
    var file = createdFiles.next();
    var fileName = file.getName();
    file.setName(fileName + "_" + createdFolder.getName());
  }
}