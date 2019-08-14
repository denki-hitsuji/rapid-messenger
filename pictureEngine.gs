/* pictureEngine 
画像アップロードに関する処理を提供します。
関連ファイル：upload.html+upload_js.html(アップローダー／ビューワー)
*/


function openSidebar(){
  var htmlOutput = 
 HtmlService.createTemplateFromFile('upload').evaluate();
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function createImageFolder(){
  // 出力先のフォルダを生成
  Logger.log("出力先のフォルダを生成");
  
  var sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  Logger.log("スプレッドシートID：" + sheetId);
  var file = DriveApp.getFileById(sheetId);
  var thisFolder  = file.getParents().next();
  var folderName = "Image";
  
  while(thisFolder.getFoldersByName(folderName).hasNext()){
    var child = thisFolder.getFoldersByName(folderName).next();
    return child;
  }
  
  var newFolder = thisFolder.createFolder(folderName);
  newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);

  Logger.log("生成しました：" + folderName);
  
  return newFolder;
}

// 画像をアップロードします。
function uploadImage(formObject) {
  Logger.log("processing form");
  Logger.log("myFile is:" + formObject.myFile);

  if(typeof(formObject.myFile) == "undefined" || formObject.myFile.size == 0){
    return;
  }
  
  var formBlob = formObject.myFile;
  var imgFolder = createImageFolder();
    
  var driveFile =imgFolder.createFile(formBlob);
//    var driveFile = DriveApp.createFile(formBlob);
  //  var image = {myfile:formBlob};
  
  if (formBlob){
    var fileId = driveFile.getId();
    var rowNum = formObject.rowNum;
    Logger.log("rowNum:" + rowNum);
    savePictureId(rowNum, fileId);
    return fileId;
  }else{
    return "";
  }
}

function getRowById(sheet, id){
  Logger.log("getRowById:start");
  var idCol = 1;
  var row = 2;
  var tf = sheet.createTextFinder(id);
  var cell = tf.findNext();
  if(cell == null) return null;
  Logger.log("found:" + cell.getA1Notation());
  while(cell.getColumn() != idCol){
    cell = tf.findNext();
  }
  Logger.log("returned:" + cell.getA1Notation());
  return cell.getRow();
}

function getPictureCell(rowId){
    // imgフォルダから、行番号に相当するファイルを取得
  // ファイルのIDを取り出す
  // ファイルのIDを使って、イメージファイルを取得
  // タグにして貼り付け
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = getRowById(sheet, rowId);
  if(row == null)  return null;

  var pictureCol = "AL"
  return sheet.getRange(pictureCol + row);
}

function getPictureId　(rowId) {
  var cell = getPictureCell(rowId);
  if(cell) return cell.getValue();
}

function savePictureId　(rowId, pictureId) {
  var cell = getPictureCell(rowId);
  if(cell) cell.setValue(pictureId);
}

var thisSheet;
function getThisSheet(){
if( thisSheet == null)
  thisSheet = SpreadsheetApp.getActiveSheet();
  
 return thisSheet;
}

function getPartnerName(rowId){
  // 名前を取得
  // ファイルのIDを取り出す
  // ファイルのIDを使って、イメージファイルを取得
  // タグにして貼り付け
  var row = getRowById(getThisSheet(), rowId);
  //Logger.log("row is " + row);

  var personName = getRange(getThisSheet(), row, NAME_COL).getValue();
  Logger.log("name is " + personName);
  return personName; 
}


