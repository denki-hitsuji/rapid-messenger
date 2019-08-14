var _ = underscoreGS;

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

// 水平線の手前まで、ドキュメントの内容を取得します。
function getMailTemplateTitle(document){
    var searchTypeParagraph = DocumentApp.ElementType.PARAGRAPH;
    var searchTypeHR = DocumentApp.ElementType.HORIZONTAL_RULE;
    var body = document.getBody();
    var firstHR = body.findElement(searchTypeHR);
    var templateTitle = "";
    var templateParagraph = null;
    var theHr = firstHR.getElement();
    Logger.log(theHr.getParent());
    
    // 水平線(HR)の前までがタイトル。これを取得して値を差し込み、新しい本文を作る
    while (templateParagraph = body.findElement(searchTypeParagraph, templateParagraph)){
      var theParagraph = templateParagraph.getElement().asParagraph();
      Logger.log("タイトル：" + theParagraph.getText());
//      Logger.log("水平線前の段落" + theHr.getParent().getText());

      if(body.getChildIndex( theHr.getParent()) < body.getChildIndex(theParagraph)){
        Logger.log("水平線が見つかりました。" );
        break;
      }
      templateTitle += theParagraph.getText() + "\n"; //最初の段落の内容を取得
    }
    Logger.log("テンプレートタイトル： " + templateTitle); 
    return templateTitle;
}


function getMailTemplateBody(document){
    // 水平線以降の文を取得します。
    var searchTypeParagraph = DocumentApp.ElementType.PARAGRAPH;
    var searchTypeHR = DocumentApp.ElementType.HORIZONTAL_RULE;

    var body = document.getBody().copy();
    var firstHR = body.findElement(searchTypeHR);
    var theHr = firstHR.getElement();
    Logger.log(theHr.getParent());

    // 水平線の後が本文。古い本文を消し、新しい本文を追加する  
    var searchResult = null;
    while (searchResult = body.findElement(searchTypeParagraph, searchResult)) {
      var theParagraph = searchResult.getElement().asParagraph();
      if(body.getChildIndex(theParagraph) < body.getChildIndex(theHr.getParent())){
        Logger.log("水平線が見つかりました。" );
        body.removeChild(theParagraph);        
        break;
      }
    }
    return body;
}

/*
-----------------------------------------------------------
アプリの最重要処理
-----------------------------------------------------------
*/
// 配信テンプレートを生成する
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
  var endRow = getLastRowNumber(NAME_COL); 　//シートの名前の最終行を取得
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
    var fileName = number + "_" + personName + "さん" ;

    var newFile = templateFile.makeCopy(fileName, folder);
    var docIdCell = getRange(mySheet,i,docIdCol);
    var fileId = newFile.getId();
    

    docIdCell.setValue('=HYPERLINK("' + newFile.getUrl() + '","' + fileId + '")' ); // 新しいドキュメントIDを控えておく

    var newDocument = DocumentApp.openById(fileId); //ドキュメントをIDで取得
//    DocumentApp.
    if(newDocument == null) {
      Logger.log("newDocument == null")
    };
    
    var body = newDocument.getBody();
    
    var nicknameCol = getNicknameCol();
    var partnerName =  getRange(mySheet,i,NAME_COL).getValue();　// 紹介者名
    var nickname = getRange(mySheet,i,nicknameCol).getValue();　//メール内呼称
//    var strMessage =mySheet.getRange(i,3).getValue();　//メッセージ 
    var imageIdCol = "AL";
 
    var imageId = getRange(mySheet, i, imageIdCol).getValue();
    const imageType =DocumentApp.ElementType.INLINE_IMAGE;

    if(body.findText("{画像}")){  
      var imagePlaceholder = body.findText("{画像}").getElement();
      Logger.log("imagePlaceholder: " + imagePlaceholder.getText())
      var imageIndex =  body.getChildIndex(imagePlaceholder.getParent());
      Logger.log("imageIndex:" + imageIndex);
      if(imageId){
        var image = DriveApp.getFileById(imageId);
        var inlineImage = body.insertImage(imageIndex, image);
        resizeImage(inlineImage);
      }
      imagePlaceholder.removeFromParent();
    }
    
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
      .replaceText("{shortUrl}",shortUrl)
    }
    
    var replacers = getReplacingWords();
    _._each(replacers, function(rep, key){
      newBody.replaceText('{' + rep +'}', getRange(mySheet,i, key).getValue())
    }, null)

    // リンクを編集
    replaceLinks(newBody, shortUrl);
    
    Logger.log("書き込みました：" + newBody); //ドキュメントの内容をログに表示
 
  }
}



