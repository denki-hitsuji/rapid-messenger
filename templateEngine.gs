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


