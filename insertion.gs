/* insertion.gs
文字列をテンプレートに挿入する機能を提供します。
templateEngineにマージするかも
*/
function ReplacingText(symbol, replacingFunction){
    this.symbol= symbol;
    this.replacer=replacingFunction;
  
  function replace(text){
    text.replace(symbol, replacer());
  }
}


function getReplacingWords(){
  var replacers = {};
  var _ = underscoreGS;
  var sheet = SpreadsheetApp.getActiveSheet();
  var cols = ["H", "I", "J", "K", "L"];
  //{ H:自由入力1 とかになって欲しい
  _._each(cols, function(col){ 
    replacers[col] = getRange(sheet,1, col).getValue(); 
  }, null);

  Logger.log(replacers);
  return replacers;
}