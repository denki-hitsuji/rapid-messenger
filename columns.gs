var NAME_COL = 3;
var newUrlCol;
function getNewUrlCol(){
  if(newUrlCol){return newUrlCol;}
  
  newUrlCol = PropertiesService.getDocumentProperties().getProperty("newUrlCol");
  if(newUrlCol == null || newUrlCol == "undefined")
  { newUrlCol = "AA"; }
  return newUrlCol;
}
var newIdColCache;
function getNewIdCol(){
  if(newIdColCache){return newIdColCache;}
  
  newIdColCache = PropertiesService.getDocumentProperties().getProperty("newIdCol");
  if(newIdColCache == null || newIdColCache == "undefined")
  { newIdColCache = "AH"; }
  
  Logger.log("newIdColCache:" + newIdColCache + " " + typeof newIdColCache);
  return newIdColCache;
}

var originUrlColCache;
function getOriginUrlCol(){
  if(originUrlColCache){return originUrlColCache;}
  
  originUrlColCache = PropertiesService.getDocumentProperties().getProperty("originUrlCol");
  if(originUrlColCache == null || originUrlColCache === "undefined")
  { originUrlColCache = "Z"; }

  Logger.log("originUrlCol:" + originUrlColCache+ " " + typeof originUrlColCache);
  return originUrlColCache;
}
var nicknameCol;
function getNicknameCol(){
  if(nicknameCol){return nicknameCol;}
  
  nicknameCol = PropertiesService.getDocumentProperties().getProperty("nicknameCol");
  if(nicknameCol == null || nicknameCol === "undefined")
  { nicknameCol = 4; }
  return nicknameCol;
}
var mailAddressCol
function getMailAddressCol(){
  if(mailAddressCol){return mailAddressCol;}
  
  mailAddressCol = PropertiesService.getDocumentProperties().getProperty("mailAddressCol");
  if(mailAddressCol == null || mailAddressCol === "undefined")
  { mailAddressCol = 5; }
  return mailAddressCol;
}
