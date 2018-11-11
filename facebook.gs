function postMessage(psid) {
   var data =
   {
      "recipient":{
        "id":psid
      },
      "message":{
        "text":"hello, world!"
      }
   };

   var options =
   {
     "method" : "post",
     "contentType": "application/json",
     "payload" : JSON.stringify(data)

   };

   var token = PropertiesService.getScriptProperties().getProperty("PAGE_ACCESS_TOKEN");
   UrlFetchApp.fetch("https://graph.facebook.com/v2.6/me/messages?access_token=" + token, options);
}
