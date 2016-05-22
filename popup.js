var token;

function authorize() {
  var clientId = "<enter your client id>";
  var redirectUrl = "https%3A%2F%2F<enter your chrome extension id>.chromiumapp.org%2Foauth2";    
  var authUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=" + clientId + "&response_type=id_token+token&scope=openid%20https%3A%2F%2Foutlook.office.com%2Fcontacts.read&response_mode=fragment&state=167381&nonce=678910&redirect_uri=" + redirectUrl;
  chrome.identity.launchWebAuthFlow(
    { 
      interactive:true,
      url:authUrl
    },
    function(responseUrl) {
      if (responseUrl) 
        token = parseToken(responseUrl);

      chrome.tabs.query({active: true, currentWindow: true}, function(tabs) {
        var selectedTabId = tabs[0].id;
        chrome.tabs.executeScript(selectedTabId, {file: "lib/jquery-2.2.3.min.js"});
        chrome.tabs.executeScript(selectedTabId, {file: "action.js"}, function(results) {
          chrome.tabs.sendMessage(selectedTabId, token);
        });
      });
    });
};

function parseToken(responseUrl) {
    var parts = responseUrl.split("#");
    if (parts.length < 2)
      return;

    var queryParts = parts[1].split("&");
    if (queryParts.length == 0)
      return 

    var tokenQueryVar = queryParts[0];
    var token = tokenQueryVar.replace("access_token=", "");

    return token;
}

chrome.browserAction.onClicked.addListener(function(tab) { 
  authorize();
});

