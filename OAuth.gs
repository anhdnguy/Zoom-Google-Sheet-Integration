var Client_ID = 'Client_ID';
var Client_Secret = 'Client_Secret';
var auth = Utilities.base64Encode(Client_ID + ':' + Client_Secret);

function RefreshToken() {  
  var keys_api = PropertiesService.getDocumentProperties();
  var refresh_token = keys_api.getProperty('refresh_token');
  var url_refresh = 'https://zoom.us/oauth/token?grant_type=refresh_token&refresh_token=' + refresh_token;
  var Headers_refresh = {
    'Authorization': 'Basic ' + auth
  };
  
  var request_refresh = {
    'headers' : Headers_refresh,
    'method' : 'POST'
  };
  var response = UrlFetchApp.fetch(url_refresh,request_refresh);
  
  var content = response.getContentText();
  
  var json = JSON.parse(content);
  keys_api.setProperties ({
    'access_token' : json['access_token'],
    'refresh_token' : json['refresh_token']
  });
  Logger.log(keys_api.getProperty('access_token'));
}
