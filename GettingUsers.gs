// Getting the list of users that have Licensed account type.

function gettingUsers() {
  var spreadsheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users Info');
  var response;
  var url;
  var lstUsers = [];
  var response;
  var count;
  var token = PropertiesService.getDocumentProperties().getProperty('access_token');
  var Headers = {
    'Authorization' : 'Bearer ' + token
  };
  
  var request = {
    'headers' : Headers,
    'method' : 'GET'
  };

  var j = 0;
  var group_id_range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data_Reference').getRange(4, 1, 7, 2).getValues();

  for (var a = 0; a < 8; a++) {
    count = a+1;
    url = 'https://api.zoom.us/v2/users?status=active&page_size=300&page_number=' + count;
    response = JSON.parse(UrlFetchApp.fetch(url,request).getContentText());
    for (var i = 0; i < response.page_size; i++) {
      if (response.users[i] != null && response.users[i].type == 2) {
        lstUsers[j] = [];
        lstUsers[j][0] = [response.users[i].email];
        lstUsers[j][1] = [response.users[i].type];
        lstUsers[j][2] = [response.users[i].id];
        lstUsers[j][3] = [response.users[i].dept];
        lstUsers[j][4] = [response.users[i].group_ids];
        for (var n = 0; n < group_id_range.length; n++) {
          if (lstUsers[j][4] == group_id_range[n][0]) {
            lstUsers[j][4] = group_id_range[n][1];
          }
        }
        j++;
      }
    }
  }
  spreadsheet2.getRange(2, 1, lstUsers.length, 5).setValues(lstUsers);
}
