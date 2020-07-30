// Create a menu item and the button to trigger the action
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Zoom Integration');
  menu.addItem('Create Zoom Meetings', 'crtZoom').addToUi();
  menu.addItem('Recommend Zoom Account', 'ZoomAvailability').addToUi();
  menu.addItem('Clear Entry', 'clearEntry').addToUi();
}

// Clearing entry when the user want to work on a new batch
function clearEntry () {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var numRows = sheet.getLastRow();
  var numColumns = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, numRows, numColumns).clearContent();
}

// Getting Zoom Meetings order by Account and Start time ASC
function getMeetings () {
  var lstMeetings = [];
  var startdate;
  var starttime;
  var endtime;
  var url;
  // Getting API key to get meetings
  var response;
  var token = PropertiesService.getDocumentProperties().getProperty('access_token');
  var Headers = {
    'Authorization' : 'Bearer ' + token
  };
  var request = {
    'headers' : Headers,
    'method' : 'GET'
  };
  
  //Getting Zoom Accounts Ids
  var maxRow_sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1].getLastRow() - 1;
  var lstAcc = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1].getRange(2, 1, maxRow_sheet2, 3).getValues();
  
  // Get Zoom Account
  for (var i = 0; i < lstAcc.length; i++) {
    lstMeetings[i] = [];
    url = 'https://api.zoom.us/v2/users/' + lstAcc[i][2] + '/meetings?type=upcoming&page_size=300';
    response = JSON.parse(UrlFetchApp.fetch(url, request).getContentText());
    
    //Get meetings DateTime ASC
    for (var j = 0; j < response.total_records; j++) {
      
      starttime = new Date(response.meetings[j].start_time);
      endtime = starttime.getTime() + (response.meetings[j].duration * 60000);
      lstMeetings[i][j] = {
        'start_time' : starttime.getTime(),
        'end_time' : endtime
      };
    }
  }
  return lstMeetings;
}

// Check and return Recommended Zoom Accounts
// The order of accounts that will be taken to consider is based on how we sort them on the accounts tabs
// If we are creating mutiple simultaneous meetings, the script won't recommend the same account for them.
// Instead, if an account is reserved for a meeting, the script will add the meeting time to the timeline. Therefore,
// there won't be confliction.

function ZoomAvailability() {
  // Getting Meeting Info
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheets()[0];
  var numRows = sheet.getLastRow()-1;
  var numColumns = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, numRows, numColumns).getValues();
  var starttime, endtime;
  var username = [];
  var enddate, createddate;
  
  // Getting Zoom Accounts
  var maxRow_sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1].getLastRow() - 1;
  var lstAcc = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1].getRange(2, 1, maxRow_sheet2, 3).getValues();
  
  var lstScheduledMeetings = getMeetings();
  var temp = [];
  
  for (var i = 0; i < lstScheduledMeetings.length; i++) {
    temp[i] = [];
    for (var j = 0; j < lstScheduledMeetings[i].length; j++) {
      temp[i][j] = lstScheduledMeetings[i][j];
    }
  }
  
  createddate = new Date();
  createddate = {
    'start_time' : createddate.getTime(),
    'end_time' : createddate.getTime()
  };
  
  for (var i = 0; i < numRows; i++) {
    starttime = new Date(data[i][2].setHours(data[i][3].getHours()));
    starttime = new Date(data[i][2].setMinutes(data[i][3].getMinutes()-30));
    starttime = starttime.getTime();
    var temp_starttime = starttime;
    endtime = new Date(data[i][2].setHours(data[i][4].getHours()));
    endtime = new Date(data[i][2].setMinutes(data[i][4].getMinutes()));
    endtime = endtime.getTime();
    var temp_endtime = endtime;
    if (data[i][10] == 0) {
      data[i][10] = new Date(data[i][2]);
    }
    enddate = new Date(data[i][10]);
    enddate = {
      'start_time' : enddate.getTime(),
      'end_time' : enddate.getTime()
    };
    
    for (var x = 0; x < temp.length; x++) {
      temp[x].splice(0, 0, createddate);
      temp[x].splice(temp[x].length, 0, enddate);
      while (temp_starttime <= enddate.end_time) {
        username[i] = 0;
        for (var y = 1; y < temp[x].length; y++) {
          if (temp_starttime > temp[x][y-1].end_time && temp_endtime < temp[x][y].start_time) {
            username[i] = [lstAcc[x][0]];
            temp[x].splice(y, 0, {
                           'start_time' : temp_starttime,
                           'end_time' : temp_endtime
                           });
            break;
          }
        }
        if (username[i] == 0) {
          temp[x] = [];
          var z = 0;
          while (z < lstScheduledMeetings[x].length) {
            temp[x][z] = lstScheduledMeetings[x][z];
            z++;
          }
          temp_starttime = starttime;
          temp_endtime = endtime;
          break;
        }
        else {
          lstScheduledMeetings[x].splice(y-1, 0, {
                           'start_time' : temp_starttime,
                           'end_time' : temp_endtime
                           });
          temp_starttime = new Date(temp_starttime);
          temp_starttime = temp_starttime.setDate(temp_starttime.getDate()+7);
          temp_endtime = new Date(temp_endtime);
          temp_endtime = temp_endtime.setDate(temp_endtime.getDate()+7);
        }
      }
      if (username[i] != 0) {
        temp[x].splice(0, 1);
        temp[x].splice(temp[x].length-1, 1);
        break;
      }
    }
    if (username[i] == 0) {
      username[i] = ['Cant find any Zoom Account'];
    }
  }
  
  sheet.getRange(2, 12, numRows, 1).setValues(username);
}

// Select Zoom Account
function whichAccount(username) {
  var getId = [];
  var acc_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var acc_sheet = acc_spreadsheet.getSheets()[2];
  var numRows = acc_sheet.getLastRow() - 1; // Number of rows
  var numColumns = acc_sheet.getLastColumn(); // Number of columns
  var data = acc_sheet.getRange(2, 1, numRows, numColumns).getValues();
  for (var i = 0; i < username.length; i++) {
    for (var j = 0; j < numRows; j++) {
      if (username[i] == data[j][0]) {
      getId[i] = data[j][2];
      }
    }
  }
  return getId;
}

// Creating New Meetings
// 
function newMeeting(data, lstId) {
  var response;
  var token = PropertiesService.getDocumentProperties().getProperty('access_token');
  var url;
  var request = [];
  var dummy;
  var type = [];
  var start_time = [];
  var recur_day = [];
  var end_date_time = [];
  var Zoomlinks = [];
  var body;
  var recur_body;
  
  var Headers = {
    'Authorization' : 'Bearer ' + token,
    'Content-Type' : 'application/json'
  };
  
  var Settings = {
    'meeting_authentication' : true
  };
  
  for (var i = 0; i < data.length; i++) {
    start_time[i] = new Date (Utilities.formatDate(data[i][2], "America/Los_Angeles", "yyyy-MM-dd") + ' ' + Utilities.formatDate(data[i][3], "America/Los_Angeles", "HH:mm:ss"));
    start_time[i] = Utilities.formatDate(start_time[i], "America/Los_Angeles", "yyyy-MM-dd'T'HH:mm:ss");
    end_date_time[i] = Utilities.formatDate(new Date(data[i][10]), "America/Los_Angeles", "yyyy-MM-dd'T'HH:mm:ss'Z'");
    dummy = new Date(data[i][2]);
    recur_day[i] = dummy.getDay() + 1;
    if (data[i][7] == 'Yes') {
      type[i] = 8;
      if (data[i][8] == 'Weekdays') {
        recur_body = {
          'type' : '2',
          'weekly_days' : recur_day[i],
          'end_date_time' : end_date_time[i]
        };
      }
      else {
        recur_body = {
          'type' : '3',
          'monthly_week' : data[i][9],
          'monthly_week_day' : recur_day[i],
          'end_date_time' : end_date_time[i]
        };
      }
    }
    else {
      type[i] = 2;
    }
    
    body = {
      'topic' : data[i][0],
      'type' : type[i],
      'start_time' : start_time[i],
      'duration' : data[i][5],
      'password' : data[i][6],
      'timezone' : 'America/Los_Angeles',
      'recurrence' : recur_body,
      'settings' : Settings
    };
    
    request = {
      'headers' : Headers,
      'method' : 'POST',
      'payload' : JSON.stringify(body),
    };
        
    url = 'https://api.zoom.us/v2/users/' + lstId[i] + '/meetings';
    
    response = JSON.parse(UrlFetchApp.fetch(url,request).getContentText());
    Zoomlinks[i] = [response['join_url']];
  }
  
  return Zoomlinks;
  
}

// Control function
function crtZoom() {  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheets()[0];
  var numRows = sheet.getLastRow()-1;
  var numColumns = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, numRows, numColumns).getValues();
  var lstId = whichAccount(sheet.getRange(2, 12, numRows, 1).getValues());
  var Zoomlink = [];
  var topic_day, topic_acc, password_name;
  
  // Update the Topic of the meeting and Randomly create the meeting password with 10 characters.
  for (var i = 0; i < data.length; i++) {
    topic_acc = data[i][11].toUpperCase();
    topic_acc = topic_acc.split('@')[0];
    topic_day = Utilities.formatDate(new Date(data[i][2]), "America/Los_Angeles", "EEEE");
    data[i][0] = data[i][1] + ' - ' + topic_acc + ' - ' + topic_day + ' - ' +
      Utilities.formatDate(new Date(data[i][3]), "America/Los_Angeles", "h:mm a") + ' - ' + Utilities.formatDate(new Date(data[i][4]), "America/Los_Angeles", "h:mm a") + ' PST';
    sheet.getRange(i+2, 1).setValue(data[i][0]);
    
    data[i][6] = Math.random().toString(36).slice(-10);
    sheet.getRange(i+2, 7).setValue(data[i][6]);
  }
  
  Zoomlink = newMeeting(data,lstId);
  
  for (var i = 0; i < Zoomlink.length; i++) {
    Zoomlink[i][0] = Zoomlink[i][0].split('?')[0];
  }
  
  sheet.getRange(2, 13, numRows, 1).setValues(Zoomlink);
  
}
