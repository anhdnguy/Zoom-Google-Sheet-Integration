# Zoom and Google Spreadsheet Integration

This Integraton can be used for organizations that use Zoom as their primary video communications platform. It helps the Admininstrators choose the availabe sub Zoom account under their organizations to schedule a meeting. All of that can be done in one single Google Spreadsheet.

The Integration is built using API method. Google Spreadsheet will send an API GET request to Zoom to receive the list of meetings from all accounts that are used for scheduling. Based on the the date and time entered by users, it will recommend the users which Zoom account is available. If the users choose to schedule the meeting on the recommended account, Google Spreadsheet will send an API POST request to Zoom to create the meeting. If Zoom receives the request successfully, Zoom will respond with the meeting URL.

1. Create a blank Google Spreadsheet with the following columns.<br><br>
![SpreadSheet Image](pictures/Google_Sheet.png)

2. Using the Zoom account with the Admin previlege, set up a Zoom OAuth App with the appropriate scopes under Zoom Marketplace. The scopes can be group:write:admin/meeting:master/meeting:write:admin/user:master/user:write:admin. Insert the Spreadsheet's URL into the Whitelist URL.

3. Get the sub Zoom accounts along with their Ids that will be used for scheduling meetings. Add the accounts information on a different tab. For the purpose of my project, I have 2 separated tabs for accounts reference. The script will use Sheet 1 to find the available accounts and Sheet 2 to find the account to schedule the meeting on:
  - Sheet 1 contains the list of accounts which are strictly used for customers and services meeting.
  - Sheet 2 contains accounts that are listed in Sheet 1 and a few more special accounts.

4. Open Script Editor on the Spreadsheet and start coding.

Below is an example of how the integration works
![example_gif](pictures/example2.gif)
