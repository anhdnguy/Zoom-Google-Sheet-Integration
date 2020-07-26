# Zoom-Google-Sheet-Integration

This Integraton can be used for organizations that use Zoom as their primary video communications platform.

This peice of code helps the Admininstrators choose the availabe sub Zoom account under their organization to schedule a meeting. All of that can be done in 1 single Google Spreadsheet.

1. Setting up a Zoom OAuth with the appropriate scopes.

2. Getting the sub Zoom accounts along with their Ids that will be used for scheduling meetings. In this case, it will be all Zoom accounts that have the Licensed account type.

3. This code will find the Zoom account that is available for the inserted start date, start time, and end time. Once it found the account and the user has verified it, it will create the Zoom meeting on that chosen account.
