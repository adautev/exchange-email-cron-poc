# Office 365 email parser
Retrieves emails under a service account by filtering them.
Integrated with Azure app insights.

## Configuration
File users.txt must be in the root program folder, containing comma separated UIDs for users for which emails should be read. 

### Microsoft Graph authentication
Create an Azure app and give Mail.Read API access. Get App secret and configure it into appsettings.json.

:)
