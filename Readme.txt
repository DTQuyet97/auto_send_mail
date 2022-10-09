[Module import]:
- python 3.4 or above (recommended 3.6)
- O365
- openpyxl

[Pre]: need to register your application at Azure App Registrations (Just once)
* Step to register an application for this tool to send email:
1. Login at Azure Portal: https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade
2. Create an app. Set a name
3. In Supported account types choose "Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook.com)", if you are using a personal account.
4. Set the redirect uri (Web) to: https://login.microsoftonline.com/common/oauth2/nativeclient and click register
5. Write down the Application (client) ID. You will need this value
6. Under "Certificates & secrets", generate a new client secret. Set the expiration preferably to never. Write down the value of the client secret created now. It will be hidden later on
7. Under Api Permissions:
- When authenticating "on behalf of a user":
	1. add the delegated permissions for Microsoft Graph you want (see scopes)
	2. It is highly recommended to add "offline_access" permission

* Open and modify "input.txt": given some values (App_Id and Client_Secret just need modify once when sender doesn't modify)
- App_Id: ID in step 5
- Client_Secret: client secret in step 6
- CC: Whom you want to CC when send email
- Path_Device: absolute path which contain information to send email
- Project: Current project which will in Subject and Sign of email.

[Execute]:
- Run file Assets_Confirm_Email.py
- Do follow command in Console output: visit a url link to get tokens
- Copy and paste url token back to Console and enter to continue
- The process will continue sending email

Path: D:\psx_trunk\trunk\Plans\CM\Assets\LoanList\Customer_Device_info.xlsx