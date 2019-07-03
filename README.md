# Office365LicenseReport
Powershell tool that connects and gets Office 365 License Report.

Powershell 5.1 is required and .net framework 4.7
You must have the tenantlist.csv in the same directory as the powershell script and fill the file with a list of your tenant IDs and a corresponding name.
The script will prompt for a user name and password and you must supply one that has permissions to access the tennant. 2fa cannot be enabled on this user.
