CheckSuspiciousActivity.ps1 - Email Forwarding Audit Script
===========================================================

Description
-----------
This script checks for suspicious activities related to email forwarding changes in Microsoft Exchange and Active Directory. Developed by the Jordan NCSC DF team, it helps identify unauthorized changes and attempts to compromise email accounts.

Prerequisites
-------------
- PowerShell version 5.1 or later
- Active Directory PowerShell module installed
- Microsoft Exchange Management Shell installed
- Administrative access to the Exchange server and Active Directory

Usage
-----
1. Open PowerShell with administrative privileges.
2. Navigate to the folder where "CheckSuspiciousActivity.ps1" is located.
3. Run the script by typing `.\CheckSuspiciousActivity.ps1` and pressing Enter.

The script will guide you through the following steps:

1. Enable the Admin Audit Log (if not already enabled).
2. Enter the suspected email address.
3. Choose a date range option (last month, last 3 months, last 6 months, or last year).
4. Enter the username.
5. Decide whether to save the audit logs to a CSV file or not.

The script will display the audit logs on the screen at the end of the process.

Support
-------
If you encounter any issues or require further assistance, please contact the Jordan NCSC DF team.

-------

Here are the instructions for installing and setting up the required components to run the "CheckSuspiciousActivity.ps1" script:

Install PowerShell 5.1 or later:

PowerShell 5.1 is included in Windows 10 and Windows Server 2016 by default.
If you're using an earlier version of Windows, you can download and install the Windows Management Framework 5.1, which includes PowerShell 5.1: https://www.microsoft.com/en-us/download/details.aspx?id=54616
Install Active Directory PowerShell module:

On Windows Server, you can install the Active Directory module as a part of the Remote Server Administration Tools (RSAT) feature. Run the following command in PowerShell with administrator privileges:
mathematica
Copy code
PS> Install-WindowsFeature RSAT-AD-PowerShell
On Windows 10, you can download and install RSAT for your version of Windows 10 from the following link: https://www.microsoft.com/en-us/download/details.aspx?id=45520
Install Microsoft Exchange Management Shell:

For Exchange Server 2016 or 2019, follow the instructions in this Microsoft article: https://docs.microsoft.com/en-us/exchange/connectors/powershell/install-powershell
For Exchange Online (Office 365), you can install the Exchange Online Management module by running the following command in PowerShell with administrator privileges:
mathematica
Copy code
PS> Install-Module -Name ExchangeOnlineManagement
Note that this requires PowerShellGet, which is included in PowerShell 5.1.
Administrative access to the Exchange server and Active Directory:

Make sure the user running the script has administrative privileges on the Exchange server and Active Directory. This typically includes members of the "Domain Admins" and "Organization Management" security groups.
