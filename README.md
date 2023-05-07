Check Suspicious Activity PowerShell Script (CheckSuspiciousActivity.ps1)
--------------------------------------------------------------------------

Description:
This script helps you investigate suspicious email forwarding activity in your Microsoft Exchange environment. It searches for Set-Mailbox audit logs in Exchange, changes to the msExchGenericForwardingAddress property in Active Directory, and then displays the results on the screen. Optionally, you can save the results to a CSV file.

Requirements:
- PowerShell version 5.1 or later
- Active Directory PowerShell module installed
- Microsoft Exchange Management Shell installed
- Administrative access to the Exchange server and Active Directory

How to install requirements:
1. PowerShell 5.1 or later:
   PowerShell 5.1 is included with Windows Server 2016 and Windows 10. For earlier versions of Windows, you can install Windows Management Framework (WMF) 5.1 to update PowerShell. Download WMF 5.1 from here: https://www.microsoft.com/en-us/download/details.aspx?id=54616

2. Active Directory PowerShell module:
   The Active Directory PowerShell module is included with the Remote Server Administration Tools (RSAT) package. To install RSAT, follow these instructions for your specific operating system:
   - Windows Server 2016 or later: RSAT is included as an optional feature. Run this command in PowerShell as an administrator: `Install-WindowsFeature RSAT-AD-PowerShell`
   - Windows 10: Download RSAT from here: https://www.microsoft.com/en-us/download/details.aspx?id=45520 and follow the installation instructions.

3. Microsoft Exchange Management Shell:
   To install the Exchange Management Shell, you must install the Exchange server on the machine where you plan to run the script, or install the Exchange admin center tools. For more information, refer to the following Microsoft documentation: https://docs.microsoft.com/en-us/exchange/exchange-management-shell

Usage:
1. Open PowerShell as an administrator.
2. Navigate to the folder containing the CheckSuspiciousActivity.ps1 script.
3. Run the script by entering ".\CheckSuspiciousActivity.ps1" (without quotes) and pressing Enter.

During the execution of the script, you will be prompted to provide the following information:
- Suspected email address
- Date range for the search (last month, last 3 months, last 6 months, or last year)
- Username associated with the suspected email address
- Domain Controller's name
- Whether to save the results to a CSV file or not

The script will display the search results on the screen, and if you choose to save the results to a CSV file, you will be prompted to provide the file path.

Please note that the script checks for the required PowerShell version and modules and enables the Admin Audit Log. If there are any errors during the script execution, it will display error messages and continue with the remaining steps.
