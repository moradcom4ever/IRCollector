# IRCollector
Email Forwarding Audit Script - This script will check for suspicious activities related to email forwarding changes in Exchange.

To execute the batch script, simply double-click the "CheckSuspiciousActivity.bat" file or run it from the Command Prompt. 
The script will enable the admin audit log (if not already enabled), search for Set-Mailbox cmdlet logs related to the suspected email address, and export the logs to a CSV file named "AuditLogs.csv" in the same directory.

You can open the "AuditLogs.csv" file using Microsoft Excel or another spreadsheet application to review the logs and determine if there were any suspicious activities on the suspected machine.
