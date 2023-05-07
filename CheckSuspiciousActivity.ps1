# Jordan NCSC
# Digital Forensics Team
# May 2023
###############################################################################################################
#

# Enable Admin Audit Log if it's not already enabled
Set-AdminAuditLogConfig -AdminAuditLogEnabled $true

# Prompt the user for the suspected email address
$suspectedEmailAddress = Read-Host "Enter the suspected email address"

# Search for Set-Mailbox audit logs related to the suspected mailbox
$auditLogs = Search-AdminAuditLog -Cmdlets Set-Mailbox -Parameters Identity -ObjectIds $suspectedEmailAddress

# Prompt the user for the output CSV file name
$csvFileName = Read-Host "Enter the output CSV file name (without extension)"
$csvFilePath = $csvFileName + ".csv"

# Export audit logs to a CSV file
$auditLogs | Export-Csv -Path $csvFilePath -NoTypeInformation

Write-Host "Audit logs exported to $csvFilePath"
