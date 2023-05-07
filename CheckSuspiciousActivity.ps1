# Welcome message
Write-Host "Jordan NCSC, DF team - Email Forwarding Audit Script" -ForegroundColor Cyan
Write-Host "This script will check for suspicious activities related to email forwarding changes in Exchange." -ForegroundColor Cyan
Write-Host ""

# Enable Admin Audit Log if it's not already enabled
Write-Host "Step 1: Enabling the Admin Audit Log (if not already enabled)..." -ForegroundColor Green
Set-AdminAuditLogConfig -AdminAuditLogEnabled $true
Write-Host "Admin Audit Log enabled." -ForegroundColor Yellow
Write-Host ""

# Prompt the user for the suspected email address
Write-Host "Step 2: Enter the suspected email address..." -ForegroundColor Green
$suspectedEmailAddress = Read-Host "Email address"
Write-Host ""

# Search for Set-Mailbox audit logs related to the suspected mailbox
Write-Host "Step 3: Searching for Set-Mailbox audit logs related to the suspected mailbox..." -ForegroundColor Green
$auditLogs = Search-AdminAuditLog -Cmdlets Set-Mailbox -Parameters Identity -ObjectIds $suspectedEmailAddress
Write-Host "Audit logs retrieved." -ForegroundColor Yellow
Write-Host ""

# Prompt the user for the output CSV file name
Write-Host "Step 4: Enter the output CSV file name (without extension)..." -ForegroundColor Green
$csvFileName = Read-Host "CSV file name"
$csvFilePath = $csvFileName + ".csv"
Write-Host ""

# Export audit logs to a CSV file
Write-Host "Step 5: Exporting audit logs to a CSV file..." -ForegroundColor Green
$auditLogs | Export-Csv -Path $csvFilePath -NoTypeInformation
Write-Host "Audit logs exported to $csvFilePath" -ForegroundColor Yellow
Write-Host ""

# Display the audit logs on the screen
Write-Host "Step 6: Displaying audit logs on the screen..." -ForegroundColor Green
$auditLogs | Format-Table -AutoSize
Write-Host "Operation completed." -ForegroundColor Yellow
