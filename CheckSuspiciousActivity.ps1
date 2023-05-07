# Import ActiveDirectory module
Import-Module ActiveDirectory

# Welcome message
Write-Host "Jordan NCSC, DF team - Email Forwarding Audit Script" -ForegroundColor Cyan
Write-Host "This script will check for suspicious activities related to email forwarding changes in Exchange and Active Directory." -ForegroundColor Cyan
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

# Define the date range options
$today = (Get-Date).Date
$dateOptions = @{
    "1" = @{
        "label" = "last month"
        "start" = $today.AddMonths(-1)
    }
    "2" = @{
        "label" = "last 3 months"
        "start" = $today.AddMonths(-3)
    }
    "3" = @{
        "label" = "last 6 months"
        "start" = $today.AddMonths(-6)
    }
    "4" = @{
        "label" = "last year"
        "start" = $today.AddYears(-1)
    }
}

# Prompt the user to choose a date range option
Write-Host "Step 3: Choose a date range option:" -ForegroundColor Green
$dateOptions.Keys | ForEach-Object { Write-Host "[$_]: $($dateOptions[$_].label)" }
$selectedOption = Read-Host "Enter the number of the desired option"
$selectedStartDate = $dateOptions[$selectedOption].start

# Search for Set-Mailbox audit logs related to the suspected mailbox
Write-Host "Step 4: Searching for Set-Mailbox audit logs related to the suspected mailbox in Exchange..." -ForegroundColor Green
$exchangeAuditLogs = Search-AdminAuditLog -Cmdlets Set-Mailbox -Parameters Identity -ObjectIds $suspectedEmailAddress -StartDate $selectedStartDate
Write-Host "Exchange audit logs retrieved." -ForegroundColor Yellow
Write-Host ""

# Prompt the user for the username
Write-Host "Step 5: Enter the username..." -ForegroundColor Green
$username = Read-Host "Username"

# Get the user's distinguished name
$UserDistinguishedName = (Get-ADUser -Identity $username).DistinguishedName

# Query the AD replication attribute metadata for changes to the msExchGenericForwardingAddress property
Write-Host "Step 6: Searching for changes to the msExchGenericForwardingAddress property in Active Directory..." -ForegroundColor Green
$Changes = Get-ADReplicationAttributeMetadata -Object $UserDistinguishedName -Properties msExchGenericForwardingAddress -ChangedAfter $selectedStartDate
Write-Host "Active Directory audit logs retrieved." -ForegroundColor Yellow
Write-Host ""

# Loop through the changes and look for corresponding Security events with event ID 5136
$adAuditLogs = $Changes | ForEach-Object {
    $Metadata = $_
    $Event = Get-WinEvent -FilterHashtable @{
        LogName   = 'Security'
        Id        = 5136
        StartTime = $Metadata.LastOriginatingChangeTime
    } | Where-Object { $_.Properties[2].Value -eq $Metadata.AttributeName -and $_.Properties[1].Value -eq $UserDistinguishedName }
	if ($Event) {
    $Event | Select-Object TimeCreated,
        @{Name='ChangedBy'; Expression={($_.Properties[0].Value)}},
        @{Name='AttributeName'; Expression={($_.Properties[2].Value)}},
        @{Name='OldValue'; Expression={($_.Properties[4].Value)}},
        @{Name='NewValue'; Expression={($_.Properties[5].Value)}}
	}
}
Write-Host ""

#Combine Exchange and Active Directory audit logs
$auditLogs = $exchangeAuditLogs + $adAuditLogs

#Ask the user if they want to save the output to a CSV file
Write-Host "Step 7: Do you want to save the audit logs to a CSV file? (yes/no)" -ForegroundColor Green
$saveToCsv = Read-Host "Answer"
$saveToCsv = $saveToCsv.ToLower() -eq "yes"

if ($saveToCsv) {
	# Prompt the user for the output CSV file name
	Write-Host "Enter the output CSV file name (without extension)..." -ForegroundColor Green
	$csvFileName = Read-Host "CSV file name"
	$csvFilePath = $csvFileName + ".csv"
	Write-Host ""
	
	# Export audit logs to a CSV file
	Write-Host "Step 8: Exporting audit logs to a CSV file..." -ForegroundColor Green
	$auditLogs | Export-Csv -Path $csvFilePath -NoTypeInformation
	Write-Host "Audit logs exported to $csvFilePath" -ForegroundColor Yellow
	Write-Host ""
}

#Display the audit logs on the screen
Write-Host "Step 9: Displaying audit logs on the screen..." -ForegroundColor Green
$auditLogs | Format-Table -AutoSize
Write-Host "Operation completed." -ForegroundColor Yellow
