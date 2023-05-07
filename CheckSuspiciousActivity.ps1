# Check if PowerShell version is 5.1 or later
Write-Host "Checking PowerShell version..." -ForegroundColor Green
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Host "PowerShell 5.1 or later is required. Please update PowerShell and try again." -ForegroundColor Red
}

# Check if Active Directory PowerShell module is installed
Write-Host "Checking Active Directory PowerShell module..." -ForegroundColor Green
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "Active Directory PowerShell module is not installed. Please install it and try again." -ForegroundColor Red
}

# Import necessary modules
Write-Host "Importing required modules..." -ForegroundColor Green
try {
    Import-Module ActiveDirectory
    . $env:ExchangeInstallPath\bin\RemoteExchange.ps1
    Connect-ExchangeServer -auto
} catch {
    Write-Host "Error loading required modules: $_" -ForegroundColor Red
}

# Enable Admin Audit Log
Write-Host "Enabling Admin Audit Log..." -ForegroundColor Green
try {
    Set-AdminAuditLogConfig -AdminAuditLogEnabled $true
} catch {
    Write-Host "Error enabling Admin Audit Log: $_" -ForegroundColor Red
}

# Prompt the user for the suspected email address
Write-Host "Step 1: Enter the suspected email address" -ForegroundColor Green
$suspectedEmail = Read-Host "Suspected Email Address"
Write-Host ""

# Prompt the user to choose a date range option
Write-Host "Step 2: Choose a date range option:" -ForegroundColor Green
Write-Host "[1]: last month" -ForegroundColor Yellow
Write-Host "[2]: last 3 months" -ForegroundColor Yellow
Write-Host "[3]: last 6 months" -ForegroundColor Yellow
Write-Host "[4]: last year" -ForegroundColor Yellow

$dateRangeChoice = Read-Host "Enter the number for the desired date range"
Write-Host ""

switch ($dateRangeChoice) {
    "1" { $startDate = (Get-Date).AddMonths(-1) }
    "2" { $startDate = (Get-Date).AddMonths(-3) }
    "3" { $startDate = (Get-Date).AddMonths(-6) }
    "4" { $startDate = (Get-Date).AddYears(-1) }
    default { Write-Host "Invalid choice. Exiting."; exit }
}

# Search for Set-Mailbox audit logs in Exchange
Write-Host "Searching for Set-Mailbox audit logs in Exchange..." -ForegroundColor Green
try {
    $searchResults = Search-AdminAuditLog -Cmdlets Set-Mailbox -StartDate $startDate -ObjectIds $suspectedEmail
} catch {
    Write-Host "Error searching for Set-Mailbox audit logs: $_" -ForegroundColor Red
}

# Prompt the user for the username
Write-Host "Step 3: Enter the username" -ForegroundColor Green
$username = Read-Host "Username"
Write-Host ""

# Prompt the user for the Domain Controller
Write-Host "Step 4: Enter the Domain Controller's name" -ForegroundColor Green
$domainController = Read-Host "Domain Controller"
Write-Host ""

# Search for changes to the msExchGenericForwardingAddress property in Active Directory
Write-Host "Searching for changes to the msExchGenericForwardingAddress property in Active Directory..." -ForegroundColor Green
try {
    $UserDistinguishedName = (Get-ADUser -Identity $username).DistinguishedName
    $Changes = Get-ADReplicationAttributeMetadata -Object $UserDistinguishedName -Properties msExchGenericForwardingAddress -Server $domainController
} catch {
    Write-Host "Error searching for changes to msExchGenericForwardingAddress property: $_" -ForegroundColor Red
}

# Export changes to a CSV file
Write-Host "Exporting changes to a temporary CSV file..." -ForegroundColor Green
try {
    $Changes | ForEach-Object {
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
    } | Export-Csv -Path "SuspiciousActivity.csv" -NoTypeInformation
} catch {
    Write-Host "Error exporting changes to CSV: $_" -ForegroundColor Red
}

# Print the results on the screen
Write-Host "Search Results:" -ForegroundColor Green
Import-Csv -Path "SuspiciousActivity.csv" | Format-Table -AutoSize

# Prompt the user to choose whether to save the results to a CSV file or not
Write-Host "Step 5: Save the results to a CSV file?" -ForegroundColor Green
Write-Host "[1]: Yes" -ForegroundColor Yellow
Write-Host "[2]: No" -ForegroundColor Yellow

$saveChoice = Read-Host "Enter the number for your choice"
Write-Host ""

switch ($saveChoice) {
    "1" {
        $filePath = Read-Host "Enter the file path where you want to save the results (e.g., C:\Results.csv)"
        Export-Csv -Path $filePath -InputObject $(Import-Csv -Path "SuspiciousActivity.csv") -NoTypeInformation
        Write-Host "Results saved to $filePath" -ForegroundColor Green
    }
    "2" { Write-Host "Results not saved to a CSV file." -ForegroundColor Green }
    default { Write-Host "Invalid choice. Results not saved to a CSV file." -ForegroundColor Red }
}
