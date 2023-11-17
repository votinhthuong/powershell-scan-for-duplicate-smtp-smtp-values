# PowerShell script to find duplicate SMTP addresses in Active Directory and output to CSV

# Ensure the ActiveDirectory module is available
try {
    $moduleName = "ActiveDirectory"
    $module = Get-Module -ListAvailable -Name $moduleName

    if (-not $module) {
        Write-Host "ActiveDirectory module is not available. Attempting to install..."
        
        # Attempt to install the module
        Install-Module -Name $moduleName -Force -AllowClobber

        # Recheck if the module is installed
        $module = Get-Module -ListAvailable -Name $moduleName
        if (-not $module) {
            throw "Failed to install ActiveDirectory module."
        }
    }

    # Import the module
    Import-Module $moduleName -ErrorAction Stop
    Write-Host "`n"
    Write-Host "ActiveDirectory module loaded successfully."
    Write-Host "`n"
} catch {
    Write-Error "An error occurred: $_"
    exit
}
# Script for validating Active Directory credentials and attempting a connection

try {
    # Input Domain
    $domain = Read-Host -Prompt "Enter the Active Directory domain"
    if (-not $domain -or $domain -notmatch '^[A-Za-z0-9.-]+\.[A-Za-z]{2,}$') {
        throw "Invalid domain format!"
    }

    # Input Username
    Write-Host "`n"
    $username = Read-Host -Prompt "Enter your username (DOMAIN\username)"
    if (-not $username -or $username -notmatch '^[A-Za-z0-9.-]+\\[A-Za-z0-9.-]+$') {
        throw "Invalid username format! Use format DOMAIN\\username."
    }

    # Input Password
    $password = Read-Host -Prompt "Enter your password" -AsSecureString
    if (-not $password) {
        throw "No password entered!"
    }

    # Create credential object
    $credential = New-Object System.Management.Automation.PSCredential ($username, $password)

    # ... previous parts of the script ...

    # Attempt to connect to Active Directory
    try {
        Import-Module ActiveDirectory -ErrorAction Stop

        # Test connection to Active Directory (suppress output)
        $null = Get-ADUser -Filter {Name -like "*"} -Credential $credential -ErrorAction Stop
        Write-Host "`n"
        Write-Host "Successfully connected to Active Directory domain '$domain' with provided credentials."
    } catch {
        throw "Failed to connect to Active Directory domain '$domain' with provided credentials. Error: $_"
    }

    # ... rest of the script ...


} catch {
    Write-Error "An error occurred: $_"
    exit
}
Write-Host "`n"
# User Input for Output File Path
try {
    $outputFilePath = Read-Host -Prompt "Enter the full path for the output CSV file"
    if (-not $outputFilePath) {
        throw "No output file path entered"
    }
} catch {
    Write-Error "An error occurred: $_"
    exit
}

# Scan for Duplicate SMTP/smtp Values and Output Results
try {
    $allUsers = Get-ADUser -Filter * -Properties DisplayName, ProxyAddresses
    $smtpAddresses = $allUsers | Select-Object -ExpandProperty ProxyAddresses | Where-Object {$_ -clike "SMTP:*" -or $_ -clike "smtp:*"}
    $duplicateSMTP = $smtpAddresses | Group-Object | Where-Object { $_.Count -gt 1 }
    $results = @()

    foreach ($dup in $duplicateSMTP) {
        $users = $allUsers | Where-Object { $_.ProxyAddresses -contains $dup.Name }
        foreach ($user in $users) {
            $output = "Duplicate SMTP/smtp: $($dup.Name) - Display Name: $($user.DisplayName), Username: $($user.SamAccountName)"
            Write-Host "`n"
            Write-Host $output
            $results += New-Object PSObject -Property @{
                DuplicateSMTP = $dup.Name
                DisplayName = $user.DisplayName
                Username = $user.SamAccountName
            }
        }
    }

    # Write results to CSV file
    $results | Export-Csv -Path $outputFilePath -NoTypeInformation
    Write-Host "`n"
    Write-Host "Results have been written to $outputFilePath"

} catch {
    Write-Error "Error processing data: $_"
    exit
}
Write-Host "`n"
Write-Host "Script execution completed!"
