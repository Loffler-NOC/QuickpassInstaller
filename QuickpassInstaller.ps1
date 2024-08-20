#Check if quickpass is already installed
$programName = "Quickpass Agent"

# Check if the program is installed
$installed = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -eq $programName }

if ($installed -ne $null) {
    Write-Output "$programName is already installed."
    exit 0
} else {
    # Check 32-bit registry if not found in 64-bit
    $installed = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -eq $programName }
    if ($installed -ne $null) {
        Write-Output "$programName is already installed."
        exit 0
    } else {
        Write-Output "$programName is not installed. Continuing to installation."
    }
}

if (!$env:QPAgentID) {
    Write-Output "Quickpass site ID not found, please check site variables for QPAgentID and make sure it exists and is filled out"
    Write-Output "Quickpass site ID reporting as $env:QPAgentID"
    exit 1
}
elseif ($env:QPAgentID -eq 0) {
    Write-Output "Quickpass site ID still has the default value of 0, please check site variables for QPAgentID and make sure it exists and is filled out"
    exit 1
}
else {
    Write-Output "Quickpass ID is $env:QPAgentID, proceeding to next step"
}

##Quickpass Installation PowerShell Script


$Path = "."
$DownloadURL = "https://storage.googleapis.com/qp-installer/production/Quickpass-Agent-Setup.exe"
$Output = $path + "\Quickpass-Agent-Setup.exe"


#Edit These Values for your Install Token and Agent ID Inside quotation Marks


$QPInstallTokenID = "$env:QPInstallTokenID"
$QPAgentID = "$env:QPAgentID"

#Edit RegionID for EU Tenant ONLY
#RegionID = "EU" for EU Tenant
#RegionID = "NA" for North America/Oceania Tenant

$RegionID = "NA"

#adds quotes to Installation Parameter

$QPInstallTokenIDBLQt = """$QPInstallTokenID""" 
$QPAgentIDDBlQt = """$QPAgentID"""
$Region = """$RegionID"""

#Restart Options

<#Restart Commands 
.NET lower than 4.7.2 
.NET 4.7.2 or Higher Already Installed

No value Specified 
After installation of .NET completes the system will automatically be restarted & After admin login, installation of the Agent system will complete and system will NOT be rebooted 
After installation of the Agent system will NOT be rebooted

/NORESTART 
After installation of .NET completes the system will NOT automatically be restarted & After admin login, installation of the Agent will complete and system will NOT be rebooted 
After installation of the Agent system will NOT be rebooted

/FORCERESTART 
After installation of .NET completes the system will automatically be restarted & After admin login, installation of the Agent will complete and system will NOT be rebooted 
After installation of the Agent system will NOT be rebooted

RESTART=1
After installation of .NET completes the system will automatically be restarted & After admin login, installation of the Agent will complete and system will be rebooted 
After installation of the Agent system will be rebooted
#>


$RestartOption = "/NORESTART"



#MSA vs Local System Service Options
<#MSA Commands

No Value Specified
The Agent will use the Local System Account to run the service

MSA=0
The Agent will use the Local System Account to run the service

MSA=1
A Managed Service Account will be created to run the Service 
NOTE: This is only used for Domain Controllers.  All other system types this command will be ignored.

#>

$MSAOption = "MSA=1"





#Test if download destination folder exists, create folder if required
If(Test-Path $Path)
{write-host "Destination folder exists"}else{
#Create Directory to download quickpass installer into
write-host "Creating folder $Path"
md $Path
}

#Begin download of Quickpass Agent
write-host "Beginning download of the quickpass agent"
Start-BitsTransfer -Source $DownloadURL -Destination $Output
write-host "Variables in use for Quickpass Agent installation"
write-host "Software Path: $Output"
write-host "Installation Token: $QPInstallTokenID"
write-host "Customer ID $QPAgentID"
write-host "Restart option Selected $RestartOption"
write-host "MSA Creation Selected $MSAOption"

write-host "Beginning installation of Quickpass"


Try
{
Start-Process "$Output" -ArgumentList "/quiet $RestartOption INSTALLTOKEN=$QPInstallTokenIDBLQt CUSTOMERID=$QPAgentIDDBlQt REGION=$Region $MSAOption" -ErrorAction Stop -Wait -NoNewWindow
}
Catch
{
$ErrorMessage = $_.Exception.Message
write-host "Install error was: $ErrorMessage"
exit 1
}

#Check if quickpass was successfully installed
$programName = "Quickpass Agent"

# Check if the program is installed
$installed = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -eq $programName }

if ($installed -ne $null) {
    Write-Output "$programName successfully installed."
    exit 0
} else {
    # Check 32-bit registry if not found in 64-bit
    $installed = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -eq $programName }
    if ($installed -ne $null) {
        Write-Output "$programName successfully installed."
        exit 0
    } else {
        Write-Output "$programName failed to install"
        exit 1
    }
}
