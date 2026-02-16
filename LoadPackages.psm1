
Function Load-Packages {
    param ([string] $directory = '.\')
    
    try {
        if(-not (Get-Module -ListAvailable -Name "Microsoft.Graph.Mail")) {
            Install-Module -Name "Microsoft.Graph.Mail" -Scope CurrentUser
            Import-Module -Name Microsoft.Graph.Mail
        }
        if(-not (Get-Module -ListAvailable -Name "MSAL.PS")) {
            Install-Module -Name "MSAL.PS" -Scope CurrentUser
            Import-Module -Name MSAL.PS
        }
        if(-not (Get-Module -ListAvailable -Name "Microsoft.Xrm.Tooling.CrmConnector.PowerShell")) {
            Install-Module -Name Microsoft.Xrm.Tooling.CrmConnector.PowerShell
        }
        if(-not (Get-Module -ListAvailable -Name "Microsoft.PowerApps.Administration.PowerShell")) {
            #Install-Module -Name Microsoft.PowerApps.Administration.PowerShell
        }
        #Add-Type -Path ".\bin\Microsoft.Xrm.Sdk.dll"
    }
    catch [System.Exception] {
        Write-Host "Message: $($_.Exception.Message)"
        Write-Host "StackTrace: $($_.Exception.StackTrace)"
        Write-Host "LoaderExceptions: $($_.Exception.LoaderExceptions)"
    }
}

Export-ModuleMember -Function Load-Packages