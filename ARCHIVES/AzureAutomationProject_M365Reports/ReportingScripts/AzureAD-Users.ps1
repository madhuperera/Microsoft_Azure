# #---------------------------------------------------------[About]------------------------------------------------------------------------------
# Quick Description of the author etc
# -----------------------------------------------------------------------------------------------------------------------------------------------


# #---------------------------------------------------------[Help]-------------------------------------------------------------------------------
<#
#>
# -----------------------------------------------------------------------------------------------------------------------------------------------


# #---------------------------------------------------------[Paramters]--------------------------------------------------------------------------
[cmdletbinding()]

param (
    [Parameter(
        Mandatory=$true)]
        [string]$ClientsServiceAccountName,
    [Parameter(
        Mandatory=$true)]
        [string]$EITAutomationAccountName,
    [Parameter(
            Mandatory=$true)]
            [string]$ClientName,
    [Parameter(
        Mandatory=$false)]
        [switch]$SendO365LicensingReport,
    [Parameter(
        Mandatory=$false)]
        [switch]$EnableVerbose
)
# -----------------------------------------------------------------------------------------------------------------------------------------------


#-----------------------------------------------------------[Function]---------------------------------------------------------------------------
function Test-AzureAutomationEnvironment
{
    if ($env:AUTOMATION_ASSET_ACCOUNTID)
    {
        Write-Verbose "This script is executed in Azure Automation"
    }
    else
    {
        $ErrorMessage = "This script is NOT executed in Azure Automation."
        throw $ErrorMessage
    }
}


function Connect-Office365Online 
{
    param ($Credential)
    try
    {
        Write-Output "Connecting to Office 365 Online"
        Connect-MsolService -Credential $Credential
    }
    catch 
    {
        Write-Error -Message $_.Exception
        Stop-AutomationScript -Status Failed
    }
    Write-Verbose "Successfully connected to Office 365 Online"
}

function Connect-SharePointOnline
{
    param ($Credential, $Url)  
    try
    {
        Write-Output "Connecting to SharePoint Online"   
        Connect-PnPOnline -Url $Url -Credentials $Credential -ErrorAction Stop
    }
    catch
    {
        Write-Error -Message $_.Exception
        Stop-AutomationScript -Status Failed
    }
}

function Disconnect-SharePointOnline
{
    try
    {
        Write-Output "Disconnecting from SharePoint Online"
        Disconnect-PnPOnline
    }
    catch
    {
        Write-Error -Message $_.Exception
        Stop-AutomationScript -Status Failed
    }
    Write-Verbose "Successfully disconnected from SharePoint Online"
}

function Stop-AutomationScript
{
    param(
        [ValidateSet("Failed","Success")]
        [string]
        $Status = "Success"
        )
    Write-Output ""
    Disconnect-SharePointOnline
    if ($Status -eq "Success")
    {
        Write-Output "Script successfully completed"
    }
    elseif ($Status -eq "Failed")
    {
        Write-Output "Script stopped with an Error"
    }
    Break
}

function Remove-TemporaryFiles
{
    param($FileToBeDeleted)

    try
    {
        Write-Output "Removing temporary files..."
        Remove-Item -LiteralPath $FileToBeDeleted -Force -ErrorAction Stop
    }
    catch
    {
        Write-Error -Message $_.Exception
        Stop-AutomationScript -Status Failed
    }
    Write-Output "Successfully removed temporary files..."
}

function Remove-SharePointListItems
{
    param($ListName)

    Write-Output "Removing SharePoint List Items"
    try
    {
        $CurrentItems = Get-PnPListItem -List $ListName -Verbose -ErrorAction Stop
        foreach ($item in $CurrentItems)
        {
            Remove-PnPListItem -Identity $item.Id -List $ListName -Force -ErrorAction Stop
        }
    }
    catch
    {
        Write-Error -Message $_.Exception
        Stop-AutomationScript -Status Failed
    }
    Write-Output "Successfully Emptied SharePoint List"
}

function Update-AzureADUsersSPOList
{
    param($ListName,$Items)
    Write-Output "Updating SharePoint List - Azure AD Users - with the latest data"
    try
    {
        foreach ($item in $Items)
        {
            $TempOutputCatcher = Add-PnPListItem -List $ListName `
                -Values @{"Title" = $item.UserPrincipalName; "DisplayName" = $item.DisplayName;`
                    'Licenses' = $item.Licenses ;`
                    'IsLicensed' = $item.IsLicensed ;`
                    'UserType' = $item.UserType  ;`
                    'BlockedSignIn' = $item.BlockedSignIn ;`
                    'MFAStatus' = $item.MFAStatus ;`
                    'DefaultMFAMethod' = $item.DefaultMFAMethod ;`
                    'MFAPhoneNumber' = $item.MFAPhoneNumber ;`
                    'MFAEmail' = $item.MFAEmail ;`
                    'MFAAlternativePhoneNumber' = $item.MFAAlternativePhoneNumber    } `
                -Verbose -ErrorAction Stop
            $TempOutputCatcher = ""
        }
        Write-Output "Uploading items to SharePoint List completed successfully"
    }
    catch
    {
        Write-Error -Message $_.Exception
        Stop-AutomationScript -Status Failed
    }
}

function Update-MFAStatusTableSPOList 
{
    param($ListName,$MFAEnforced,$MFAEnabled,$MFADisabled)
    Write-Output "Updating SharePoint List - MFA Status Table - with the latest data"
    try
    {
        
        $TempOutputCatcher = Add-PnPListItem -List $ListName `
            -Values @{"Title" = "Enforced"; "Count" = $MFAEnforced }`
            -Verbose -ErrorAction Stop
        $TempOutputCatcher = ""
        
        $TempOutputCatcher = Add-PnPListItem -List $ListName `
            -Values @{"Title" = "Enabled"; "Count" = $MFAEnabled }`
            -Verbose -ErrorAction Stop
        $TempOutputCatcher = ""

        $TempOutputCatcher = Add-PnPListItem -List $ListName `
            -Values @{"Title" = "Disabled"; "Count" = $MFADisabled }`
            -Verbose -ErrorAction Stop
        $TempOutputCatcher = ""

        Write-Output "Uploading items to SharePoint List completed successfully"
    }
    catch
    {
        Write-Error -Message $_.Exception
        Stop-AutomationScript -Status Failed
    }
}

function Update-UserTypeTableSPOList
{
    param($ListName,$NumMembers,$NumGuests)
    Write-Output "Updating SharePoint List - User Type Table - with the latest data"
    try
    {
        
        $TempOutputCatcher = Add-PnPListItem -List $ListName `
            -Values @{"Title" = "Guest"; "Count" = $NumGuests }`
            -Verbose -ErrorAction Stop
        $TempOutputCatcher = ""
        
        $TempOutputCatcher = Add-PnPListItem -List $ListName `
            -Values @{"Title" = "Member"; "Count" = $NumMembers }`
            -Verbose -ErrorAction Stop
        $TempOutputCatcher = ""

        Write-Output "Uploading items to SharePoint List completed successfully"
    }
    catch
    {
        Write-Error -Message $_.Exception
        Stop-AutomationScript -Status Failed
    }
}


# -----------------------------------------------------------------------------------------------------------------------------------------------


#-----------------------------------------------------------[Declaration]------------------------------------------------------------------------
$ReportSmtpServer = "CHANGE ME" # CHANGE ME Ex: smtp.gmail.com
$ReportSmtpPort = 25 # CHANGE ME Ex: 25
$ReportSmtpPSCredentialName = $ClientsServiceAccountName
$ReportSmtpFrom = "CHANGE ME" # CHANGE ME Ex: sender@gmail.com
$ReportSmtpTo = "CHANGE ME" # CHANGE ME Ex: receiver@outlook.com

# SendOffice365GroupsReport Variables
$SendO365LicensingReportSubject = "Azure Automation Report: Office 365 Licensing Report - " + $ClientName
$SendO365LicensingReportBody = "CSV file attached, containing Office 365 Licensing Report" 

# SharePoint Information
$SPOSiteUrl = "CHANGE ME" + $ClientName # CHANGE ME Ex: https://<your_tenancy>.sharepoint.com/sites/"
$AzureADUsersListName = "CHANGE ME" # CHANGE ME Ex: "AzureADUsers"
$MFAStatusTableListName = "CHANGE ME" # CHANGE ME Ex: "MFAStatusTable"
$UserTypeTableListName = "CHANGE ME" # CHANGE ME Ex: "UserTypeTable"

# -----------------------------------------------------------------------------------------------------------------------------------------------


#-----------------------------------------------------------[Execution]--------------------------------------------------------------------------
Test-AzureAutomationEnvironment


# Ackowledge the paramaters
Write-Output "::: Parameters :::"
Write-Output "ClientsServiceAccountName:        $ClientsServiceAccountName"
Write-Output "SharePointPSCredentialName:       $EITAutomationAccountName"
Write-Output "SendO365LicensingReport:          $SendO365LicensingReport"
Write-Output "EnableVerbose:                    $EnableVerbose"
Write-Output ""


# Handle Verbose Preference
if ($EnableVerbose -eq $true)
{
    $VerbosePreference = "Continue"
}


# Get AutomationPSCredential
Write-Output "::: Microsoft Online Connection :::"
try
{
    Write-Output "Importing Microsoft Online"
    $O365Credential = Get-AutomationPSCredential -Name $ClientsServiceAccountName -ErrorAction Stop
}
catch 
{
    Write-Error -Message $_.Exception
    Stop-AutomationScript -Status Failed
}
Write-Verbose "Successfully imported Microsoft Online Credentials"


# Get SPOPSCredential
Write-Output "::: SharePoint Online Connection :::"
try
{
    Write-Output "Importing SharePoint Online Automation Credential"
    $SPOCredential = Get-AutomationPSCredential -Name $EITAutomationAccountName -ErrorAction Stop
}
catch 
{
    Write-Error -Message $_.Exception
    Stop-AutomationScript -Status Failed
}
Write-Verbose "Successfully imported SharePoint Online Automation credentials"


# Connect to Office 365 Online
Connect-Office365Online -Credential $O365Credential
Write-Output ""

# Initializing MFA Status Variables
$MFAEnabled = 0
$MFAEnforced = 0
$MFADisabled = 0

# Initializing User Type Variables
$NumOfGuestUsers = 0
$NumOfMembers = 0

try
{
    $AllUsers = Get-MsolUser -All -ErrorAction Stop

    $AzureADUserData = foreach ($User in $AllUsers)
    {
        # Initialize Variables
        $UserLicensesString = ""   
        $MFAStatus = "Disabled"
        $DefaultMFAMethod = ""
        $MFAPhoneNumber = ""
        $MFAEmail = ""
        $MFAAlternativePhoneNumber = ""
        
        # Retrieving Licensing Details
        if ($User.IsLicensed)
        {
            $UserLicenses = $User.Licenses
            $UserLicensesArray = $UserLicenses | ForEach-Object { ($_.AccountSkuId -split ":")[1]  }
            $UserLicensesString = $UserLicensesArray -join ", "
        }
        
        # Retrieving MFA Details
        If ($User.StrongAuthenticationRequirements.State -ne $null)
        {
            $MFAStatus = $User.StrongAuthenticationRequirements.State
            $DefaultMFAMethod = ($User.StrongAuthenticationMethods | Where-Object {$_.IsDefault -eq $true}).MethodType
            switch ($MFAStatus)
            {
                "Enforced"  { $MFAEnforced += 1 }
                "Enabled"   { $MFAEnabled += 1 }
            }
        }
        else
        {
            $MFADisabled += 1     
        }

        if ($User.StrongAuthenticationUserDetails.PhoneNumber)
        {
            $MFAPhoneNumber = $User.StrongAuthenticationUserDetails.PhoneNumber
        }

        if ($User.StrongAuthenticationUserDetails.Email)
        {
            $MFAEmail = $User.StrongAuthenticationUserDetails.Email
        }

        if ($User.StrongAuthenticationUserDetails.AlternativePhoneNumber)
        {
            $MFAAlternativePhoneNumber = $User.StrongAuthenticationUserDetails.AlternativePhoneNumber
        }

        # Counting Number of Guest and Internal Users
        $UserType = $User.UserType.toString()
        If ($UserType -eq "Guest")
        {
            $NumOfGuestUsers += 1
        }
        else
        {
            $NumOfMembers += 1    
        }


        $ULProps = @{'DisplayName' = $User.DisplayName;
                        'Licenses' = $UserLicensesString;
                        'UserPrincipalName' = $User.UserPrincipalName;
                        'IsLicensed' = $User.IsLicensed;
                        'UserType' = $UserType;
                        'BlockedSignIn' = $User.BlockCredential;
                        'MFAStatus' = $MFAStatus;
                        'DefaultMFAMethod' = $DefaultMFAMethod;
                        'MFAPhoneNumber ' = $MFAPhoneNumber;
                        'MFAEmail' = $MFAEmail;
                        'MFAAlternativePhoneNumber' = $MFAAlternativePhoneNumber   }
            New-Object -TypeName PSObject -Property $ULProps
    }
}
catch
{
    Write-Error -Message $_.Exception
    Stop-AutomationScript -Status Failed
}


# Send Mail Report
if ($SendO365LicensingReport)
{
    Write-Output "::: Send Office 365 Licensing Report :::"
    
    $ReportTime = Get-Date -Format "MM-dd-yyyy_HH-mm-ss"

    Write-Output "Generate Office 365 Licensing Report CSV file"
    try
    {
        $CSVFileName = "O365Licensing_" + $ReportTime + ".csv"
        $CSVFilePath = $env:TEMP + "\" + $CSVFileName
        $AzureADUserData | Export-CSV -LiteralPath $CSVFilePath -Encoding Unicode -NoTypeInformation -Delimiter "`t" -ErrorAction Stop
    }
    catch 
    {
        Write-Error -Message $_.Exception
    }

    Write-Output "Send e-mail to '$ReportSmtpTo'"
    try
    {
        Send-MailMessage -To $ReportSmtpTo -From $ReportSmtpFrom -Subject $SendO365LicensingReportSubject `
                -Body $SendO365LicensingReportBody -BodyAsHtml -Attachments $CSVFilePath -SmtpServer $ReportSmtpServer `
                -Port $ReportSmtpPort -ErrorAction Stop
    }
    catch 
    {
        Write-Error -Message $_.Exception
    }
}

# Connect to SharePoint Online
Connect-SharePointOnline -Credential $SPOCredential -Url $SPOSiteUrl
Write-Output ""

# Emptying the SharePoint List first
Remove-SharePointListItems -ListName $AzureADUsersListName
Remove-SharePointListItems -ListName $MFAStatusTableListName
Remove-SharePointListItems -ListName $UserTypeTableListName

# Updating SharePoint
Update-AzureADUsersSPOList -ListName $AzureADUsersListName -Items $AzureADUserData
Update-MFAStatusTableSPOList -ListName $MFAStatusTableListName -MFAEnforced $MFAEnforced -MFAEnabled $MFAEnabled -MFADisabled $MFADisabled
Update-UserTypeTableSPOList -ListName $UserTypeTableListName -NumMembers $NumOfMembers -NumGuests $NumOfGuestUsers

# Script Completed
Stop-AutomationScript -Status Success
# -----------------------------------------------------------------------------------------------------------------------------------------------