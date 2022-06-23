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
            Mandatory=$true)]
            [string]$ClientIPAddress,
    [Parameter(
        Mandatory=$false)]
        [switch]$SendO365FailedLoginsReport,
    [Parameter(
        Mandatory=$false)]
        [switch]$EnableVerbose,
    [Parameter(
        Mandatory=$true)]
        [int] $ReportPeriod=1,
    [Parameter(
        Mandatory=$true)]
        [int] $DaysToKeep=30
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


function Connect-ExchangeOnline 
{
    param ($Credential,$Commands)
    try
    {
        Write-Output "Connecting to Exchange Online"
        Get-PSSession | Remove-PSSession       
        $Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential `
        -Authentication Basic -AllowRedirection
        Import-PSSession -Session $Session -DisableNameChecking:$true -AllowClobber:$true -CommandName $Commands | Out-Null
    }
    catch 
    {
        Write-Error -Message $_.Exception
        Stop-AutomationScript -Status Failed
    }
    Write-Verbose "Successfully connected to Exchange Online"
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

function Disconnect-ExchangeOnline 
{
    try
    {
        Write-Output "Disconnecting from Exchange Online"
        Get-PSSession | Remove-PSSession       
    }
    catch 
    {
        Write-Error -Message $_.Exception
        Stop-AutomationScript -Status Failed
    }
    Write-Verbose "Successfully disconnected from Exchange Online"
}

function Disconnect-SharePointOnline
{
    try
    {
        Write-Output "Disconnecting from Exchange Online"
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
    Disconnect-ExchangeOnline
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

function Remove-OldSharePointListItems
{
    param($ListName,$Days)

    Write-Output "Removing SharePoint List Items in $ListName older than $Days"
    try
    {
        $CurrentItems = Get-PnPListItem -PageSize 5000 -List $ListName -Verbose -ErrorAction Stop
        foreach ($item in $CurrentItems)
        {
            if (  ($item.FieldValues.TimeStamp) -gt ((Get-Date).AddDays(-$Days))   )
            {
                Write-Output "Item ($item.Id) is not $Days days old"
            }
            else 
            {
                Remove-PnPListItem -Identity $item.Id -List $ListName -Force -ErrorAction Stop
            }
        }
    }
    catch
    {
        Write-Error -Message $_.Exception
        Stop-AutomationScript -Status Failed
    }
    Write-Output "Successfully Emptied SharePoint List"
}

function Update-Office365FailedLoginsSPOList
{
    param($ListName,$Items)
    Write-Output "Updating SharePoint List - Office 365 Failed Logins - with the latest data"
    try
    {
        foreach ($item in $Items)
        {
            $TempOutputCatcher = Add-PnPListItem -List $ListName `
                -Values @{"Title" = $item.CreationTime; "User" = $item.User;`
                    'Action' = $item.Action ;`
                    'TimeStamp' = $item.TimeStamp ;`
                    'Status' = $item.Status  ;`
                    'Actor' = $item.Actor ;`
                    'ClientIP' = $item.ClientIP ;`
                    'ActorIPAddress' = $item.ActorIPAddress ;`
                    'Reason' = $item.Reason ;`
                    'City' = $item.City ;`
                    'Country' = $Item.Country } `
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

function Convert-TimeStampToSPODateTime
{
    param($DateText)
    $SPODateTime = ""
    try
    {
        $tsDate = ($DateText -split "T")[0]
        $tsTime = ($DateText -split "T")[1]

        $tsYear = ($tsDate -split "-")[0]
        $tsMonth = ($tsDate -split "-")[1]
        $tsDay = ($tsDate -split "-")[2]

        $tsHour = ($tsTime -split ":")[0]
        $tsMinute = ($tsTime -split ":")[1]

        $SPODateTime = $tsMonth + "/" + $tsDay + "/" + $tsYear + " " + $tsHour + ":" + $tsMinute
        $SPODateTime
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
$IPStackAccessKey = "CHANGE ME" # CHANGE ME Ex: Access Key from IP Stack Website

# SendOffice365GroupsReport Variables
$SendO365FailedLoginsReportSubject = "Azure Automation Report: Office 365 Failed Logins Report - " + $ClientName
$SendO365FailedLoginsReportBody = "CSV file attached, containing Office 365 Failed Logins Report" 

# SharePoint Information
$SPOSiteUrl = "CHANGE ME" + $ClientName # CHANGE ME Ex: https://<your_tenancy>.sharepoint.com/sites/"
$Office365FailedLoginsListName = "CHANGE ME" # CHANGE ME Ex: "Office365FailedLogins"

# -----------------------------------------------------------------------------------------------------------------------------------------------


#-----------------------------------------------------------[Execution]--------------------------------------------------------------------------
Test-AzureAutomationEnvironment


# Ackowledge the paramaters
Write-Output "::: Parameters :::"
Write-Output "ClientsServiceAccountName:    $ClientsServiceAccountName"
Write-Output "SharePointPSCredentialName:    $EITAutomationAccountName"
Write-Output "SendO365FailedLoginsReport:    $SendO365FailedLoginsReport"
Write-Output "EnableVerbose:                 $EnableVerbose"
Write-Output "ReportPeriod:                  $ReportPeriod"
Write-Output ""


# Handle Verbose Preference
if ($EnableVerbose -eq $true)
{
    $VerbosePreference = "Continue"
}


# Get AutomationPSCredential
Write-Output "::: Exchange Connection :::"
try
{
    Write-Output "Importing Exchange Online and Azure AD Credentials"
    $ExOCredential = Get-AutomationPSCredential -Name $ClientsServiceAccountName -ErrorAction Stop
}
catch 
{
    Write-Error -Message $_.Exception
    Stop-AutomationScript -Status Failed
}
Write-Verbose "Successfully imported Exchange Online and Azure AD Credentials"


# Get SPOPSCredential
Write-Output "::: Connection :::"
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
Write-Verbose "Successfully imported credentials"


# Connect to Exchange Online
Connect-ExchangeOnline -Credential $ExOCredential -Commands "Get-UnifiedGroup","Search-UnifiedAuditLog","Get-UnifiedGroupLinks","Get-MailboxFolderStatistics"
Write-Output ""

$ToDate = (Get-Date).ToString("dd-MMMM-yyyy")
$FromDate = ((Get-Date).AddDays(-$ReportPeriod)).ToString("dd-MMMM-yyyy")

try
{
    $Records = (Search-UnifiedAuditLog -StartDate $FromDate -EndDate $ToDate -Operations "UserLoginFailed" -ResultSize 5000 -ErrorAction Stop)    
}
catch
{
    Write-Error -Message $_.Exception
    Stop-AutomationScript -Status Failed
}

If ($Records.Count -eq 0)
{
    Write-Output "No group creation records found." 
}
else
{
    try
    {
        # Get the IP Location for Customer's Primary IP Address
        Write-Verbose "Getting the IP Information for Client Main IP Address"
        $APIUri = "http://api.ipstack.com/" + $ClientIPAddress + $IPStackAccessKey
        $IPResults = Invoke-restmethod -method get -uri $APIUri

        # Create a New PowerShell Object and assign the details of client IP address
        $IPObject = @{
            "IPAddress" = $IPResults.IP;
            "City" = $IPResults.City;
            "Country" = $IPResults.country_name
        }
        New-Object -TypeName psobject -Property $IPObject

        # Create an Array to store a known IP list
        $KnownIpObjectsArray = @()
        $KnownIpObjectsArray += $IPObject

        # Create an Array to store a known IP list
        $KnownIpObjectsArray = @()
        $KnownIpObjectsArray += $IPObject

        # Keeping track of number of API Calls to IP Stack
        $TotalNumOfKnownRequests = 0
        $TotalNumOfAPICalls = 0

        $AllFailedLogins = foreach ($Rec in $Records)
        {
            $AuditData = ConvertFrom-Json $Rec.Auditdata

            $TimeStamp = Convert-TimeStampToSPODateTime -DateText $AuditData.CreationTime
            $ClientIP = $AuditData.ClientIP

            # Checking if the IP Address has already been checked
            if ( $KnownIpObjectsArray.IPAddress -contains $ClientIP )
            {
                Write-Verbose "$ClientIP is already known by the script...."
                $IPObjectDetails = $KnownIpObjectsArray | Where-Object {$_.IPAddress -eq $ClientIP}
                $CountryName = $IPObjectDetails.Country
                $CityName = $IPObjectDetails.City

                $TotalNumOfKnownRequests += 1
            }
            else
            {
                # Getting the IP Address Location Details
                Write-Verbose "$ClientIP : Connecting to IP Stack for IP Location Data..."
                $APIUri = "http://api.ipstack.com/" + $ClientIP + $IPStackAccessKey
                $IPResults = Invoke-restmethod -method get -uri $APIUri
                $CountryName = $IPResults.country_name
                $CityName = $IPResults.city

                #Creating a new object with the new details
                $IPObject = @{
                    "IPAddress" = $IPResults.IP;
                    "City" = $IPResults.City;
                    "Country" = $IPResults.country_name
                }
                $temp = New-Object -TypeName psobject -Property $IPObject

                # Adding to the known IP list
                $KnownIpObjectsArray = @()
                $KnownIpObjectsArray += $IPObject

                $TotalNumOfAPICalls += 1
            }
           
            $props = @{
                'TimeStamp'   = $TimeStamp;
                'CreationTime'= $AuditData.CreationTime;
                'User'        = $AuditData.UserId;
                'Action'      = $AuditData.Operation;
                'Status'      = $AuditData.ResultStatus;
                'Actor'       = $AuditData.Actor[1].Id;
                'ClientIP'    = $ClientIP;
                'City'        = $CityName;
                'Country'     = $CountryName;
                'ActorIPAddress' = $AuditData.ActorIpAddress;
                'Reason'      = $AuditData.LogonError
            }
        New-Object -TypeName PSObject -Property $props
        }

        Write-Output "Total Number of Known IP Locations: $TotalNumOfKnownRequests"
        Write-Output "Total Number of New API Calls: $TotalNumOfAPICalls"
        Write-Output "Total Number of IP Addresses Checked: $($AllFailedLogins.Count)"
    }
    catch
    {
        Write-Error -Message $_.Exception
        Stop-AutomationScript -Status Failed
    }
}

# Send Mail Report
if ($SendO365FailedLoginsReport)
{
    Write-Output "::: Send Failed Login Report Report :::"
    
    $ReportTime = Get-Date -Format "MM-dd-yyyy_HH-mm-ss"

    Write-Output "Generate Failed Logins Report CSV file"
    try
    {
        $CSVFileName = "O365FailedLogins_" + $ReportTime + ".csv"
        $CSVFilePath = $env:TEMP + "\" + $CSVFileName
        $AllFailedLogins | Export-CSV -LiteralPath $CSVFilePath -Encoding Unicode -NoTypeInformation -Delimiter "`t" -ErrorAction Stop
    }
    catch 
    {
        Write-Error -Message $_.Exception
    }

    Write-Output "Send e-mail to '$ReportSmtpTo'"
    try
    {
        Send-MailMessage -To $ReportSmtpTo -From $ReportSmtpFrom -Subject $SendO365FailedLoginsReportSubject `
                -Body $SendO365FailedLoginsReportBody -BodyAsHtml -Attachments $CSVFilePath -SmtpServer $ReportSmtpServer `
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
Remove-OldSharePointListItems -ListName $Office365FailedLoginsListName -Days $DaysToKeep

# Updating SharePoint
If ($Records.Count -ne 0)
{
    Update-Office365FailedLoginsSPOList -ListName $Office365FailedLoginsListName -Items $AllFailedLogins
}
else
{
    Write-Output "Skipping item update to SharePoint as not items were found"     
}


# Script Completed
Stop-AutomationScript -Status Success
# -----------------------------------------------------------------------------------------------------------------------------------------------