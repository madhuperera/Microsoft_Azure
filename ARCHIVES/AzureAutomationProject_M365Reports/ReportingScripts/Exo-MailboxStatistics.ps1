# #---------------------------------------------------------[About]------------------------------------------------------------------------------
# Author: Madhu Perera
# -----------------------------------------------------------------------------------------------------------------------------------------------


# #---------------------------------------------------------[Help]-------------------------------------------------------------------------------

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
        [switch]$SendMailboxReport,
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
    Write-Output "Successfully connected to Exchange Online"
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

function Remove-SharePointListItems
{
    param($ListName)

    Write-Output "Removing SharePoint List Items"
    try
    {
        $CurrentItems = Get-PnPListItem -List $ListName -PageSize 5000 -Verbose -ErrorAction Stop
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

function Update-AllMailboxesSPOList
{
    param($ListName,$Items)
    Write-Output "Updating SharePoint List with latest data"
    try
    {
        foreach ($item in $Items)
        {
            Write-Output "Uploading data to SharePoint"
            Write-Output ($item.MailboxName)
            $TempOutputCatcher = Add-PnPListItem -List $ListName `
                -Values @{"Title" = $item.MailboxName; "PrimaryEmailAddress" = $item.PrimaryEmailAddress;`
                    'MailboxSizeInGB' = $item.MailboxSizeInGB ;`
                    'MailboxSizeInMB' = $item.MailboxSizeInMB  ;`
                    'MailboxCapacityInGB' = $item.MailboxCapacityInGB ;`
                    'MailboxUsage' = $item.MailboxUsage ;`
                    'LastMailboxAccessTime' = $item.LastLogonTime ;`
                    'LastMailboxAccessDate' = $item.LastLogonDate ;`
                    'MailboxIdentity' = $item.MailboxIdentity ;`
                    'ActiveInLast30Days' = $item.ActiveInLast30Days ;`
                    'OtherEmailAddresses' = $item.OtherEmailAddresses ;`
                    'ForwardingARecipient' = $item.ForwardingARecipient ;`
                    'ForwardingEmailAddress' = $Item.ForwardingEmailAddress ;`
                    'MailboxAuditingEnabled' = $item.MailboxAuditingEnabled ;`
                    'MailboxAuditHistory' = $item.MailboxAuditHistory ;`
                    'OutlookMobileEnabled' = $item.OutlookMobileEnabled ;`
                    'UniversalOutlookEnabled' = $item.UniversalOutlookEnabled ;`
                    'MAPIEnabled' = $item.MAPIEnabled ;`
                    'ImapEnabled' = $item.ImapEnabled ;`
                    'PopEnabled' = $item.PopEnabled ;`
                    'OWAEnabled' = $item.OWAEnabled ;`
                    'OwaMailboxPolicy' = $item.OwaMailboxPolicy ;`
                    'ActiveSyncEnabled' = $item.ActiveSyncEnabled ;`
                    'HasActiveSyncDevicePartnership' = $item.HasActiveSyncDevicePartnership ;`
                    'ActiveSyncMailboxPolicy' = $item.ActiveSyncMailboxPolicy ;`
                    'InboxRuleRecipients' = $item.InboxRuleRecipients  } `
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

# -----------------------------------------------------------------------------------------------------------------------------------------------


#-----------------------------------------------------------[Declaration]------------------------------------------------------------------------
$ReportSmtpServer = "CHANGE ME" # CHANGE ME Ex: smtp.gmail.com
$ReportSmtpPort = 25 # CHANGE ME Ex: 25
$ReportSmtpPSCredentialName = $ClientsServiceAccountName
$ReportSmtpFrom = "CHANGE ME" # CHANGE ME Ex: sender@gmail.com
$ReportSmtpTo = "CHANGE ME" # CHANGE ME Ex: receiver@outlook.com

# SendInboxRuleForwardingReport Variables
$SendMailboxReportSubject = "Azure Automation Report: Mailbox Report - " + $ClientName
$SendMailboxReportBody = "CSV file attached, containing Mailbox Report" 

# SharePoint Information
$SPOSiteUrl = "CHANGE ME" + $ClientName # CHANGE ME Ex: https://<your_tenancy>.sharepoint.com/sites/"
$AllMailboxesListName = "CHANGE ME" # CHANGE ME Ex: "AllMailboxes"
# -----------------------------------------------------------------------------------------------------------------------------------------------


#-----------------------------------------------------------[Execution]--------------------------------------------------------------------------
# Check if script is executed in Azure Automation
Test-AzureAutomationEnvironment


# Ackowledge the paramaters
Write-Output "::: Parameters :::"
Write-Output "ClientsServiceAccountName:    $ClientsServiceAccountName"
Write-Output "SharePointPSCredentialName:    $EITAutomationAccountName"
Write-Output "SendMailboxReport:             $SendMailboxReport"
Write-Output "EnableVerbose:                 $EnableVerbose"
Write-Output ""


# Handle Verbose Preference
if ($EnableVerbose -eq $true)
{
    $VerbosePreference = "Continue"
}


# Get AutomationPSCredential
Write-Output "::: Connection :::"
try
{
    Write-Output "Importing Automation Credential"
    $Credential = Get-AutomationPSCredential -Name $ClientsServiceAccountName -ErrorAction Stop
}
catch 
{
    Write-Error -Message $_.Exception
    Stop-AutomationScript -Status Failed
}
Write-Verbose "Successfully imported credentials"


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
Connect-ExchangeOnline -Credential $Credential -Commands "Get-Mailbox","Get-Recipient","Get-AcceptedDomain","Get-MailboxStatistics","Get-CASMailbox","Get-MobileDevice","Get-InboxRule"
Write-Output ""

# Getting a list of Internal Domains
$AllInternalDomains = (Get-AcceptedDomain).DomainName


# Import All Mailboxes
try
{
    Write-Verbose "Importing List of Mailboxes"
    $Mailboxes = Get-Mailbox -Filter "RecipientTypeDetails -eq 'UserMailbox' -or `
                                    RecipientTypeDetails -eq 'SharedMailbox' -or `
                                    RecipientTypeDetails -eq 'RoomMailbox' -or `
                                    RecipientTypeDetails -eq 'EquipmentMailbox'" `
                                    -ResultSize Unlimited -ErrorAction Stop
}
catch 
{
    Write-Error -Message $_.Exception
    Stop-AutomationScript -Status Failed
}

if (!$Mailboxes)
{
    $ErrorMessage = "No Mailboxes Found!"
    Stop-AutomationScript -Status Failed
}

Write-Verbose "Successfully Imported List of Mailboxes"
Write-Verbose ""

# Process All Mailboxes
Write-Verbose "Starting the process of Mailboxes"
Write-Verbose ""
$AllMailboxes = foreach ($Item in $Mailboxes)
{
    # Collecting Mailbox Statistics
    try
    {
        Write-Verbose "Collecting the Mailbox Statistics for "
        Write-Verbose ($item.DisplayName)
        Write-Verbose ""
        $ItemStatistics = Get-MailboxStatistics -Identity $Item.Id -ErrorAction SilentlyContinue
        $TotalItemSizeInBytes = $ItemStatistics.TotalItemSize.Value -replace "(.*\()|,| [a-z]*\)", ""
        $TotalItemSizeInMB = $TotalItemSizeInBytes / 1MB
        $TotalItemSizeInGB = $TotalItemSizeInBytes / 1GB
        $ProhibitSendReceiveQuota = $Item.ProhibitSendReceiveQuota -replace "(.*\()|,| [a-z]*\)", ""
        $ProhibitSendReceiveQuotaInMB = $ProhibitSendReceiveQuota / 1MB
        $ProhibitSendReceiveQuotaInGB = $ProhibitSendReceiveQuota / 1GB
        $MailboxUsage = ($TotalItemSizeInBytes / $ProhibitSendReceiveQuota)

        # Getting the last logon date and time
        $LastLogonTime = "No Time Recorded"
        If ($ItemStatistics.LastLogonTime)
        {
            $LastLogonTime = $ItemStatistics.LastLogonTime
            $LastLogonDate = '{0:MM/dd/yyyy HH:MM}' -f ($ItemStatistics.LastLogonTime)
        }
        else
        {
            $LastLogonDate = $item.WhenMailboxCreated
            $LastLogonDate = '{0:MM/dd/yyyy HH:MM}' -f ($LastLogonDate)
        }
       

        # Checking if the mailbox is logged on within last 30 days
        $ActiveWithinLast30Days = $false
        If ($ItemStatistics.LastLogonTime)
        {
            if (($ItemStatistics.LastLogonTime) -gt (Get-Date).AddDays(-30))
            {
                $ActiveWithinLast30Days = $true
            }
        }

        # Getting a list of Email Aliases
        $Aliases = $Item.EmailAddresses
        $OtherEmailAddresses = ""
        For ($i = 0; $i -le $Aliases.Count-1; $i++)
        {
            if (($Aliases[$i] -notlike "SIP*") -and ($Aliases[$i] -notlike "SPO*") -and ($Aliases[$i] -notlike "x500*") -and ($Aliases[$i] -notlike ("SMTP:" + $Item.PrimarySmtpAddress)))
            {
               # $OtherEmailAddresses = $OtherEmailAddresses + $Aliases[$i] + " | "
               if ($OtherEmailAddresses.Length -eq 0)
               {
                    $OtherEmailAddresses = ($Aliases[$i] -split ":")[1]
               }
               else
               {
                    $OtherEmailAddresses = $OtherEmailAddresses + " + "
                    $OtherEmailAddresses += ($Aliases[$i] -split ":")[1]
               }
            }
        }

        # Getting a Forwarding Address Details
        $ForwardingSmtpAddress = ""
        if ($Item.ForwardingSmtpAddress)
        {
            $ForwardingSmtpAddress = $Item.ForwardingSmtpAddress
        }

        $ForwardingRecipient = ""
        if ($Item.ForwardingAddress)
        {
            $ForwardingRecipient = $Item.ForwardingAddress
        }


        # Getting a Audit Details
        if ($item.AuditEnabled)
        {
            $MailboxAuditingEnabled = $true

            [int]$MailboxAuditHistory = ($item.AuditLogAgeLimit -split ":")[0]
        }
        else
        {
            $MailboxAuditingEnabled = $false
            [int]$MailboxAuditHistory = 0
        }

        # Getting Mailbox Protocol Status
        $ItemProtocols = ($item | Get-CASMailbox)

        # Processing Mailbox Inbox Rules
        $RuleExternalRecipientsString = ""
        $ItemRules = Get-InboxRule -Mailbox $item.PrimarySmtpAddress -WarningAction SilentlyContinue
        if ($ItemRules)
        {
            $ItemForwardingRules = $ItemRules | Where-Object {$_.ForwardTo -or $_.ForwardAsAttachmentTo}

            if ($ItemForwardingRules)
            {
                $RuleRecipients = @()
                $RuleExternalRecipients = @()
                foreach ($Rule in $ItemForwardingRules)
                {
                    # Collecting an array of recipients in a rule
                    
                    $RuleRecipients = $Rule.ForwardTo | Where-Object {$_ -match "SMTP"}
                    $RuleRecipients += $Rule.ForwardAsAttachmentTo | Where-Object {$_ -match "SMTP"}

                    # Collecting an array of external recipients in a rule
                    
                    foreach ($Recipient in $RuleRecipients)
                    {
                        $RuleEmailAddress = ($Recipient -split "SMTP:")[1].Trim("]")
                        $RuleEmailDomain = ($RuleEmailAddress -split "@")[1]

                        if ($AllInternalDomains -notcontains $RuleEmailDomain)
                        {
                            $RuleExternalRecipients += $RuleEmailAddress
                        }
                        else
                        {
                            Write-Verbose "$RuleEmailAddress - Internal Domain, Excluded Domain or Excluded Email Address"    
                        }

                    }

                    if ($RuleExternalRecipients)
                    {
                        $RuleExternalRecipientsString = $RuleExternalRecipients -join ", "
                    }
                }
            }
        }

        

        $props = @{
            'MailboxName' = $Item.DisplayName;
            'PrimaryEmailAddress' = $Item.PrimarySmtpAddress;
            'MailboxSizeInMB' = $TotalItemSizeInMB;
            'MailboxSizeInGB' = $TotalItemSizeInGB;
            'MailboxCapacityInMB' = $ProhibitSendReceiveQuotaInMB;
            'MailboxCapacityInGB' = $ProhibitSendReceiveQuotaInGB;
            'MailboxUsage' = $MailboxUsage;
            'MailboxIdentity' = $Item.Identity;
            'LastLogonTime' = $LastLogonTime;
            'LastLogonDate' = $LastLogonDate;
            'ActiveInLast30Days' = $ActiveWithinLast30Days;
            'OtherEmailAddresses' = $OtherEmailAddresses;
            'ForwardingEmailAddress' = $ForwardingSmtpAddress;
            'ForwardingARecipient' = $ForwardingRecipient;
            'MailboxAuditingEnabled' = $MailboxAuditingEnabled;
            'MailboxAuditHistory' = $MailboxAuditHistory;
            'OutlookMobileEnabled' = $ItemProtocols.OutlookMobileEnabled.toString();
            'UniversalOutlookEnabled' = $ItemProtocols.UniversalOutlookEnabled.toString();
            'MAPIEnabled' = $ItemProtocols.MAPIEnabled.toString();
            'ImapEnabled' = $ItemProtocols.ImapEnabled.toString();
            'PopEnabled' = $ItemProtocols.PopEnabled.toString();
            'OWAEnabled' = $ItemProtocols.OWAEnabled.toString();
            'OwaMailboxPolicy' = $ItemProtocols.OwaMailboxPolicy;
            'ActiveSyncEnabled' = $ItemProtocols.ActiveSyncEnabled.toString();
            'HasActiveSyncDevicePartnership' = $ItemProtocols.HasActiveSyncDevicePartnership.toString();
            'ActiveSyncMailboxPolicy' = $ItemProtocols.ActiveSyncMailboxPolicy;
            'InboxRuleRecipients' = $RuleExternalRecipientsString
        }
        New-Object -TypeName PSObject -Property $props
    }
    catch
    {
        Write-Error -Message $_.Exception
        Stop-AutomationScript -Status Failed
    }
    
}
Write-Verbose "Successfully Processed the List of Mailboxes"
Write-Verbose ""


# Import All Recipients
Write-Verbose "::: Import Exchange Recipient List :::"
try
{
    Write-Verbose "Importing List of Recipients"
    $Recipients = Get-Recipient -ResultSize Unlimited -ErrorAction Stop `
        | Select-Object Identity,RecipientType,ExternalEmailAddress
}
catch 
{
    Write-Error -Message $_.Exception
    Stop-AutomationScript -Status Failed
}
Write-Verbose "Successfully Imported List of Recipients"
Write-Verbose ""


# Connect to SharePoint Online
Connect-SharePointOnline -Credential $SPOCredential -Url $SPOSiteUrl
Write-Output ""



# Send Mail Report
if ($SendMailboxReport)
{
    Write-Output "::: Send Mail Report :::"
    
    $ReportTime = Get-Date -Format "MM-dd-yyyy_HH-mm-ss"

    Write-Output "Generate Mailbox Report CSV file"
    try
    {
        $CSVFileName = "MailboxReport_" + $ReportTime + ".csv"
        $CSVFilePath = $env:TEMP + "\" + $CSVFileName
        $AllMailboxes | Export-CSV -LiteralPath $CSVFilePath -Encoding Unicode -NoTypeInformation -Delimiter "`t" -ErrorAction Stop
    }
    catch 
    {
        Write-Error -Message $_.Exception
    }

    Write-Output "Send e-mail to '$ReportSmtpTo'"
    try
    {
        Send-MailMessage -To $ReportSmtpTo -From $ReportSmtpFrom -Subject $SendMailboxReportSubject `
                -Body $SendMailboxReportBody -BodyAsHtml -Attachments $CSVFilePath -SmtpServer $ReportSmtpServer `
                -Port $ReportSmtpPort -ErrorAction Stop
    }
    catch 
    {
        Write-Error -Message $_.Exception
    }
}

# Emptying the SharePoint List first
Remove-SharePointListItems -ListName $AllMailboxesListName

# Updating SharePoint
Update-AllMailboxesSPOList -ListName $AllMailboxesListName -Items $AllMailboxes

# Script Completed
Stop-AutomationScript -Status Success
# -----------------------------------------------------------------------------------------------------------------------------------------------