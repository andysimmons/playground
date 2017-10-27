#requires -Version 5
<#
.NOTES
    Created on:   10/5/2017
    Created by:   Andy Simmons
    Organization: St. Luke's Health System
    Filename:     Invoke-MailHandler.ps1
    Version:      1.2

    TODO:       - Add option in rule to move message to a custom folder 
                  after processing. (DONE)
                - Add support for 'dynamic' logging, to parse out interesting
                  content from the message data to identify a specific entity, 
                  and inject that into the log message, to enable event
                  correlation at the device level. (DONE)
                - Web interface for rule management

.SYNOPSIS
    Searches an Exchange mailbox for specific messages and generates
    corresponding Windows events, to enable integration with SCOM.

.DESCRIPTION
    Written to automate some NOC workflows that currently rely on humans
    monitoring a shared mailbox.

.PARAMETER Mailbox
    Email address of the mailbox to be processed.

.PARAMETER FolderName
    Mailbox folder containing messages to be processed.

.PARAMETER ItemLimit
    Maximum number of messages to retrieve for processing.

.PARAMETER SqlServer
    Hostname/FQDN/IP address of the SQL server containing the rules database.

.PARAMETER Database
    Name of the Database instance containing the rules.

.PARAMETER RuleTable
    Name of the table containing the rules.

.PARAMETER LogOnly
    Generate event logs, but do not modify the original email item's subject
    to indicate the message has been processed.
#>
using namespace Microsoft.Exchange.WebServices.Data
using namespace System.Net.Mail
using namespace System.Data.SqlClient
using namespace System.Text.RegularExpressions

[CmdletBinding(SupportsShouldProcess)]
param
(
    [Parameter(Mandatory)]
    [MailAddress]
    $Mailbox,

    [Parameter(Mandatory)]
    [string]
    $FolderName,
    
    [int]
    $ItemLimit = 10,
    
    [string]
    $SqlServer = 'IHTDBA01',
    
    [string]
    $Database = 'NOCRules',

    [string]
    $RuleTable = 'dbo.MailReaderRules',

    [switch]
    $LogOnly
)

#region Classes
class MailHandlerRule
{
    # Properties
    [int]         $Id
    [string]      $Name
    [MailAddress] $Mailbox
    [string]      $FolderName
    [string]      $DestinationFolderName
    [regex]       $SenderPattern
    [regex]       $SubjectPattern
    [regex]       $BodyPattern
    [string]      $LogServer
    [string]      $LogName
    [string]      $LogSource
    [int]         $EventId
    [string]      $EntryType
    [string]      $LogMessage
    [bool]        $IsDynamic
    [string]      $DynamicSource
    [string]      $ProcessedText
    [bool]        $Enabled
    [int]         $Priority
    [bool]        $CaseSensitive
    [bool]        $SourceVerified
    
    # Constructor
    MailHandlerRule ([object] $InputObject)
    {
        $o = $InputObject

        $this.Id             = $o.Id
        $this.Name           = $o.Name
        $this.Mailbox        = $o.Mailbox
        $this.FolderName     = $o.FolderName
        $this.LogServer      = $o.LogServer
        $this.LogName        = $o.LogName
        $this.LogSource      = $o.LogSource
        $this.EventId        = $o.EventId
        $this.EntryType      = $o.EntryType
        $this.LogMessage     = $o.LogMessage
        $this.IsDynamic      = $o.IsDynamic
        $this.Enabled        = $o.Enabled
        $this.Priority       = $o.Priority
        $this.CaseSensitive  = $o.CaseSensitive
        $this.SourceVerified = $false

        # Sanitize some stuff, probably will do that in the rule management interface too.
        if ($o.DynamicSource) { $this.DynamicSource = $o.DynamicSource}
        elseif ($o.IsDynamic) { $this.DynamicSource = 'Body' }
        
        if ($o.ProcessedText) { $this.ProcessedText = $o.ProcessedText }
        else                  { $this.ProcessedText = '[PROCESSED]' }
        
        if ($o.DestinationFolderName -eq $o.FolderName) { $this.DestinationFolderName = '' }
        else                                            { $this.DestinationFolderName = $o.DestinationFolderName }

        # The -imatch/-cmatch operators are misleading, and won't override case options from 
        # the regex to the right of the operator. We'll work around this by setting options
        # explicitly on each [regex] property, and just use -match for clarity.

        # always ignore case on SenderPattern
        $this.SenderPattern = [regex]::new($o.SenderPattern, [RegexOptions]::IgnoreCase)

        # use rule preference for matching subject and body content
        if ($this.CaseSensitive) { $regexOptions = [RegexOptions]::None }
        else                     { $regexOptions = [RegexOptions]::IgnoreCase }

        $this.SubjectPattern = [regex]::new($o.SubjectPattern, $regexOptions)
        $this.BodyPattern    = [regex]::new($o.BodyPattern, $regexOptions)
    }

    # Register the log source specified by this rule (to allow writing to the event log)
    [void] RegisterLogSource ()
    {
        try 
        { 
            $nelParams = @{
                LogName      = $this.LogName
                Source       = $this.LogSource
                ComputerName = $this.LogServer
                ErrorAction  = 'Stop'
            }
            New-Eventlog @nelParams 
            Write-Verbose "Registered $($this.LogName) log source '$($this.LogSource)' on $($this.LogServer)."
            $this.SourceVerified = $true
        }

        # If the log source already exists, suppress the error and shoot the message to verbose.
        catch [InvalidOperationException]
        { 
            $this.SourceVerified = $true
            Write-Verbose $_.Exception.Message 
        } 
        
        catch { throw $_.Exception }
    }
}

class MailHandler
{
    # Properties
    [MailAddress]       $Mailbox
    [string]            $FolderName
    [MailAddress]       $Sender
    [string]            $Subject
    [string]            $Body
    [bool]              $IsLogged
    [MailHandlerRule[]] $RuleSet
    [MailHandlerRule]   $Rule

    # Constructor
    MailHandler (
        [MailAddress]       $Mailbox, 
        [string]            $FolderName, 
        [MailAddress]       $Sender, 
        [string]            $Subject, 
        [string]            $Body, 
        [MailHandlerRule[]] $RuleSet
    ) 
    { 
        $this.Mailbox    = $Mailbox
        $this.FolderName = $FolderName
        $this.Sender     = $Sender
        $this.Subject    = $Subject
        $this.Body       = $Body
        
        $this.Initialize($RuleSet)
    }

    # Methods
    hidden [void] Initialize ([MailHandlerRule[]] $RuleSet) 
    {
        $this.IsLogged = $false

        if ($this.Validate())
        {
            # Load all enabled rules for this mailbox/folder
            $this.RuleSet = $RuleSet | Where-Object { 
                $_.Enabled -and 
                $_.Mailbox -eq $this.Mailbox -and
                $_.FolderName -eq $this.FolderName
            } | Sort-Object -Property 'Priority'

            # Select and assign a rule to this message
            if ($this.RuleSet) { $this.Rule = $this.FindMatch() }
            else               { Write-Warning 'No rules apply to the specified mailbox/folder.' }
        }
        else
        {
            Write-Warning "Not enough message content to process rules for item: '$($this.Subject)'"
        }
    }

    # Returns true if there's enough data to process rules
    hidden [bool] Validate ()
    {
        return $this.Mailbox -and $this.FolderName -and $this.Sender -and $this.Subject -and $this.Body
    }

    # Return the best match from the associated rule set, if any
    [MailHandlerRule] FindMatch ()
    {
        foreach ($rule in $this.RuleSet)
        {
            if ($this.TestMatch($rule)) { return $rule }
        }     

        return $null
    }

    # Compares the message against a single rule, returning true if it matches and the
    # message subject hasn't been tattooed as processed.
    hidden [bool] TestMatch ([MailHandlerRule] $Rule)
    {
        $processedPattern = [regex]::new('^{0}' -f [regex]::Escape($Rule.ProcessedText))

        return (
            ($this.Subject -notmatch $processedPattern) -and
            ($this.Mailbox       -eq $Rule.Mailbox) -and
            ($this.FolderName    -eq $Rule.FolderName) -and
            ($this.Sender     -match $Rule.SenderPattern) -and
            ($this.Subject    -match $Rule.SubjectPattern) -and
            ($this.Body       -match $Rule.BodyPattern) 
        )
    }
    
    # Invokes the selected rule on this message (create Windows event)
    [void] InvokeRule ()
    {
        try
        {
            $weParams = @{
                ComputerName = $this.Rule.LogServer
                LogName      = $this.Rule.LogName
                Source       = $this.Rule.LogSource
                EventId      = $this.Rule.EventId
                EntryType    = $this.Rule.EntryType
                Message      = $this.GetLogMessage()
                ErrorAction  = 'Stop'
            }
            
            Write-EventLog @weParams
            $this.IsLogged = $true
        }
        catch { throw $_.Exception }
    }

    # Compares the rule and email message, determines the appropriate event log 
    # message string, and returns it.
    [string] GetLogMessage () 
    {
        if ($this.Rule.IsDynamic) 
        {
            # Generate a dynamic log message, treating the rule's "LogMessage" property as a regex, 
            # and returning the first match from the specified source (body/subject/sender)
            $sourceContent = switch ($this.Rule.DynamicSource) 
            {
                'Body'    { $this.Body }
                'Subject' { $this.Subject }
                'Sender'  { $this.Sender }
                default
                { 
                    Write-Warning "Invalid dynamic message source '$($this.Rule.DyanamicSource)' specified. Defaulting to message body."
                    $this.Body 
                } 
            }

            # We'll use case options from the rule preference
            if ($this.Rule.CaseSensitive) { $regexOptions = [RegexOptions]::None }
            else                          { $regexOptions = [RegexOptions]::IgnoreCase }
           
            $message = [regex]::Match($sourceContent, $this.Rule.LogMessage, $regexOptions)
        }
        else 
        {
            # If the rule specifies a literal event log message, we'll use that.
            # If not, we'll use content from the mail message.
            if ($this.Rule.LogMessage) { $message = $this.Rule.LogMessage }
            else                       { $message = "$($this.Subject)`n`nMessage Body:`n$($this.Body)" }
        }
        
        # Absolute max message length is probably a little bigger, but Windows throws
        # vague error messages if you get just under the documented max of 32 KB - 2 bytes.
        $maxLength = 31KB
        if ($message.Length -gt $maxLength) { $message = $message.SubString(0, $maxLength) }

        return $message
    }
}
#endregion Classes


#region Functions
<#
.SYNOPSIS
    Retrieves message processing rules from a database.

.PARAMETER SqlServer
    Hostname/FQDN/IP address of the SQL server containing the rules database.

.PARAMETER Database
    Name of the Database instance containing the rules.

.PARAMETER RuleTable
    Name of the table containing the rules.

.PARAMETER Mailbox
    Email address of the mailbox to be processed.

.PARAMETER FolderName
    Mailbox folder containing messages to be processed.

.PARAMETER IncludeDisabled
    Includes disabled rules in the result.
#>
function Get-MailHandlerRules
{
    [CmdletBinding()]
    param (
        [string] $SqlServer  = $SqlServer,
        [string] $Database   = $Database,
        [string] $RuleTable  = $RuleTable,
        [string] $Mailbox    = $Mailbox,
        [string] $FolderName = $FolderName,
        [switch] $IncludeDisabled
    )

    Write-Verbose "Opening connection to $Database database on $SqlServer"
    $connection = [SqlConnection]::new()
    $connection.ConnectionString = "Data Source=$SqlServer;Integrated Security=True;Initial Catalog=$Database"
    $connection.Open()

    # build base query
    $sql = "SELECT 
                Id,
                RuleName,
                Enabled,
                Priority,
                Mailbox,
                FolderName,
                DestinationFolderName,
                SenderPattern,
                SubjectPattern,
                BodyPattern,
                LogServer,
                LogName,
                LogSource,
                EventId,
                EntryType,
                LogMessage,
                IsDynamic,
                DynamicSource,
                ProcessedText,
                CaseSensitive
            FROM
                $RuleTable"
    
    # apply some optional filters based on invocation parameters
    $filter = @(
        if ($Mailbox)          { "Mailbox like '$Mailbox'" }
        if ($FolderName)       { "FolderName like '$FolderName'"}
        if (!$IncludeDisabled) { 'Enabled = 1' }
    )
    if ($filter) { $sql += "`nWHERE " + ($filter -join " AND`n") }
    
    # finalize and execute query
    $sql += "`nORDER BY Priority ASC"
    Write-Verbose $sql

    $cmd = [SqlCommand]::new($sql, $connection)
    $reader = $cmd.ExecuteReader()

    # create a [MailHandlerRule] from each record
    while ($reader.Read())
    {
        $ruleData = [pscustomobject] @{
            Id                    = [int]    $reader.Item('Id')
            Name                  = [string] $reader.Item('RuleName')
            Enabled               = [bool]   $reader.Item('Enabled')
            Priority              = [int]    $reader.Item('Priority')
            Mailbox               = [string] $reader.Item('Mailbox')
            FolderName            = [string] $reader.Item('FolderName')
            DestinationFolderName = [string] $reader.Item('DestinationFolderName')
            SenderPattern         = [regex]  $reader.Item('SenderPattern')
            SubjectPattern        = [regex]  $reader.Item('SubjectPattern')
            BodyPattern           = [regex]  $reader.Item('BodyPattern')
            LogServer             = [string] $reader.Item('LogServer')
            LogName               = [string] $reader.Item('LogName')
            LogSource             = [string] $reader.Item('LogSource')
            EventId               = [int]    $reader.Item('EventId')
            EntryType             = [string] $reader.Item('EntryType')
            LogMessage            = [string] $reader.Item('LogMessage')
            IsDynamic             = [bool]   $reader.Item('IsDynamic')
            DynamicSource         = [string] $reader.Item('DynamicSource')
            ProcessedText         = [string] $reader.Item('ProcessedText')
            CaseSensitive         = [bool]   $reader.Item('CaseSensitive')
        }

        [MailHandlerRule]::new($ruleData)
    }

    Write-Verbose 'Closing SQL connection'
    $reader.Close()
    $connection.Close()
}

<#
.SYNOPSIS
    Registers the log sources for one or more mail handler rules.

.DESCRIPTION
    This accomplishes the same end-result as calling RegisterLogSource() on a 
    collection of [MailHanderRule] objects, but it's more efficient.
    
    It only attempts one registration for each unique log source in the 
    collection, rather than once for every rule.

.PARAMETER RuleSet
    Rules specifying the log sources to be registered. 
#>
function Register-MailHandlerLogSource
{
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]	
        [MailHandlerRule[]] 
        $RuleSet
    )
    # Group the rule set by explicit log source (server + log + source)
    $ruleGroups = @($ruleSet | Group-Object -Property LogServer, LogName, LogSource)
    
    foreach ($ruleGroup in $ruleGroups) 
    {
        $logServer, $logName, $logSource = $ruleGroup.Name -split ', '
        if ($PSCmdlet.ShouldProcess("[$logServer]: $logName\$logSource", 'register log source'))
        {
            # Call the registration method on the first rule in each group, and if successful
            # mark the whole group as verified.
            $ruleGroup.Group[0].RegisterLogSource()
        }
    }
}

<#
.SYNOPSIS
    Reads a mailbox folder and processes mail handler rules.

.DESCRIPTION
    Checks a mailbox for messages that match specific sender/subject/body
    patterns as defined by a set of mail handler rules, and invokes those
    rules as appropriate.

    Requires the Exchange Web Services (EWS) API. Tested with v2.2.

.PARAMETER Mailbox
    Mail address associated with the mailbox (used for autodiscovery of Exchange server)

.PARAMETER FolderName
    Mailbox folder containing messages to be processed.

.PARAMETER RuleSet
    Collection of rules for message processing

.PARAMETER EwsAssembly
    Exchange Web Services API DLL path

.PARAMETER Credentials
    Exchange mailbox credentials
    
.PARAMETER ItemLimit
    Maximum number of items to retrieve

.PARAMETER LogOnly
    Generate event logs, but do not modify the original email item to
    indicate successful processing.
#>
function Invoke-MailHandler
{
    [CmdletBinding(SupportsShouldProcess)]
    param 
    (
        [Parameter(Mandatory)]
        [MailAddress]
        $Mailbox,
        
        [Parameter(Mandatory)]
        [string]
        $FolderName,
       
        [Parameter(Mandatory)]
        [MailHandlerRule[]]
        $RuleSet,
        
        [ValidateScript({Test-Path $_})]
        [string]
        $EwsAssembly = "${env:ProgramFiles}\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll",
        
        [Net.NetworkCredential]
        $Credentials = [Net.CredentialCache]::DefaultNetworkCredentials,

        [int]
        $ItemLimit = $ItemLimit,
        
        [switch]
        $LogOnly = $LogOnly
    )

    [void] [Reflection.Assembly]::LoadFile($EwsAssembly)
    
    # Configure the Exchange service via autodiscovery
    $exchangeService = [ExchangeService]::new()
    $exchangeService.Credentials = $Credentials
    $exchangeService.AutoDiscoverUrl($Mailbox)

    # Build a custom property set to access the message body as plain text
    $desiredProps = [PropertySet]::new()
    $desiredProps.BasePropertySet = [BasePropertySet]::FirstClassProperties
    $desiredProps.RequestedBodyType = [BodyType]::Text

    Write-Verbose "Retrieving the $ItemLimit most recent '$FolderName' items for $Mailbox."
    $mailboxFolderId = [FolderId]::new($FolderName, [Mailbox]::new($Mailbox))
    $itemView = [ItemView]::new($ItemLimit)
    $searchResult = $exchangeService.FindItems($mailboxFolderId, $itemView)

    # Process oldest messages first
    $items = @($searchResult.Items)
    [array]::Reverse($items)

    # If the ruleset contains any destination folder names, perform a single lookup to find them all,
    # and stuff those into a collection we can reference during message processing.
    $destFolderNames = $RuleSet.Where({$_.DestinationFolderName}).DestinationFolderName | Sort-Object -Unique
    if ($destFolderNames)
    {
        Write-Verbose "Searching '$FolderName' for subfolders referenced by the current ruleset..."
        $containment = [ContainmentMode]::FullString
        $comparison  = [ComparisonMode]::IgnoreCase
        $folderProp  = [FolderSchema]::DisplayName

        $filters = $destFolderNames.ForEach({[SearchFilter+ContainsSubstring]::new($folderProp, $_, $containment, $comparison)})
        $compoundFilter = [SearchFilter+SearchFilterCollection]::new([LogicalOperator]::Or, $filters)

        $folderView = [FolderView]::new($destFolderNames.Count)

        $destFolders = $exchangeService.FindFolders($mailboxFolderId, $compoundFilter, $folderView)
    }
    
    Write-Verbose "Processing $($RuleSet.Count) rules on $($items.Count) messages ..."
    foreach ($item in $items)
    {
        $item.Load($desiredProps)

        # create a mail handler for this message
        $mailHandler = [MailHandler]::new($Mailbox, $FolderName, $item.Sender.Address, $item.Subject, $item.Body.Text, $RuleSet)
        
        # If the mail handler assigned a rule to this message, invoke it
        if ($mailHandler.Rule -and $PSCmdlet.ShouldProcess($mailHandler.Subject, "Invoke rule: $($mailHandler.Rule.Name)"))
        {
            # Log message
            try { $mailHandler.InvokeRule() }
            catch 
            {
                Write-Warning $_.Exception.Message
                continue
            }
            
            # Tattoo the email's subject text as processed
            if ($mailHandler.IsLogged -and !$LogOnly) 
            {
                Write-Verbose "Marking item '$($item.Subject)' as processed"
                $item.Subject = "$($mailHandler.Rule.ProcessedText) $($item.Subject)"
                $item.IsRead = $true
                $item.Update([ConflictResolutionMode]::AutoResolve)

                # If the rule specifies a destination folder, move it
                if ($mailHandler.Rule.DestinationFolderName -and $destFolders)
                {
                    $destFolder = ($destFolders.Where({$_.DisplayName -eq $mailHandler.Rule.DestinationFolderName}))[0]
                    
                    if ($destFolder) 
                    {
                        Write-Verbose "Moving item from '$FolderName' to '$($destFolder.DisplayName)'" 
                        [void] $item.Move($destFolder.Id)
                    }
                    else 
                    {
                        Write-Warning "Couldn't find '$FolderName' subfolder: '$($mailHandler.Rule.DestinationFolderName)'!"
                    }                    
                }
            }
        }        
    }
}
#endregion Functions


#region Main

# Retrieve rule set for this mailbox/folder
$ruleSet = @(Get-MailHandlerRules)

# Process rules (if any)
if ($ruleSet)
{
    try { Register-MailHandlerLogSource -RuleSet $ruleSet -ErrorAction Stop }
    catch
    { 
        Write-Error 'Failed to register the log source for one or more rules. Aborting.'
        throw $_.Exception
    }

    $imhParams = @{
        Mailbox    = $Mailbox
        FolderName = $FolderName
        RuleSet    = $ruleSet
        ItemLimit  = $ItemLimit
        LogOnly    = $LogOnly
    }
    Invoke-MailHandler @imhParams 
}
#endregion Main
