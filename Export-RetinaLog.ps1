#requires -Version 5
<#
.NOTES
    Created on:   10/2/2017
    Created by:   Andy Simmons
    Organization: St. Luke's Health System
    Filename:     Export-RetinaLog.ps1

.SYNOPSIS
    Checks an Exchange mailbox for a specific type of message, and if found,
    exports the message content as a Windows event.

.DESCRIPTION
    Written to automate some NOC workflows that currently rely on humans
    monitoring a shared mailbox.
#>
using namespace Microsoft.Exchange.WebServices.Data

[CmdletBinding(SupportsShouldProcess)]
param
(
	[Parameter(Mandatory)]
	[Net.Mail.MailAddress]
	$Mailbox,

	[Parameter(Mandatory)]
	[string]
	$FolderName,
    
	[int]
	$ItemLimit = 10,
    
	[string]
	$SqlServer = 'IHTDBA01',
    
	[string]
	$Database = 'NOCRules'
)

#region Classes
class MailReader
{
	# Properties
	[Net.Mail.MailAddress] $Mailbox
	[string] $FolderName
	[Net.Mail.MailAddress] $Sender
	[string] $Subject
	[string] $Body
	[bool] $IsLogged
	[MailReaderRule[]] $RuleSet
	[MailReaderRule] $Rule

	# Constructor
	MailReader (
		[Net.Mail.MailAddress] $Mailbox, 
		[string] $FolderName, 
		[Net.Mail.MailAddress] $Sender, 
		[string] $Subject, 
		[string] $Body, 
		[MailReaderRule[]] $RuleSet
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
	[void] Initialize ([MailReaderRule[]] $RuleSet) 
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
		}
	}

	[bool] Validate ()
	{
		return $this.Mailbox -and $this.FolderName -and $this.Sender -and $this.Subject -and $this.Body
	}

	[MailReaderRule] FindMatch ()
	{
		foreach ($rule in $this.RuleSet)
		{
			if ($this.TestMatch($rule)) { return $rule }
		}     

		return $null
	}

	hidden [bool] TestMatch ([MailReaderRule] $Rule)
	{
		# Returns true if this rule matches the message, and the message
		# hasn't been marked as processed already
        $processedPattern = "^$([regex]::Escape($rule.ProcessedText))"

		return ( 
			($this.Subject -notmatch $processedPattern) -and
			($this.Mailbox       -eq $rule.Mailbox) -and
			($this.FolderName    -eq $rule.FolderName) -and
			($this.Sender     -match $rule.SenderPattern) -and
			($this.Subject    -match $rule.SubjectPattern) -and
			($this.Body       -match $rule.BodyPattern)
		)
	}
    
	[void] InvokeRule ()
	{
		# If the rule specifies an event log message, we'll use that.
		# If not, we'll use the subject from the mail message.
        if ($this.Rule.LogMessage) 
        { 
            $message = $this.Rule.LogMessage 
        }
        else 
        { 
            $message = $this.Subject + "`n`nMessage Body: `n" + $this.Body
        }

        # Absolute max message length is probably a little bigger, but Windows throws
        # vague error messages if you get just under the documented max of 32 KB - 2 bytes.
        $maxLength = 31KB
        if ($message.Length -gt $maxLength) 
        {
            $message = $message.SubString(0, $maxLength)
        }

		try
		{
			$weParams = @{
				ComputerName = $this.Rule.LogServer
				LogName      = $this.Rule.LogName
				Source       = $this.Rule.LogSource
				EventId      = $this.Rule.EventId
				EntryType    = $this.Rule.EntryType
				Message      = $message
				ErrorAction  = 'Stop'
			}
            
			Write-EventLog @weParams
			$this.IsLogged = $true
		}
		catch 
		{
			throw $_.Exception
		}
	}
}

class MailReaderRule
{
	# Properties
	[int]    $Id
	[string] $Name
	[Net.Mail.MailAddress] $Mailbox
	[string] $FolderName
	[regex]  $SenderPattern
	[regex]  $SubjectPattern
	[regex]  $BodyPattern
	[string] $LogServer
	[string] $LogName
	[string] $LogSource
	[int]    $EventId
	[string] $EntryType
	[string] $LogMessage
	[string] $ProcessedText
	[bool]   $Enabled
	[int]    $Priority
    
	# Constructors
	MailReaderRule ([object] $InputObject)
	{
        $o = $InputObject
        $this.Initialize($o.Id, $o.Name, $o.Mailbox, $o.FolderName, $o.SenderPattern, $o.SubjectPattern, 
            $o.BodyPattern, $o.LogServer, $o.LogName, $o.LogSource, $o.EventId, $o.EntryType, $o.LogMessage, 
            $o.ProcessedText, $o.Enabled, $o.Priority)
	}

	MailReaderRule (
		[int]    $Id,
		[string] $Name,
		[Net.Mail.MailAddress] $Mailbox,
		[string] $FolderName,
		[regex]  $SenderPattern,
		[regex]  $SubjectPattern,
		[regex]  $BodyPattern,
		[string] $LogServer,
		[string] $LogName,
		[string] $LogSource,
		[int]    $EventId,
		[string] $EntryType,
		[string] $LogMessage,
		[string] $ProcessedText,
		[bool]   $Enabled,
		[int]    $Priority
	)
	{
		$this.Initialize($Id, $Name, $Mailbox, $FolderName, $SenderPattern, $SubjectPattern, $BodyPattern, $LogServer, 
			$LogName, $LogSource, $EventId, $EntryType, $LogMessage, $ProcessedText, $Enabled, $Priority)
	}

	# Methods
	[void] Initialize (
		[int]    $Id,
		[string] $Name,
		[Net.Mail.MailAddress] $Mailbox,
		[string] $FolderName,
		[regex]  $SenderPattern,
		[regex]  $SubjectPattern,
		[regex]  $BodyPattern,
		[string] $LogServer,
		[string] $LogName,
		[string] $LogSource,
		[int]    $EventId,
		[string] $EntryType,
		[string] $LogMessage,
		[string] $ProcessedText,
		[bool]   $Enabled,
		[int]    $Priority
	)
	{
		$this.Id             = $Id
		$this.Name           = $Name
		$this.Mailbox        = $Mailbox
		$this.FolderName     = $FolderName
		$this.SenderPattern  = $SenderPattern
		$this.SubjectPattern = $SubjectPattern
		$this.BodyPattern    = $BodyPattern
		$this.LogServer      = $LogServer
		$this.LogName        = $LogName
		$this.LogSource      = $LogSource
		$this.EventId        = $EventId
		$this.EntryType      = $EntryType
		$this.LogMessage     = $LogMessage
		$this.Enabled        = $Enabled
		$this.Priority       = $Priority

		if ($ProcessedText) { $this.ProcessedText = $ProcessedText }
		else                { $this.ProcessedText = '[PROCESSED] ' }
	}
    
	[void] RegisterLogSource ()
	{
		# Register the log source for this rule (to allow writing to the event log)
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
		}

		# Redirect error to the verbose stream if the log source already exists
		catch [InvalidOperationException] { Write-Verbose $_.Exception.Message } 
        
		catch { throw $_.Exception }
	}

	[string] ToString () { return $this.Name }
}
#endregion Classes

#region Functions
<#
.SYNOPSIS
    Retrieves message processing rules from a database.

.PARAMETER SqlServer
    SQL Server hostname/IP/FQDN

.PARAMETER Database
    Database name
#>
function Get-MailReaderRules
{
	[CmdletBinding()]
	param (
		[string] $SqlServer = 'IHTDBA01',
		[string] $Database = 'NOCRules',
		[string] $Mailbox = $Mailbox,
		[string] $FolderName = $FolderName,
		[switch] $IncludeDisabled
    )

	Write-Verbose "Opening connection to $Database database on $SQLServer"
	$connection = New-Object Data.SqlClient.SqlConnection
	$connection.ConnectionString = "Data Source=$SqlServer;" +
	"Integrated Security=True;" +
	"Initial Catalog=$Database"
	$connection.Open()

	$query = 'SELECT 
                Id,
                RuleName,
                Enabled,
                Priority,
                Mailbox,
                FolderName,
                SenderPattern,
                SubjectPattern,
                BodyPattern,
                LogServer,
                LogName,
                LogSource,
                EventId,
                EntryType,
                LogMessage,
                ProcessedText
            FROM
                dbo.MailReaderRules'
    
	$ruleFilter = @()
	if ($Mailbox)          { $ruleFilter += "Mailbox like '$Mailbox'" }
	if ($FolderName)       { $ruleFilter += "FolderName like '$FolderName'"}
	if (!$IncludeDisabled) { $ruleFilter += "Enabled = 1" }
    
	if ($ruleFilter)
	{
		$query += "`nWHERE " + $($ruleFilter -join " AND`n")
	}
    
	$query += "`nORDER BY Priority ASC"

	write-verbose $query

	Write-Verbose 'Loading rules'
	$cmd = New-Object Data.SqlClient.SqlCommand $query, $connection
	$reader = $cmd.ExecuteReader()

	while ($reader.Read())
	{
		$ruleData = [pscustomobject] @{
			Id             = [int]    $reader.Item('Id')
			Name           = [string] $reader.Item('RuleName')
			Enabled        = [bool]   $reader.Item('Enabled')
			Priority       = [int]    $reader.Item('Priority')
			Mailbox        = [string] $reader.Item('Mailbox')
			FolderName     = [string] $reader.Item('FolderName')
			SenderPattern  = [regex]  $reader.Item('SenderPattern')
			SubjectPattern = [regex]  $reader.Item('SubjectPattern')
			BodyPattern    = [regex]  $reader.Item('BodyPattern')
			LogServer      = [string] $reader.Item('LogServer')
			LogName        = [string] $reader.Item('LogName')
			LogSource      = [string] $reader.Item('LogSource')
			EventId        = [int]    $reader.Item('EventId')
			EntryType      = [string] $reader.Item('EntryType')
			LogMessage     = [string] $reader.Item('LogMessage')
			ProcessedText  = [string] $reader.Item('ProcessedText')
		}

		[MailReaderRule]::New($ruleData)
	}

	Write-Verbose 'Closing SQL connection'
	$reader.Close()
	$connection.Close()
}

<#
.SYNOPSIS
Processes "trigger" email messages.

.DESCRIPTION
Checks a mailbox for messages that match specific sender and subject
patterns, returns a few interesting properties, and deletes those messages.

Requires the Exchange Web Services (EWS) API. Tested with v2.2.

.PARAMETER EwsAssembly
Exchange Web Services API DLL path

.PARAMETER Credentials
Exchange mailbox credentials

.PARAMETER Mailbox
Mail address associated with the mailbox (used for autodiscovery of Exchange server)

.PARAMETER ItemLimit
Maximum number of items to retrieve

.PARAMETER RuleSet
Collection of rules for message processing
#>
function Invoke-MailReader
{
	[CmdletBinding(SupportsShouldProcess)]
	param 
	(
		[ValidateScript({Test-Path $_})]
		[string]
		$EwsAssembly = "${env:ProgramFiles}\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll",
        
		[Net.NetworkCredential]
		$Credentials = [Net.CredentialCache]::DefaultNetworkCredentials,
        
		[Parameter(Mandatory)]
		[Net.Mail.MailAddress]
		$Mailbox,
        
		[int]
		$ItemLimit = 50,

		[Parameter(Mandatory)]
		[string]
		$FolderName,

		[MailReaderRule[]]
		$RuleSet
	)

	[void] [Reflection.Assembly]::LoadFile($EwsAssembly)
    
	# Configure the Exchange service via autodiscovery
	$exchangeService = New-Object -TypeName 'ExchangeService'
	$exchangeService.Credentials = $Credentials
	$exchangeService.AutoDiscoverUrl($Mailbox)

	$mailFolder = [Folder]::Bind($exchangeService, $FolderName)
    
	# Build a custom property set to access the message body as plain text
	$desiredProps = New-Object -TypeName 'PropertySet'
	$desiredProps.BasePropertySet = [BasePropertySet]::FirstClassProperties
	$desiredProps.RequestedBodyType = [BodyType]::Text

	# Auto-resolve message update conflicts
	$resolveConflicts = [ConflictResolutionMode]::AutoResolve

	# Set the deletion mode to move the message to deleted items
	#$deleteMode = [DeleteMode]::MoveToDeletedItems

	Write-Verbose "Retrieving the $ItemLimit most recent messages for $Mailbox ..."
	$items = $mailFolder.FindItems($ItemLimit)
    
	foreach ($item in $items)
	{
		$item.Load($desiredProps)

		$reader = [MailReader]::New($Mailbox, $FolderName, $item.Sender.Address, $item.Subject, $item.Body.Text, $RuleSet)
        
		if ($reader.Rule -and $PSCmdlet.ShouldProcess($reader.Subject, "process rule: $($reader.Rule)"))
		{ 
			$reader.InvokeRule()
            
			if ($reader.IsLogged) 
			{
				Write-Verbose "Marking item '$($item.Subject)' as processed"
				$item.Subject = $reader.Rule.ProcessedText + $item.Subject
				$item.IsRead = $true
				$item.Update($resolveConflicts)
            }
		}        
	}
}
#endregion Functions

$ruleSet = @(Get-MailReaderRules)

if ($ruleSet)
{
	try
	{ 
		# Group the rule set by explicit log source (server + log + source)
		$ruleGroups = @($ruleSet | Group-Object -Property LogServer, LogName, LogSource)
        
		foreach ($ruleGroup in $ruleGroups) 
		{
			if ($PSCmdlet.ShouldProcess($ruleGroup.Name, 'register event log source'))
			{
				# Call the registration method on the first rule in each group.
				$ruleGroup.Group[0].RegisterLogSource()
			}
		}
	}
	catch
	{ 
		Write-Error "Failed to register the log source for one or more rules. Aborting"
		throw $_.Exception
	}
	$reader = Invoke-MailReader -Mailbox $Mailbox -FolderName $FolderName -RuleSet $ruleSet -ItemLimit $ItemLimit
	$reader    
}
