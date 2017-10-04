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
[CmdletBinding(SupportsShouldProcess)]
param
(
	[string] $SqlServer = 'IHTDBA01',
    
	[string] $Database = 'NOCRules',

	[Parameter(Mandatory)]
	[string]
	$ProcessedText
)

#region Classes
class MailReader
{
	# Properties
	[Net.Mail.MailAddress] $Sender
	[string] $Subject
	[string] $Body
	[bool] $MatchesRule
	[bool] $IsLogged
	[MailReaderRule[]] $RuleSet

	# Constructor
	MailReader([Net.Mail.MailAddress] $Sender, [string] $Subject, [string] $Body, [MailReaderRule[]] $RuleSet) 
	{ 
		$this.Sender = $Sender
		$this.Subject = $Subject
		$this.Body = $Body
		
		$this.RuleSet = $RuleSet.Where({$_.Enabled})
		$this.MatchesRule = $false
		$this.IsLogged = $false
	}
    
	# Methods
	[bool] Validate ()
	{
		# Ensure sender, subject, and body aren't empty, and at least one enabled rule exists
		return (
			($this.Sender -and $this.Subject -and $this.Body) -and 
			$this.RuleSet.Where({$_.Enabled})
		)
	}

	hidden [bool] TryMatch ([MailReaderRule] $Rule)
	{
		return (
			($this.Subject -notmatch $rule.ProcessedText) -and
			($this.Sender -match $rule.SenderPattern) -and
			($this.Subject -match $rule.SubjectPattern) -and
			($this.Body -match $rule.BodyPattern)
		)
	}

	[void] Log ([MailReaderRule] $Rule)
	{
		# If the rule specifies an event log message, we'll use that.
		# If not, we'll use the subject from the mail message.
		if ($Rule.LogMessage) { $Message = $Rule.LogMessage }
		else                  { $Message = $this.Subject }

		try
		{
			$weParams = @{
				ComputerName = $Rule.LogServer
				LogName      = $Rule.LogName
				Source       = $Rule.LogSource
				EventId      = $Rule.EventId
				EntryType    = $Rule.EntryType
				Message      = $Message
				ErrorAction  = 'Stop'
			}
            
			Write-EventLog @weParams
			$this.IsLogged = $true
		}
		catch 
		{
			Write-Error $_.Exception.Message
		}
	}
}

class MailReaderRule
{
	# Properties
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
	
	# Constructors
	MailReaderRule ([pscustomobject] $InputObject)
	{
		$this.SenderPattern = $InputObject.SenderPattern
		$this.SubjectPattern = $InputObject.SubjectPattern
		$this.BodyPattern = $InputObject.BodyPattern
		$this.LogServer = $InputObject.LogServer
		$this.LogName = $InputObject.LogName
		$this.LogSource = $InputObject.LogSource
		$this.EventId = $InputObject.EventId
		$this.EntryType = $InputObject.EntryType
		$this.LogMessage = $InputObject.LogMessage
		$this.ProcessedText = $InputObject.ProcessedText
		$this.Enabled = $InputObject.Enabled	
	}

	MailReaderRule (
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
		[bool]   $Enabled
	)
	{
		$this.SenderPattern = $SenderPattern
		$this.SubjectPattern = $SubjectPattern
		$this.BodyPattern = $BodyPattern
		$this.LogServer = $LogServer
		$this.LogName = $LogName
		$this.LogSource = $LogSource
		$this.EventId = $EventId
		$this.EntryType = $EntryType
		$this.LogMessage = $LogMessage
		$this.ProcessedText = $ProcessedText
		$this.Enabled = $Enabled
	}

	# No methods
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
		[string] $Database = 'NOCRules'
	)

	Write-Verbose "Opening connection to $Database database on $SQLServer"
	$connection = New-Object Data.SqlClient.SqlConnection
	$connection.ConnectionString = "Data Source=$SqlServer;" +
	"Integrated Security=True;" +
	"Initial Catalog=$Database"
	$connection.Open()

	$query = "select 
                Id,
                RuleName,
                Enabled,
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
                LogMessage
            from
                dbo.MailReaderRules"

	Write-Verbose 'Loading rules'
	$cmd = New-Object Data.SqlClient.SqlCommand $query, $connection
	$reader = $cmd.ExecuteReader()

	while ($reader.Read())
	{
		[pscustomobject] [ordered] @{
			Id             = [int]    $reader.Item('Id')
			RuleName       = [string] $reader.Item('RuleName')
			Enabled        = [bool]   $reader.Item('Enabled')
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
		}
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

.PARAMETER SenderPattern
    Regular expression describing eligible "From" email address(es)

.PARAMETER SubjectPattern
    Regular expression describing eligible "Subject" text

.PARAMETER BodyPattern
    Regular expression describing eligible "Body" text

.EXAMPLE
    Receive-TriggerMessage -Mailbox 'joe@abc.tld'
#>
function Export-TriggerMessage 
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

		[Parameter(Mandatory)]
		[regex]
		$SenderPattern,

		[Parameter(Mandatory)]
		[regex]
		$SubjectPattern,

		[Parameter(Mandatory)]
		[regex]
		$BodyPattern
	)

	[void] [Reflection.Assembly]::LoadFile($EwsAssembly)
    
	# Configure the Exchange service via autodiscovery
	$exchangeService = New-Object -TypeName 'Microsoft.Exchange.WebServices.Data.ExchangeService'
	$exchangeService.Credentials = $Credentials
	$exchangeService.AutoDiscoverUrl($Mailbox)

	$mailFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $FolderName)
    
	# Build a custom property set to access the message body as plain text
	$desiredProps = New-Object -TypeName 'Microsoft.Exchange.WebServices.Data.PropertySet'
	$desiredProps.BasePropertySet = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties
	$desiredProps.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text

	# Auto-resolve message update conflicts
	$resolveConflicts = [Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve

	# Set the deletion mode to move the message to deleted items
	$deleteMode = [Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems

	Write-Verbose "Retrieving the $ItemLimit most recent messages for $Mailbox ..."
	$items = $mailFolder.FindItems($ItemLimit)
    
	foreach ($item in $items)
	{
		$item.Load($desiredProps)

		$itemIsTrigger = ($item.Subject -notmatch "^$ProcessedText") -and 
		($item.Subject -match $SubjectPattern)      -and
		($item.Sender -match $SenderPattern)        -and
		($item.Body -match $BodyPattern)
        
		if ($itemIsTrigger)
		{
			if ($PSCmdlet.ShouldProcess($item.Subject, 'process'))
			{
				Write-Verbose "TRIGGER FOUND! $($item.Subject)"
				# Return a custom object with some interesting message properties for easy export
				[pscustomobject] [ordered] @{
					DateTimeSent     = $item.DateTimeSent
					DateTimeReceived = $item.DateTimeReceived
					SenderAddress    = $item.Sender.Address
					Subject          = $item.Subject
					Body             = $item.Body.Text
					InternetHeaders  = $item.InternetMessageHeaders -join "`n"
				}

				# Prepend the "processed" text to the subject, mark message as read, and delete it.
				$item.Subject = "$ProcessedText$($item.Subject)"
				$item.IsRead = $true
				$item.Update($resolveConflicts)
				$item.Delete($deleteMode)
			}
		}
		elseif ($item.Subject -match "^$ProcessedText")
		{
			Write-Warning "Item may have been partially processed on a previous run. Skipping: $($item.Subject)"
		}
		else
		{
			Write-Verbose "Skipping item: $($item.Subject)"
		}
	}
}
#endregion Functions

# Retrieve new trigger messages
