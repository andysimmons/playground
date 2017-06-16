#requires -Version 4
<#
.NOTES
    Created on:   6/14/2016
    Created by:   Andy Simmons
    Organization: St. Luke's Health System
    Filename:     Get-MailMessage.ps1

.SYNOPSIS
    Tinkering around with a simple Exchange mail client to retrieve messages.

.DESCRIPTION
    Checks a mailbox for messages that match specific sender and subject
    patterns, returns the delivery timestamp(s), and deletes those messages.

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

.EXAMPLE
    Get-MailMessage -Mailbox 'joe@abc.tld'
#>
function Get-MailMessage
{
    [CmdletBinding()]
    param 
    (
        [ValidateScript({Test-Path $_})]
        [string]
        $EwsAssembly = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll',
        
        [Net.NetworkCredential]
        $Credentials = [Net.CredentialCache]::DefaultNetworkCredentials,
        
        [Parameter(Mandatory, HelpMessage = 'Primary email address of the Exchange mailbox')]
        [Net.Mail.MailAddress]
        $Mailbox,
        
        [int]
        $ItemLimit = 5,

        [regex]
        $SenderPattern = 'sender@some\.tld',

        [regex]
        $SubjectPattern = 'something interesting'
    )

    [void] [Reflection.Assembly]::LoadFile($EwsAssembly)
    
    # Configure the Exchange service via autodiscovery
    $exchangeService = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new()
    $exchangeService.Credentials = $Credentials
    $exchangeService.AutodiscoverUrl($Mailbox)

    $inboxPath = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox
    $inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $inboxPath)
    
    # Build a custom property set to access the message body as plain text
    $defaultProps = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties
    $desiredProps = [Microsoft.Exchange.WebServices.Data.PropertySet]::new($defaultProps)
    $desiredProps.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text

    # Auto-resolve message update conflicts
    $resolveConflicts = [Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve

    Write-Verbose "Retrieving the $ItemLimit most recent messages for $Mailbox ..."
    $items = $inbox.FindItems($ItemLimit)
    
    foreach ($item in $items)
    {
        Write-Verbose "Toggling 'Read' status on item '$($item.Subject)'"

        $item.Load($desiredProps)
        $item.IsRead = !$item.IsRead
        $item.Update($resolveConflicts)
    }
}
