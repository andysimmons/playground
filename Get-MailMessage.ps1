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
    Requires the Exchange Web Services (EWS) API. Tested with v2.2.

.PARAMETER EwsAssembly
    Path to the EWS DLL

.PARAMETER Credentials
    Credentials used to connect to EWS

.PARAMETER Mailbox
    Mail address associated with the mailbox (used for autodiscovery of Exchange server)

.EXAMPLE
    Get-MailMessage -Mailbox 'joeschmoe@abc.tld'
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
        $ItemLimit = 5
    )

    # Load the EWS assembly and configure the service
    [void] [Reflection.Assembly]::LoadFile($EwsAssembly)
    
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
    $autoResolve = [Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve

    Write-Verbose "Retrieving the $ItemLimit most recent messages for $Mailbox ..."
    $items = $inbox.FindItems($ItemLimit)
    
    foreach ($item in $items)
    {
        Write-Verbose "Toggling 'Read' status on item '$($item.Subject)'"

        $item.Load($desiredProps)
        $item.IsRead = !$item.IsRead
        $item.Update($autoResolve)
    }
}
