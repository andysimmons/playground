#requires -Module TunableSSLValidator
#requires -Version 5

using namespace Microsoft.PowerShell.Commands
using namespace System.Net
using namespace System.Xml
using namespace System.Runtime.InteropServices
using namespace System.Collections.Generic

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [pscredential]
    $Credential,

    [int[]]
    $Id = 6223..6226,

    [uri]
    $BaseUri = 'https://rhapsodyindev01.slhs.org:8444',

    [switch]
    $Insecure
)

enum CommPointAction 
{
    START
    STOP
    RESTART
}

class CommPoint
{
    # properties
    hidden [pscredential] $Credential
    hidden [Uri]          $BaseUri
    
    [int]      $Id
    [string]   $Name
    [string]   $FolderPath
    [string]   $Mode
    [string]   $State
    [string]   $CommPointType
    [timespan] $InputIdleTime
    [int]      $OutQueueSize
    [timespan] $OutputIdleTime
    [timespan] $Uptime
    [int]      $ConnectionCount
    [int]      $SentCount
    [bool]     $SkipCertificateCheck

    # constructors
    CommPoint (
        [Uri]          $BaseUri,
        [int]          $Id,
        [pscredential] $Credential
    )
    {
        $this.BaseUri = $BaseUri
        $this.Id = $Id
        $this.Credential = $Credential
        $this.SkipCertificateCheck = $false

        $this.Initialize()
    }

    CommPoint (
        [Uri]          $BaseUri,
        [int]          $Id,
        [pscredential] $Credential,
        [bool]         $SkipCertificateCheck
    )
    {
        $this.BaseUri = $BaseUri
        $this.Id = $Id
        $this.Credential = $Credential
        $this.SkipCertificateCheck = $SkipCertificateCheck

        $this.Initialize()
    }

    # methods
    hidden [void] Initialize ()
    {
        try { $this.Refresh() }
        catch
        {
            Write-Error "Initialization failed."
            throw $_.Exception
        }
    }

    [void] Refresh ()
    {
        try
        {
            $irmParams = @{
                Uri         = $("{0}api/commpoint/{1}" -f $this.BaseUri, $this.Id)
                Credential  = $this.Credential
                Method      = 'GET'
                ErrorAction = 'Stop'
                Insecure    = $
            }
            $data = (Invoke-RestMethod @irmParams).Data
        }
        catch { throw $_.Exception }

        # assign properties that can be coerced as-is
        $usableProps = 'Name', 'FolderPath', 'Mode', 'State', 'CommPointType', 'OutQueueSize', 'ConnectionCount', 'SentCount'
        $usableProps.ForEach({$this.$_ = $data.$_})
 
        # deserialize timespan properties
        $xmlTsProps = 'InputIdleTime', 'OutPutIdleTime', 'Uptime'
        foreach ($prop in $xmlTsProps)
        {
            try { $this.$prop = [XmlConvert]::ToTimeSpan($data.$prop) }
            catch { $this.$prop = 0 }
        }
    }
<#
    hidden [void] UpdateSession ()
    {
        if (!$this.TestSession()) { $this.RepairSession() }
    }

    hidden [bool] TestSession()
    {
        try
        {
            $iwrParams = @{
                Uri         = $this.BaseUri
                WebSession  = $this.Session
                Method      = 'GET'
                ErrorAction = 'Stop'
            }
            $response = Invoke-WebRequest @iwrParams
            $summary = Get-ResponseSummary $response

            if ($response.BaseResponse.StatusDescription -ne 'OK') { throw [ExternalException] $summary }
            else { return $true }
        }
        catch { return $false }
    }

    hidden [void] RepairSession()
    {
        try
        {
            $psCreds = $this.Session.Credential | ConvertFrom-NetworkCredential
            $newSession = New-WebSession -Uri $this.BaseUri -Credential $psCreds -ErrorAction Stop

            if ($newSession) { $this.Session = $newSession }
        }
        catch 
        {
            Write-Error "Couldn't establish new session with $($this.BaseUri)." 
            throw $_.Exception
        }
    }
#>
    hidden [void] Invoke ([CommPointAction] $Action)
    {
        $this.Refresh()

        $headers = [Dictionary[[String],[String]]]::new()
        $headers.Add('Content-Type','text/plain')

        $iwrParams = @{
            Uri         = $("{0}api/commpoint/{1}/state" -f $this.BaseUri, $this.Id)

            Method      = 'PUT'
            Body        = $Action
            ContentType = 'Text/Plain'
            Credential  = $this.Credential
            ErrorAction = 'Continue'
        }
        
        $response = Invoke-WebRequest @iwrParams
        $summary = Get-ResponseSummary $response

        switch ($response.BaseResponse.StatusCode -as [int]) 
        {
            204 { Write-Verbose "$Action request accepted. ($summary)" }
            500 { throw [ExternalException] "State change failed! ($summary)" }
            400 
            {
                $longMessage = "Valid actions are supposedly $([Enum]::GetNames([CommPointAction]) -join ', '), " + 
                "but the server says it can't $Action communication points. ($summary)"
                throw [ArgumentOutOfRangeException] $longMessage
            }
            default { throw "This wasn't in the docs. Not sure what happened. ($summary)" }
        }
    }
}

<#
function New-WebSession 
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [pscredential]
        $Credential,
    
        [Parameter(Mandatory)]
        [uri]
        $Uri,
    
        [switch]
        $Insecure
    )

    try
    {
        $iwrParams = @{
            Credential      = $Credential
            Uri             = $Uri
            Method          = 'Get'
            SessionVariable = 'session'
            Insecure        = $Insecure
            ErrorAction     = 'Stop'
        }
        $response = Invoke-WebRequest @iwrParams
        $summary = Get-ResponseSummary $response

        if ($response.BaseResponse.StatusDescription -ne 'OK') { throw [ExternalException] $summary }
        else 
        {
            Write-Verbose $summary 
            return $session
        }
    }
    catch
    {
        Write-Error "Couldn't connect to API! URI: $BaseUri"
        throw $_.Exception
    }
}
#>
function Get-ResponseSummary
{
    [OutputType([string])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
        [HtmlWebResponseObject[]]
        $Response
    )

    process 
    {    
        foreach ($r in $Response)
        {
            $blanks = @(
                $r.BaseResponse.ProtocolVersion
                $r.BaseResponse.StatusCode -as [int]
                $r.BaseResponse.StatusDescription
                $r.BaseResponse.Method
                $r.BaseResponse.ResponseURI
            )

            '[HTTP/{0} {1} {2}] {3} {4}' -f $blanks
        }
    }
}

function ConvertFrom-NetworkCredential
{
    [OutputType([pscredential])]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
        [NetworkCredential[]]
        $Credential
    )

    process
    {
        foreach ($netCred in $Credential)
        {
            if ($netCred.Domain)
            { 
                $userName = '{0}\{1}' -f $netCred.Domain, $netCred.User
            }
            else { $userName = $netCred.UserName }

            [pscredential]::new($userName, $netCred.SecurePassword)
        }
    }
}


$a, $b, $c, $d = $id.ForEach({[CommPoint]::new($BaseUri, $_, $Credential)})
