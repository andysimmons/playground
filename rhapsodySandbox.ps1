#requires -Module TunableSSLValidator
#requires -Version 5
using namespace Microsoft.PowerShell.Commands
using namespace System.Xml
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [pscredential]
    $Credential,

    [int[]]
    $id = 6223..6226,

    [uri]
    $BaseURI = 'https://rhapsodyindev01.slhs.org:8444',

    [switch]
    $Insecure
)

<#
scratch area

https://rhapsodyindev01.slhs.org:8444/api/commpoint/6223
{
    "data": {
        "id": 6223,
        "name": "eICUResTest 1",
        "folderPath": "Main/zzTesting/Tims/RES Automation Testing",
        "mode": "OUTPUT",
        "state": "RUNNING",
        "commPointType": "Sink",
        "inputIdleTime": "PT27H48M27.009S",
        "outQueueSize": 0,
        "outputIdleTime": "PT27H48M27.009S",
        "uptime": "PT27H48M27.009S",
        "connectionCount": 1,
        "sentCount": 0
    },
    "error": null
}

[System.Xml.XmlConvert]::ToTimeSpan("PT27H48M27.009S")

#>

# establish session
try {
    $irmParams = @{
        Credential      = $Credential
        Uri             = $BaseURI
        Method          = 'Get'
        SessionVariable = 'session'
        Insecure        = $Insecure
        ErrorAction     = 'Stop'
    }
    Invoke-RestMethod @irmParams
    Write-Verbose "Session established: $BaseURI"
}
catch {
    Write-Error "Couldn't connect to API! URI: $BaseURI"
    throw $_.Exception
}

class CommPoint {
    # properties
    [int]               $Id
    [string]            $Name
    [string]            $FolderPath
    [string]            $Mode
    [string]            $State
    [string]            $CommPointType
    [timespan]          $InputIdleTime
    [int]               $OutQueueSize
    [timespan]          $OutputIdleTime
    [timespan]          $Uptime
    [int]               $ConnectionCount
    [int]               $SentCount
    [WebRequestSession] $Session
    [Uri]               $BaseURI


    # constructors
    CommPoint (
        [WebRequestSession] $Session,
        [Uri]               $BaseURI,
        [int]               $Id
    ) {
        $this.Id = $Id
        $this.Session = $Session
        $this.BaseURI = $BaseURI

        $this.Initialize()
    }

    # methods
    hidden [void] Initialize () {

        try {
            $irmParams = @{
                Uri         = $("{0}api/commpoint/{1}" -f $this.BaseURI, $this.Id)
                WebSession  = $this.Session
                Method      = 'Get'
                ErrorAction = 'Stop'
            }
            $data = (Invoke-RestMethod @irmParams).Data
        }
        catch {
            Write-Error "Initialization failed."
            throw $_.Exception
        }

        $this.Name = $data.Name
        $this.FolderPath = $data.folderPath
        $this.Mode = $data.mode
        $this.State = $data.State
        $this.CommPointType = $data.commPointType
        $this.InputIdleTime = [XmlConvert]::ToTimeSpan($data.inputIdleTime)
        $this.OutQueueSize = $data.outQueueSize
        $this.OutputIdleTime = [XmlConvert]::ToTimeSpan($data.outputIdleTime)
        $this.Uptime = [XmlConvert]::ToTimeSpan($data.uptime)
        $this.ConnectionCount = $data.connectionCount
        $this.SentCount = $data.sentCount
    }
}

$id.ForEach({[CommPoint]::new($session, $BaseURI, $_)})
