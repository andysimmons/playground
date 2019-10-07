# requires -Version 4.0
# requires -Modules ActiveDirectory
<#
.NOTES
    Created on:   10/4/2019
    Created by:   Andy Simmons
    Organization: St. Luke's Health System
    FileName:     Copy-ADGroupMember.ps1

.SYNOPSIS
    Copies members from one AD security group into another. 

.PARAMETER SourceGroup
    Group from which members will be copied

.PARAMETER DestGroup
    Group into which the members will be copied

.PARAMETER BatchSize
    Maximum number of members to add to a group in a single operation. All users
    will still be copied, but will be broken up into batches no larger than this.

    AD web services tends to freak out around 5k users per operation.

.PARAMETER Interval
    Time (in seconds) to wait between batches.

.EXAMPLE
    .\Copy-ADGroupMember.ps1 -SourceGroup 'GroupA' -DestGroup 'GroupB'

    Ensures that every member of 'GroupA' is also a member of 'GroupB'
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)]
    [String]
    $SourceGroup,

    [Parameter(Mandatory)]
    [string]
    $DestGroup,

    [int]
    $BatchSize = 1000,

    [int]
    $Interval = 5
)

function Add-ThrottledADGroupMember {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $Identity,

        [Parameter(Mandatory)]
        [object[]]
        $Members,

        [int]
        $BatchSize = 1000,

        [int]
        $Interval = 5
    )

    if ($Members.Count -le $BatchSize) {
        Write-Output "[$(Get-Date -f G)] Adding $($Members.Count) members to $Identity."
        Add-ADGroupMember -Identity $Identity -Members $Members
    }
    else {
        # Too many for one batch. Split this into two collections (one for now, one for later)
        $leftovers = $Members | Select-Object -Skip $BatchSize
        $Members = $Members | Select-Object -First $BatchSize

        Write-Output "[$(Get-Date -f G)] Adding $BatchSize members to $Identity."
        Add-ADGroupMember -Identity $Identity -Members $Members

        Write-Output "[$(Get-Date -f G)] Waiting $Interval seconds to add remaining $($leftovers.Count) members."
        Start-Sleep -Seconds $Interval
        $atadgmParams = @{
            Identity  = $Identity
            Members   = $leftovers
            BatchSize = $BatchSize
            Interval  = $Interval
        }
        Add-ThrottledADGroupMember @atadgmParams
    }
}

Write-Output "[$(Get-Date -f G)] Checking membership on $SourceGroup"
try {
    # Get-ADGroupMember barfs if a group contains more members than AD web services can return (default
    # is 5k at the time of writing), but we can work around that with Get-ADGroup
    $sourceADGroup = Get-ADGroup -Identity $SourceGroup -ErrorAction 'Stop' -Properties 'Members'
    $destADGroup = Get-ADGroup -Identity $DestGroup -ErrorAction 'Stop' -Properties 'Members'
}
catch {
    Write-Error "[$(Get-Date -f G)] Barfed retrieving AD groups. Bailing."
    throw $_.Exception
}

$newMember = $sourceADGroup.Members.Where({ $_ -notin $destADGroup.Members })
$newMemberCount = $newMember.Count

if (-not $newMemberCount) {
    Write-Warning "[$(Get-Date -f G)] Source group '$SourceGroup' has no members that aren't already in '$DestGroup'. Nothing to do."
    exit 0
}

if ($PSCmdlet.ShouldProcess($DestGroup, "copy $newMemberCount members from $SourceGroup")) {
    $atadgmParams = @{
        Identity  = $destADGroup
        Members   = $newMember
        BatchSize = $BatchSize
        Interval  = $Interval
    }
    Add-ThrottledADGroupMember @atadgmParams
}