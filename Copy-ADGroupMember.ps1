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
    $DestGroup
)

Write-Verbose "Checking membership on $SourceGroup"

try {
    # Get-ADGroupMember can't deal with large groups, but we can work around it
    # with Get-ADGroup.
    $sourceADGroup = Get-ADGroup -Identity $SourceGroup -ErrorAction 'Stop' -Properties 'Members'
    $destADGroup = Get-ADGroup -Identity $DestGroup -ErrorAction 'Stop'
}
catch {
    Write-Error 'Barfed retrieving AD groups. Bailing.'
    throw $_.Exception
}

$memberCount = $sourceADGroup.Members.Count

if (-not $memberCount) {
    Write-Warning "Source group '$SourceGroup' has no members. Bailing."
    exit 0
}

if ($PSCmdlet.ShouldProcess($DestGroup, "copy $memberCount members from $SourceGroup")) {
    Add-ADGroupMember -Identity $destADGroup -Members $sourceADGroup.Members
}
