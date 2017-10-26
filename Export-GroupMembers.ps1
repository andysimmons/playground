<#
.SYNOPSIS
    Exports membership from one or more AD groups to text files (one
    per group) containing the SID of each member.

.PARAMETER GroupName
    Name (SAM or Distinguished Name) of one or more AD groups

.PARAMETER ExportFolder
    Directory containing files to be exported

.PARAMETER NoClobber
    If used, existing export files will not be overwritten.

.EXAMPLE
    .\Export-GroupMembers.ps1 -GroupName (Get-Content .\GG_Groups.txt) -Verbose

    Pulls a list of groups from the file .\GG_Groups.txt and exports membership for each.
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)]
    [String[]]
    $GroupName,

    [IO.DirectoryInfo]
    $ExportFolder = '.\GroupExport',

    [switch]
    $NoClobber
)

Write-Verbose ("Looking up membership for {0} groups..." -f $GroupName.Count)
$adGroups = $GroupName | Get-ADGroup -ErrorAction Continue -Properties 'Members'
if ($adGroups.Count -ne $GroupName.Count)
{
    $difference = $GroupJName.Count - $adGroups.Count
    Write-Warning "AD lookup failed for $difference groups."
}

if (!$adGroups) { throw "AD lookup failed for all groups. Aborting." }

if (!$ExportFolder.Exists -and $PSCmdlet.ShouldProcess($ExportFolder.FullName, 'create directory')) {
    Write-Verbose "Directory '$ExportFolder' not found. Creating."
    try { New-Item -ItemType Directory -Path $ExportFolder -ErrorAction Stop > $null }
    catch { throw "Couldn't create directory '$ExportFolder'." }
}

foreach ($adGroup in $adGroups)
{

    $isEmpty = ! [bool] $adGroup.Members 
    $fileName = "{0}\{1}.txt" -f $ExportFolder.FullName, $adGroup.DistinguishedName
    if ($PSCmdlet.ShouldProcess($fileName, 'Export Membership'))
    {
        $objectSIDs = ($adGroup.Members | Get-ADObject -Properties 'ObjectSID').ObjectSID.Value
        if (!$objectSIDs)
        {
            if (!$isEmpty)
            {
                Write-Warning "SID lookup failed - exporting Distinguished Names instead."
                $content = $adGroup.Members
            }
            else { $content = $null }
        }
        else { $content = $objectSIDs }

        # We're already in a ShouldProcess code block, override whatif/confirm
        if ($NoClobber -and [IOFileInfo]$fileName.Exists)
        {
            Write-Warning "Skipping file '$fileName' because it already exists."
        }
        else
        {
            $content | Out-File -FilePath $fileName -NoClobber:$NoClobber -WhatIf:$false -Confirm:$false
        }
    }
}

"Script completed! Any exported files can be found in: {0}" -f $ExportFolder.FullName

"To get a list of object names rather than SIDs, use something like this:
(Get-Content '" + 
    $ExportFolder.FullName + "\<GroupDN>.txt'" + 
    ').ForEach({[Security.Principal.Identifier]::new($_.Trim()).Translate([Security.Principal.NTAccount]).Value'
