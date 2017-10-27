# requires -Version 4.0
# requires -Modules ActiveDirectory
<#
.NOTES
    Created on:   10/26/2017
    Created by:   Andy Simmons
    Organization: St. Luke's Health System
    FileName:     Export-GroupMembers.ps1

.SYNOPSIS
    Exports membership from one or more AD groups to some CSV files.

.PARAMETER GroupName
    Name of one or more Active Directory groups.

.PARAMETER ExportFolder
    Directory where export files will be created.

.PARAMETER NoClobber
    Don't overwrite export files if they already exist.

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

# Get-ADGroupMember can't deal with large groups, but we can work around it
# with Get-ADGroup/Get-ADObject.
$adGroups = @($GroupName | Get-ADGroup -ErrorAction Continue -Properties 'Members')
if ($adGroups.Count -ne $GroupName.Count)
{
    $difference = $GroupName.Count - $adGroups.Count
    Write-Warning "Skipping export for $difference groups (member lookup error)!"
}

if (!$adGroups) { throw "Member lookup failed for all groups. Aborting." }

# Validate the export directory
if (!$ExportFolder.Exists -and $PSCmdlet.ShouldProcess($ExportFolder.FullName, 'create directory')) 
{
    try 
    { 
        Write-Verbose "Directory '$ExportFolder' not found. Creating."
        New-Item -ItemType Directory -Path $ExportFolder -ErrorAction Stop > $null 
    }
    catch { throw "Couldn't create directory '$ExportFolder'. Bailing." }
}

Write-Verbose ("Exporting membership information for {0} groups" -f $adGroups.Count)

$exportCounter = 0

foreach ($adGroup in $adGroups)
{
    $outFile = "{0}\{1}.csv" -f $ExportFolder.FullName, $adGroup.DistinguishedName
    
    if ($PSCmdlet.ShouldProcess($outFile, 'export membership to file'))
    {
        # We just have DN strings describing members, and DNs change, so we'll grab the 
        # corresponding AD objects to export a few other identifiers.
        $adObjects = @($adGroup.Members | Get-ADObject -Properties 'ObjectSID' -ErrorAction Continue)
        
        if ($adObjects.Count -eq $adGroup.Members.Count) { $content = $adObjects }
        else 
        {
            if ($adGroup.Members)
            {
                $difference = $adGroup.Members.Count - $adObjects.Count
                Write-Warning "Member lookup failed for $difference member(s)! Exporting distinguished names only."
            
                # convert the strings to single-property objects, for consistency across exports
                $content = $adGroup.Members.ForEach({ [pscustomobject] @{ 'DistinguishedName' = $_ } })
            }
            else 
            {
                Write-Warning "'$($adGroup.Name)' group has no members. Still generating an export file, but it'll be empty..." 
                $content = $null 
            }
        }

        if ($NoClobber -and [IOFileInfo]$outFile.Exists)
        {
            Write-Warning "Skipping file '$outFile'. It's already there and I'm not supposed to clobber it."
        }
        else
        {
            try
            {
                $ecsvParams = @{
                    Path              = $outFile
                    NoTypeInformation = $true
                    Confirm           = $false
                    ErrorAction       = 'Stop'
                }
                $content | Export-Csv @ecsvParams
                $exportCounter++
               }
            catch
            {
                Write-Error "Couldn't write to export file '$outFile'!"
                Write-Error $_.Exception.Message
            }
        }
    }
}

"All done! {0} group exports can be found here: '{1}'" -f $exportCounter, $ExportFolder.FullName
