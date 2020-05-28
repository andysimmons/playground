<#
.NOTES
    Name:   ntuserTroubleshoot.ps1
    Author: Andy Simmons
    Date:   05/27/2020

.SYNOPSIS
    Work in progress. Need a way to report on systems where any of the local user
    profiles are missing NTUser.dat.

.DESCRIPTION
    Right now, this just throws an error if it detects a missing NTUser.dat.

.PARAMETER CheckAll
    This checks ALL local profiles (including the system and service profiles). I don't
    think this would be useful in production. Just dev.
#>
#Requires -RunAsAdministrator
#Requires -Version 4
[CmdletBinding()]
param (
    [switch] $CheckAll
)

$regKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'

if ($CheckAll) { 
    # grab the root for all profiles (including system/service)
    $profileRoot = (Get-ChildItem -Path $regKey).ForEach({ $_.GetValue('ProfileImagePath') })
}
else { 
    # just grab the root for full profiles
    $profileRoot = Get-ChildItem -Path $regKey | ForEach-Object {
        if ($_.GetValue('FullProfile')) { $_.GetValue('ProfileImagePath') }
    }
}

# grab file information about each NTUSER.DAT
$fileInfo = $profileRoot.ForEach({ [IO.FileInfo] "$_\NTUSER.DAT" })
$summary = $fileInfo | 
    Sort-Object -Property Exists, FullName |
    Select-Object -Property FullName, Exists | 
    Out-String

# if any of these are missing, throw an exception
if ($fileInfo.Where({-not $_.Exists})) {
    # Nic just wants true/false for now
    #throw [System.IO.FileNotFoundException] "$env:COMPUTERNAME has corrupt profile(s)!`n`n$summary"
    Write-Verbose "$env:COMPUTERNAME has corrupt profile(s)!`n`n$summary"
    $false
    $fileInfo.Where({ -not $_.Exists }).ForEach({ Write-Warning $_.FullName })
}
else {
    Write-Verbose "No (obviously) corrupt profiles found on $env:COMPUTERNAME. Hooray!`n`n$summary"
    return $true
}
