<#
.NOTES
    Name:   ntuserTroubleshoot.ps1
    Author: Andy Simmons
    Date:   05/27/2020

.SYNOPSIS
    Work in progress. Need a way to report on systems where any of the local user
    profiles are missing the user hive.

.DESCRIPTION
    Returns $true if every profile has a user hive (ntuser.dat), or $false if it is missing
    for any profile.

.PARAMETER CheckAll
    This checks all local profiles (including the system and service profiles). I don't
    think this would be useful in production. Just dev.

.PARAMETER Repair
    Specifies whether the profiles should be "repaired". The profile root will be renamed and
    the profile regkey will be deleted (after being exported to the renamed profile root).
#>
#Requires -RunAsAdministrator
#Requires -Version 5
[CmdletBinding(SupportsShouldProcess)]
param (
    [switch] $CheckAll,

    [switch] $Repair
)

class ProfileInfo {
    # properties
    [Microsoft.Win32.RegistryKey] $ProfileKey
    [IO.FileInfo]                 $UserHive

    # constructors
    ProfileInfo ([Microsoft.Win32.RegistryKey] $ProfileKey) {
        $this.ProfileKey = $ProfileKey
        $this.Refresh()
    }

    # methods
    
    # override ToString() to show profile root
    [string] ToString() {
        return $this.ProfileKey.GetValue('ProfileImagePath')
    }

    # refresh the state of the user hive
    [ProfileInfo] Refresh () {
        $this.UserHive = [IO.FileInfo] "$this\NTUSER.DAT"
        return $this
    }

    # is there an NTUSER.DAT for this profile?
    [bool] UserHiveExists() { return $this.UserHive.Exists }

    # If we're missing the user hive, rename the profile directory and
    # nuke the profile regkey. This is super quick and dirty... probably shouldn't
    # reuse this class outside this specific script 
    [void] Archive() {
        if ($this.UserHiveExists()) { 
            # don't touch it if the user hive is intact
            Write-Warning "User hive found in $this. I'm not gonna archive that..."
        }
        else {
            try {
                # rename profile root directory
                $oldPath = $this.ToString()
                $newPath = "${oldPath}_corrupt_$(Get-Date -f 'yyyyMMdd-hhmmss')"
                Rename-Item -Path $oldPath -NewName $newPath -ErrorAction Stop -Force
                
                # export profile regkey and then remove it (note:
                # I'm not aware of a PS-native way to do this, so
                # error handling is kludgy, but it's something...)
                $exportPath = "$newPath\ProfileKey.reg"
                reg export $this.ProfileKey $exportPath /y | Out-Null
                if (-not (Test-Path -Path $exportPath)) {
                    throw "Couldn't export profile registry key to '$exportPath'. " + 
                        "'$oldPath' is now '$newPath'. Leaving '$($this.ProfileKey)' in place."
                }
                Remove-Item -Path $this.ProfileKey.PSPath -ErrorAction Stop -Force

                Write-Verbose "'$oldPath' archived to '$newPath'."
                Write-Verbose "User profile regkey exported to '$exportPath'."
            }
            catch {
                Write-Error "Archive failed for '$this'!"
                Write-Error $_.Exception.Message
            }
        }
    }
}

function Get-ProfileInfo {
    [CmdletBinding()]
    param ([switch] $CheckAll)
    $parentKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'

    if ($CheckAll) { 
        # grab the all profiles (including system/service)
        $profileKey = Get-ChildItem -Path $parentKey
    }
    else { 
        # just grab full profiles
        $profileKey = (Get-ChildItem -Path $parentKey).Where( { $_.GetValue('FullProfile') } )
    }
    
    # return a collection of [ProfileInfo] objects
    $profileKey.ForEach( { [ProfileInfo]::new($_) } )
}

# grab all the profiles
$profileInfo = Get-ProfileInfo -CheckAll:$CheckAll

# are any missing the user hive?
$corruptProfile = $profileInfo.Where( { -not $_.UserHiveExists() })

if ($corruptProfile) {
    Write-Verbose "$env:COMPUTERNAME has corrupt profile(s)!"
    $false

    # archive corrupted profiles
    if ($Repair) {
        foreach ($p in $corruptProfile) {
            if ($PSCmdlet.ShouldProcess($p, 'ARCHIVE')) { $p.Archive() }
        }
    }
}
else {
    Write-Verbose "No (obviously) corrupt profiles found on $env:COMPUTERNAME. Hooray!"
    return $true
}
