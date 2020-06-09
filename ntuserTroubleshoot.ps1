<#
.NOTES
    Name:        ntuserTroubleshoot.ps1
    Author:      Andy Simmons
    Date:        05/27/2020
    Last Update: 06/09/2020

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

.PARAMETER SaveProfileKey
    Save each profile registry keys to its corresponding archived/renamed profile root directory.
    
    NOTE: This will not save profile registry keys that refer to a non-existent profile directory.
    Those will be destroyed. 

.PARAMETER LogFile
    Text file used to capture script output (PS transcript).

.PARAMETER LogName
    Specifies which Windows event log should be written to regarding profile repair.

.PARAMETER LogSource
    Specifies the log source used when writing to the Windows event log regarding profile repair.

.PARAMETER EventId
    Specifies the event ID used when writing to the Windows event log regarding profile repair.
#>
#Requires -RunAsAdministrator
#Requires -Version 5
[CmdletBinding(SupportsShouldProcess)]
param (
    [switch]
    $CheckAll,

    [switch]
    $Repair,

    [switch]
    $SaveProfileKey,

    [IO.FileInfo]
    $LogFile = "C:\logs\ntuserTroubleshoot.log",

    [string]
    $LogName = 'Application',

    [string]
    $LogSource = 'UserHiveTroubleshooter',

    [int]
    $EventId = 6032
)

#region classes

# need a custom class to add some remediation/logging capabilities
class ProfileInfo {
    # properties
    [Microsoft.Win32.RegistryKey] $ProfileKey
    [IO.FileInfo]                 $UserHive
    [string[]]                    $LogMessage

    # constructors
    ProfileInfo ([Microsoft.Win32.RegistryKey] $ProfileKey) {
        $this.ProfileKey = $ProfileKey
        $this.Refresh()
        $line = '---------------------------------------------------------------------'
        $this.LogMessage = @(
            $line
            "Profile Key: $($this.ProfileKey)"
            "User Hive (File Path): $($this.UserHive)"
            "Profile Directory Exists: $($this.UserDirExists())"
            "User Hive Exists: $($this.UserHiveExists())"
            $line
            ''
            'Log Detail:'
            ''
        )
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

    # is there a profile directory for this regkey?
    [bool] UserDirExists() { return $this.UserHive.Directory.Exists }

    # If we're missing the user hive, rename the profile directory and
    # nuke the profile regkey. This is super quick and dirty... probably shouldn't
    # reuse this class outside this specific script 
    [void] Archive([bool] $SaveProfileKey) {
        if ($this.UserHiveExists()) { 
            # don't touch it if the user hive is intact
            $logInfo = "User hive found in $this. I'm not gonna archive that..."
            $this.LogMessage += $logInfo
            Write-Warning $logInfo
        }
        else {
            try {
                # rename profile root directory (if it exists)
                if ($this.UserDirExists()) {
                    $oldPath = $this.ToString()
                    $newPath = "${oldPath}_corrupt_$(Get-Date -f 'yyyyMMdd-hhmmss')"
                    Rename-Item -Path $oldPath -NewName $newPath -ErrorAction Stop -Force
                
                    $logInfo = "'$oldPath' archived to '$newPath'."
                    $this.LogMessage += $logInfo
                    Write-Verbose $logInfo

                    # I'm not aware of a PS-native way to do export the .reg file, so
                    # error handling is kludgy here, but it's something...
                    if ($SaveProfileKey) { 
                        $exportPath = "$newPath\ProfileKey.reg"
                        reg export $this.ProfileKey $exportPath /y | Out-Null
                        if (-not (Test-Path -Path $exportPath)) {
                            $logInfo = "Couldn't export profile regkey to '$exportPath'. '$oldPath' is now '$newPath'."
                            $this.LogMessage += $logInfo
                            throw $logInfo
                        }
                        else {
                            $logInfo = "User profile regkey exported to '$exportPath'."
                            $this.LogMessage += $logInfo
                            Write-Verbose $logInfo
                        }
                    }
                }
                else {
                    $logInfo = "$($this.UserHive.Directory) is referenced by the profile key, but doesn't exist."
                    $this.LogMessage += $logInfo
                    Write-Verbose
                }

                Remove-Item -Path $this.ProfileKey.PSPath -ErrorAction Stop -Force -Recurse -Confirm:$false
                $logInfo = "Removed registry key: $($this.ProfileKey.PSPath)"
                $this.LogMessage += $logInfo
                Write-Verbose $logInfo
            }
            catch {
                $logInfo = "Archive failed for '$this'!"
                $this.LogMessage += $logInfo
                Write-Error $logInfo

                $logInfo = $_.Exception.Message
                $this.LogMessage += $logInfo
                Write-Error $logInfo
            }
        }
    }
}
#endregion classes

#region functions
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

# Register the log source specified by this rule (to allow writing to the event log)
function Register-LogSource {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $LogName,

        [Parameter(Mandatory)]
        [string]
        $LogSource,

        [string]
        $ComputerName = $env:COMPUTERNAME
    )
    try { 
        $nelParams = @{
            LogName      = $LogName
            Source       = $LogSource
            ComputerName = $ComputerName
            ErrorAction  = 'Stop'
        }
        New-Eventlog @nelParams 
        Write-Verbose "Registered $($this.LogName) log source '$($this.LogSource)' on $($this.LogServer)."
        
    }
    # If the log source already exists, suppress the error and shoot the message to verbose.
    catch [InvalidOperationException] { 
        Write-Verbose $_.Exception.Message 
    } 
    catch { throw $_.Exception }
}

#endregion functions

#region init

# Going to override $WhatIfPreference for these cmdlets since
# we always want logs.
try {
    # try creating the log directory if not already present
    if (-not $LogFile.Directory.Exists) {
        New-Item $LogFile.Directory -ItemType Directory -WhatIf:$false | Out-Null
    }
    # Try the cool/newer Start-Transcript if available
    $stOutput = Start-Transcript -IncludeInvocationHeader $LogFile -ErrorAction 'Stop' -WhatIf:$false | Out-String -Stream
    $isTranscribing = $true
    Write-Verbose $stOutput
}
catch { 
    try {
        # otherwise fall back to legacy Start-Transcript
        $stOutput = Start-Transcript $LogFile -ErrorAction 'Stop' -WhatIf:$false
        $isTranscribing = $true
        Write-Verbose $stOutput
    }
    catch { 
        Write-Warning "Couldn't write to log file '${LogFile}'. Continuing without logging."
        Write-Warning $_.Exception.Message
        $isTranscribing = $false
    }
}
#endregion init

#region main
# grab all the profiles
$profileInfo = Get-ProfileInfo -CheckAll:$CheckAll

# are any missing the user hive?
$corruptProfile = $profileInfo.Where( { -not $_.UserHiveExists() } )

if ($corruptProfile) {
    Write-Warning "$env:COMPUTERNAME has corrupt profile(s)!"

    $corruptProfile | Select-Object -Property @(
        'ProfileKey'
        'UserHive'
        @{ Name = 'UserHiveExists'; Expression = { $_.UserHiveExists() } }
        @{ Name = 'UserDirExists'; Expression = { $_.UserDirExists() } }
    ) | Format-List | Out-String | Write-Warning
    $false

    # if invoked with -Repair, then archive corrupted profiles
    if ($Repair) {
        foreach ($p in $corruptProfile) {
            if ($PSCmdlet.ShouldProcess($p, 'ARCHIVE')) {
                $p.Archive($SaveProfileKey) 

                # write to the Windows event log
                try {
                    Register-LogSource -LogName $LogName -LogSource $LogSource

                    $weParams = @{
                        ComputerName = $env:COMPUTERNAME
                        LogName      = $LogName
                        Source       = $LogSource
                        EventId      = $EventId
                        EntryType    = 'Information'
                        Message      = $p.LogMessage -join "`n"
                        ErrorAction  = 'Stop'
                    }
                    Write-EventLog @weParams
                }
                catch {
                    Write-Warning "I choked trying to write to the Windows event log..."
                    Write-Error $_.Exception
                }
            }
        }
    }
}
else {
    Write-Verbose "No (obviously) corrupt profiles found on $env:COMPUTERNAME. Hooray!"
    $true
}
if ($isTranscribing) { 
    $stOutput = Stop-Transcript | Out-String -Stream
    $isTranscribing = $false
    Write-Verbose $stOutput
}
#endregion main