#requires -Version 4.0
#requires -Modules ActiveDirectory, CimCmdlets
<#
.NOTES
    Name:    Get-ComputerUser.ps1
    Author:  Andy Simmons
    Version: 0.1.0
    URL:     https://github.com/andysimmons/playground/blob/master/Get-ComputerUser.ps1

.SYNOPSIS
	Generates a list of active users on one or more computers.

.DESCRIPTION
    This script pulls human user profiles from a Windows computer,
    pulls their name and email address from Active Direcotry, and
    spits out a summary.

.PARAMETER ComputerName
    One or more computer names to query for user profile information.

.PARAMETER ShowFullErrorDetail
    Most errors are suppressed and replaced with a terse warning, by
    default. This switch disables that suppression.

.EXAMPLE 
    Get-ComputerUser.ps1 -ComputerName 'Computer1'

    Pull a list of users from computer1.
.EXAMPLE
    'Computer1','Computer2' | Get-ComputerUser.ps1 | Format-Table -AutoSize

    Pull a list of users from computer1 and computer2.
#>
[CmdletBinding()]
param (
    [Parameter(ValueFromPipeline, Position = 0)]
    [string[]] $ComputerName = $env:COMPUTERNAME,

    [switch] $ShowFullErrorDetail
)

process
{
    foreach ($computer in $ComputerName)
    {
        try
        {
			Write-Verbose "Enumerating user profiles on $computer ..."

			$cimParams = @{
				ComputerName = $computer
				ClassName    = 'Win32_UserProfile'
				ErrorAction  = 'Stop' 
			}
            $userProfiles = (Get-CimInstance @cimParams).Where({ $_.LastUseTime -and (-not $_.Special) }) 
        }
        catch
        {
            if ($ShowFullErrorDetail) { Write-Error -ErrorRecord $_ }

            Write-Warning "Couldn't enumerate user profiles on $computer. Skipping."
            $userProfiles = $null
        }

        foreach ($userProfile in $userProfiles)
        {
            try   
            { 
                Write-Verbose "Looking up SID '$($userProfile.SID)' ..."
                $user = Get-ADUser $userProfile.SID -Properties 'Mail' -ErrorAction 'Stop'
            }
            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
            {
                if ($ShowFullErrorDetail) { Write-Error -ErrorRecord $_ }
                
                $oldUserName = $userProfile.LocalPath.Split('\')[-1]
                Write-Warning "User $oldUserName may have been deleted (no such SID in Active Directory). Skipping."
                $user = $null
            }
            catch
            {
                if ($ShowFullErrorDetail) { Write-Error -ErrorRecord $_ }
                Write-Warning "Error querying AD for user profile..."

                [PSCustomObject] @{
                    Name    = 'AD_QUERY_ERROR'
                    Mail    = 'AD_QUERY_ERROR'
                    Enabled = $null
                }
            }

            # return interesting data about the computer/profile/user
            if ($user)
            {
                [PSCustomObject] @{
                    ComputerName = $computer
                    User         = $user.Name
                    UserMail     = $user.Mail
                    LastUseTime  = $userProfile.LastUseTime
                    LocalPath    = $userProfile.LocalPath
                    Enabled      = $user.Enabled
                }
            }
        }		
    }
}
