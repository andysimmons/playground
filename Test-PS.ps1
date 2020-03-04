using namespace System.Management.Automation
#Requires -Version 5
<#
.NOTES
    Name:   Test-PS.ps1
    Author: Andy Simmons
    Date:   3/4/2020

.SYNOPSIS
    Dummy script to test PS syntax highlighting and PS Host rendering.

.DESCRIPTION
    Implements most of what I'm interested in seeing from a PowerShell syntax highlighter both
    from the editor and at runtime (in the PS host).

    I haven't found any VS Code themes that visually differentiate all PS output streams,
    generate legible progress bars, and emphasize elements I typically care about within
    script logic, so I wrote this to see it all in once place.

.PARAMETER Duration
    Duration (sec) to display a progress bar

.EXAMPLE
    1..5 | Test-PS.ps1 -Verbose
    
    Generates 5 different progress bars that respectively take 1, 2, 3, 4, and 5 
    seconds each to complete, outputting a dummy object between each.
    
    Incidentally, this also tests some rendering behavior with the PS host. I noticed
    some funk with VS Code's PowerShell Integrated Console.
    
    It looks like messages sent in rapid sequence to different streams often render out of order,
    and the progress bar will always scroll out of view as new objects hit the output stream.
    
    Microsoft's other PS hosts don't seem to be affected (powershell.exe and ISE).
 #>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param (
    [Parameter(ValueFromPipeline)]
    [ValidateScript({ $_ -ge 1 })]
    [int[]] $Duration = 5
)

begin {
    <#
    .SYNOPSIS
        Prepends a timestamp on a string.
     #>
    function Add-TimeStamp {
        param (
            [Parameter(ValueFromPipeline, Position = 0)]
            [string[]]
            $Message,

            [string]
            $Format = 'MMM dd hh:mm:ss.fffffff tt'
        )
        process { $Message.ForEach({ "[$(Get-Date -f $Format)] $_" }) }
    }

    # dummy class to generate progress bars and return objects
    class DummyClass {

        # properties
        [int] $Duration
        [string] $Noun
        [string] $Verb
        [string] $Status

        # constructors
        DummyClass ()                { $this.Initialize(10) }
        DummyClass ([int] $Duration) { $this.Initialize($Duration) }

        # methods
        hidden [void] Initialize ([int] $Duration) { 
            $this.Duration = $Duration
            $this.Status = 'Ready' 
            $this.Noun = 'dummy object'
            $this.Verb = 'Wait around for {0} seconds' -f $this.Duration
        }

        [string[]] GetAction () { return @($this.Noun, $this.Verb) }

        [DummyClass] Invoke () {
            $this.Status = 'Running'

            # generate a progress bar that takes the specified duration (sec) to
            # complete, and updates every 100 ms
            foreach ($i in 1..($this.Duration * 10)) {
                $wpParams = @{
                    Activity         = $this.Verb
                    Status           = $this.Status
                    PercentComplete  = 10 * $i / $this.Duration
                    SecondsRemaining = $this.Duration - ($i / 10)
                }
                Write-Progress @wpParams
                Start-Sleep -Milliseconds 100
            }

            $this.Status = 'Complete'
            return $this
        }
    }

    # testing string interpolation with escape characters, preference vars, and scopes
    if ($ProgressPreference -eq [ActionPreference]::SilentlyContinue) {
        $implication = "you won't see progress bars"
        Write-Warning "`$ProgressPreference is '$ProgressPreference', meaning $local:implication."
    }
}

process { 
    foreach ($d in $Duration) {
        $dummy = [DummyClass]::new($d) 
        if ($PSCmdlet.ShouldProcess( @($dummy.GetAction()) )) { $dummy.Invoke() }
    }
}

end { 
    # TODO: figure out why VS Code's integrated console sometimes renders these out of sequence
    Add-TimeStamp -Message "001 - Testing output streams. Here's (forced) verbose." | Write-Verbose -Verbose
    Add-TimeStamp -Message "002 - This is standard output." | Write-Output
    Add-TimeStamp -Message "003 - Here's a warning!"  | Write-Warning
    Add-TimeStamp -Message "004 - HERE'S AN ERROR!" | Write-Error
    Add-TimeStamp -Message "005 - Script finished." | Write-Verbose
}
