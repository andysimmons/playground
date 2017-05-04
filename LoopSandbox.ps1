<#
.NOTES
     Created on:   5/3/2016
     Created by:   Andy Simmons
     Organization: St. Luke's Health System
     Filename:     LoopSandbox.ps1

.SYNOPSIS    
    Tests loop performance.

.DESCRIPTION
    Invokes various types of loop logic against a collection, and
    reports the results.

.PARAMETER Collection
    Collection to test against.

.PARAMETER LoopLogic
    One or more scriptblocks of loop logic.
#>
#requires -version 5.0
using namespace System.Collections

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [array] $Collection,

    [scriptblock[]] $LoopLogic = @(
        { $Collection | .{ process { $_ } } }, 
        { $Collection.ForEach({ $_ }) },
        { $Collection | ForEach-Object { $_ } },
        { ForEach-Object -InputObject $Collection { $_ } },
        { foreach ($element in $Collection) { $element } }
    )
)

#region Classes
class LoopRunner 
{
    # Properties
    [scriptblock] $ScriptBlock
    [int]         $Ticks
    [array]       $Result

    # Constructor
    LoopRunner([scriptblock] $ScriptBlock) 
    { 
        $this.Initialize($ScriptBlock) 
    }
    
    # Methods
    [void] Initialize ([scriptblock] $ScriptBlock)
    {
        $this.ResetCounters()
        $this.ScriptBlock = $ScriptBlock      
    }
    
    [void] Invoke () 
    {
        $this.ResetCounters()
        $this.Ticks = (Measure-Command -Expression { $this.Result = $this.ScriptBlock.Invoke() }).Ticks
    }
    
    [void] ResetCounters ()
    {
        $this.Ticks  = -1
        $this.Result = @()
    }

    [bool] Validate ()
    {
        return (($this.Ticks -ne -1) -and ($this.Result))
    }
}
#endregion Classes

#region Functions
<#
.SYNOPSIS
    Eats a collection of LoopRunner objects and spits back a report.
#>
function Get-LoopRunnerView
{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline)]
        [LoopRunner[]] $InputObject
    )

    begin { [Collections.ArrayList] $loopView = @() }

    process
    {
        foreach ($element in $InputObject)
        {
            # Build each element "view" and add it to the view collection
            $elementView = [pscustomobject] @{
                'ScriptBlock'    = $element.ScriptBlock
                'RelativeSpeed'  = 'Unknown'
                'ItemsPerSecond' = -1
                'Ticks'          = $element.Ticks 
                'ItemsProcessed' = $element.Result.Length
                '_IsValid'       = $element.Validate()
            }

            $loopView.Add($elementView) > $null
        }
    }

    end
    {
        # Pick the fastest method and calculate relative comparisons
        $speedBaseline  = ($loopView | Measure-Object -Property 'Ticks' -Minimum).Minimum

        foreach ($element in $loopView)
        {
            if ($element._IsValid -and $element.ItemsProcessed)
            {
                $secondsElapsed         = [timespan]::FromTicks($element.Ticks).TotalSeconds
                $element.ItemsPerSecond = [math]::Round($element.ItemsProcessed / $secondsElapsed)
                $element.RelativeSpeed  = ($speedBaseline / $element.Ticks).ToString("#.#%")
            }
        }

        $loopView | Select-Object -Property * -ExcludeProperty _* | Sort-Object -Property 'Ticks' 
    }
}
#endregion Functions

$loopRunners = [LoopRunner[]] $LoopLogic

# Invoke each loop runner with a progress bar
for ($i = 0; $i -lt $loopRunners.Length; $i++) 
{
    $writeProgressParams = @{
        Id              = 10
        Activity        = "Looping through $($Collection.Length) items"
        Status          = "Method $i of $($loopRunners.Length): { $($loopRunners[$i].ScriptBlock) }"
        PercentComplete = 100 * $i / $loopRunners.Length
    }
    Write-Progress @writeProgressParams
    
    $loopRunners[$i].Invoke()
}
Write-Progress -Activity  $writeProgressParams['Activity'] -Completed

$loopView = $loopRunners | Get-LoopRunnerView
$loopView | Format-Table -AutoSize
