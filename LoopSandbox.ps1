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
    [object[]] $Collection,

    [scriptblock[]] $LoopLogic = @(
        { $Collection | .{ process { $_ } } }, 
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
    [timespan]    $ProcessingTime
    [int]         $ItemCount
    [double]      $ItemsPerSecond
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
        # Invoke the scriptblock and capture interesting metrics
        $this.ResetCounters()
        $this.ProcessingTime = Measure-Command -Expression { $this.Result = $this.ScriptBlock.Invoke() }
        $this.ItemCount      = $this.Result.Length
        $this.ItemsPerSecond = $this.getItemsPerSecond(0)
    }
    
    [void] ResetCounters ()
    {
        $this.ProcessingTime = -1
        $this.ItemCount      = -1
        $this.ItemsPerSecond = -1
        $this.Result         = @()
    }
    
    hidden [int] getItemCount ()
    {
        if ($this.Result) { return $this.Result.Length }
        else              { return -1 }
    }
    
    hidden [double] getItemsPerSecond ([int] $Precision)
    {
        if (($this.ItemCount -gt 0) -and ($this.ProcessingTime -gt 0))
        {
            return [math]::Round(($this.ItemCount / $this.ProcessingTime.TotalSeconds), $Precision)
        }
        else { return -1 }
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
                'ProcessingTime' = [math]::Round($element.ProcessingTime.TotalSeconds, 3) 
                'CollectionSize' = $element.ItemCount
                'ItemsPerSecond' = $element.ItemsPerSecond
            }

            $loopView.Add($elementView) > $null
        }
    }

    end
    {
        # Pick the fastest method, and set the relative speed on each element before returning
        $baseline = ($loopView | Measure-Object -Property 'ItemsPerSecond' -Maximum).Maximum

        foreach ($element in $loopView)
        {
            # ItemsPerSecond of -1 means the LoopRunner hasn't been invoked
            if ($element.ItemsPerSecond -ne -1)
            {
                $element.RelativeSpeed = ($element.ItemsPerSecond / $baseline).ToString("#.#%")
            }
        }

        $loopView | Sort-Object -Property 'ItemsPerSecond' -Descending
    }
}
#endregion Functions

# Instantiate a bunch of LoopRunners
$loopRunners = [LoopRunner[]] $LoopLogic

# Invoke them all with a progress bar
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

$loopRunners | Get-LoopRunnerView | Format-Table -AutoSize
