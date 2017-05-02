#requires -version 5.0
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [object[]] $Collection,

    [scriptblock[]] $LoopLogic = @(
        { $Collection | .{ process { $_ } } }, 
        { $Collection | ForEach-Object { $_ } },
        { ForEach-Object -InputObject $Collection { $_ } },
        { foreach ($element in $Collection) { $element } }
    )
)

# Invokes loops and watches performance.
class LoopRunner 
{
    # Properties
    [scriptblock] $ScriptBlock
    [timespan]    $ProcessingTime
    [int]         $ItemCount
    [double]      $ItemsPerSecond
    [array]       $Result
    
    # Constructors
    LoopRunner()
    { 
        $this.Initialize({}) 
    }
    
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
        
        # Invoke the scriptblock and capture interesting metrics
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

    [string] GetRelativeSpeed ([double] $BaseLineItemsPerSec)
    {
        if ($this.ItemsPerSecond)
        {
            $relativeSpeed = $this.ItemsPerSecond / $BaseLineItemsPerSec
            return $relativeSpeed.ToString("#%")
        }
        else { return "Unknown" }
    }
    
    hidden [int] getItemCount ()
    {
        if ($this.Result) { return $this.Result.Length }
        else              { return -1 }
    }
    
    hidden [double] getItemsPerSecond ([int]$Precision)
    {
        if ($this.ItemCount -and $this.ProcessingTime.TotalSeconds)
        {
            return [math]::Round(($this.ItemCount / $this.ProcessingTime.TotalSeconds), $Precision)
        }
        else { return -1 }
    }
}

# Re-cast the scriptblocks as LoopRunners and invoke them all against our collection.
$loopRunners = [LoopRunner[]]$LoopLogic

for ($i = 0; $i -lt $loopRunners.Length; $i++) 
{
    $writeProgressParams = @{
        Id       = 10
        Activity = "Looping through $($Collection.Length) items"
        Status   = "Method $i of $($loopRunners.Length): { $($loopRunners[$i].ScriptBlock) }"
        PercentComplete = 100 * $i / $loopRunners.Length
    }
    Write-Progress @writeProgressParams
    
    $loopRunners[$i].Invoke()
}

# Summarize results
$topSpeed = ($loopRunners | Measure-Object -Property 'ItemsPerSecond' -Maximum).Maximum

$loopRunners | Select-Object -Property ScriptBlock,
                                        @{ 
                                            n = 'ProcessingTime (sec)'
                                            e = { [math]::Round($_.ProcessingTime.TotalSeconds, 3) }
                                        }, 
                                        ItemCount, 
                                        ItemsPerSecond, 
                                        @{
                                            n = 'RelativeSpeed' 
                                            e = { $_.GetRelativeSpeed($topSpeed) }
                                        } | 
    Sort-Object -Property 'ItemsPerSecond' -Descending | 
    Format-Table -AutoSize
