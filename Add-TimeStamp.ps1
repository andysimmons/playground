<#
.SYNOPSIS
    Slaps a timestamp on a string.
#>
function Add-TimeStamp
{
    param (
        [Parameter(ValueFromPipeline, Position = 0)]
        [string[]]
        $Message,

        [string]
        $Format = 'MMM dd hh:mm:ss tt'
    )

    begin { $timeStamp = Get-Date -Format $Format }
    process { $Message.ForEach({ "[$timeStamp] $_" }) }
}
