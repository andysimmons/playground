<#
.NOTES
Name:     Get-VAppResourceInfo.ps1
Author:   Andy Simmons
Date:     02/12/2013
Last Rev: 12/29/2017

.SYNOPSIS
Analyze vApps and determine if the shares are configured appropriately.

.DESCRIPTION
Quick script to analyze the vApps and determine if the memory/CPU shares are correctly
configured to ensure consistent treatment of VMs with respect to the cluster root pool.

This probably won't work if you nest resource pools inside non-root resource pools.

.EXAMPLE
.\Get-VAppResourceInfo.ps1 -VIServer 'vcenter01','vcenter02'

Analyzes vApp shares on the 'vcenter01' and 'vcenter02' vCenter servers. 
#>
[CmdletBinding()]
param 
(
    [Parameter(Mandatory, Position = 0)]
    [string[]]
    $VIServer
)

# built-in share multipliers
$cpuLow	 = 500
$cpuNorm = 1000
$cpuHigh = 2000
$memLow	 = 5
$memNorm = 10
$memHigh = 20

# initialization
try   { Add-PSSnapin VMware.VimAutomation.Core -ErrorAction 'Stop' }
catch { throw "Power CLI required. Aborting." }

$serverList = ($VIServer -join ', ').ToUpper()

$alreadyConnected = (Test-Path vis:\) -and ((Get-Childitem vis:\).Where({$_.AsDefaultFormat.Name -eq $VIServer}))
if (-not $alreadyConnected)
{
    Write-Output "Connecting to $serverList"
    try 	{ Connect-VIServer $VIServer -ErrorAction 'Stop'  }
    catch 	{ throw "Couldn't connect to $serverList.`n$($_.Exception.Message)" }
}
    
# find vApps
try   { $vApps = @(Get-VApp -Server $VIServer -ErrorAction 'Stop') }
catch { throw "Couldn't retrieve vApps from $serverList.`n$($_.Exception.Message)" }

if (!$vApps)
{
    "No vApps found. Nothing to do."
    exit
}

# analyze vApp shares
$i = 0
$summary = foreach ($vApp in $vApps)
{
    $i++
    $wpParams = @{
        Id              = 1
        Activity        = 'Checking vApp shares: {0}/{1}' -f $i, $vApps.Count
        Status          = $vApp.Name
        PercentComplete = 100 * $i / $vApps.Count
    }
    Write-Progress @wpParams
    
    # child VM share analysis
    $j = 0
    $idealCpuShares = 0
    $idealMemShares = 0

    $VMs = @(Get-VM -Location $vApp)
    foreach ($vm in $VMs)
    {
        $j++
        $wpParams = @{
            Id              = 2
            ParentId        = 1
            Activity        = 'Checking VM shares: {0}/{1}' -f $j, $VMs.Count
            Status          = $vm.Name
            PercentComplete = 100 * $j / $VMs.Count 
        }
        Write-Progress @wpParams
        
        switch ($vm.ExtensionData.ResourceConfig.CpuAllocation.Shares.Level)
        {
            "Low"    { $vmCpuShares = [int] ($vm.NumCpu * $cpuLow) }	
            "Normal" { $vmCpuShares = [int] ($vm.NumCpu * $cpuNorm) }
            "High"   { $vmCpuShares = [int] ($vm.NumCpu * $cpuHigh) }
            default  { $vmCpuShares = $vm.ExtensionData.ResourceConfig.CpuAllocation.Shares.Shares }
        }
        $idealCpuShares += $vmCpuShares
    
        switch ($vm.ExtensionData.ResourceConfig.CpuAllocation.Shares.Level)
        {
            "Low"    { $vmMemShares = [int] ($vm.MemoryMB * $memLow) }	
            "Normal" { $vmMemShares = [int] ($vm.MemoryMB * $memNorm) }
            "High"   { $vmMemShares = [int] ($vm.MemoryMB * $memHigh) }
            default  { $vmMemShares = $vm.ExtensionData.ResourceConfig.MemoryAllocation.Shares.Shares }
        }
        $idealMemShares += $vmMemShares
    }
    
    # vApp shares OK?
    $cpuOk = $vApp.NumCpuShares -eq $idealCpuShares
    $memOk = $vApp.NumMemShares -eq $idealMemShares
    $fixMe = !($cpuOk -and $memOk)

    # check each parent til we hit the cluster root
    $container = $vApp.Parent
    while ($container.Parent) { $container = $container.Parent }
    
    # parse vCenter server name from client info
    $vCenter = [regex]::match($vApp.Client.ServerUri, '(?<=[0-9]+@).*').Value.ToUpper()
    
    # return vApp summary object
    [PSCustomObject] [ordered] @{
        Name           = $vApp.Name
        vCenter        = $vCenter
        Cluster        = $container.Name
        FixMe          = $fixMe
        vAppCpuShares  = $vApp.NumCpuShares
        IdealCpuShares = $idealCpuShares
        vAppMemShares  = $vApp.NumMemShares
        IdealMemShares = $idealMemShares
        CpuOk          = $cpuOk
        MemOk          = $memOk
    }
}

$summary | Sort-Object -Descending -Property FixMe, Cluster, Name | Out-GridView
