#requires -Module VMware.VimAutomation.Core
using namespace VMware.VimAutomation.ViCore.Impl.V1.VM
[CmdletBinding(SupportsShouldProcess)]
param (
    [string[]]
    $VIServer,

    [string[]]
    $VMName
)

# Depending on PowerCLI config, connecting without the -NotDefault triggers
# some obnoxious warnings. Functionality is the same either way.
try { Connect-VIServer -Server $VIServer -Force -NotDefault -ErrorAction 'Stop' }
catch {
    Write-Error "Error connecting to VI server(s): $($VIServer -join ', ')"
    throw $_.Exception
}

<#
.SYNOPSIS
    Filters out SRM placeholder virtual machines from a collection of VMs

.PARAMETER VirtualMachine
    One or more VM objects to be filtered

.EXAMPLE
    Get-VM -Server 'vc01','vc02' -Name 'srmprotectedvm' | Skip-PlaceHolder

    Retrieves all virtual machines from "vc01" and "vc02" named "srmprotectedvm",
    and filters out any placeholder VMs.
#>
function Skip-PlaceHolder {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [UniversalVirtualMachineImpl[]]
        $VirtualMachine
    )

    process {
        foreach ($vm in $VirtualMachine) {
            if ($_.ExtensionData.Summary.Config.ManagedBy.Type -eq 'placeholderVm') {
                Write-Verbose "Skipping '$vm' placeholder (UID: $($vm.Uid))"
            }
            else { $vm }
        }
    }
}

# get any primary VMs that match the name(s) we're looking for
$primaryVm = Get-VM -Name $VMName -Server $VIServer | Skip-PlaceHolder

# and snapshot them
foreach ($vm in $primaryVm) {
    if ($PSCmdlet.ShouldProcess($vm, 'create snapshot')) {
        New-Snapshot -Name "Andy Test" -Description "Not important" -VM $vm
    }
}
