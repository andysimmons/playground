#requires -Module VMware.VimAutomation.Core
using namespace VMware.VimAutomation.ViCore.Impl.V1.VM
[CmdletBinding(SupportsShouldProcess)]
param (
    [string[]]
    $VIServer,

    [string[]]
    $VMName
)

try { Connect-VIServer -Server $VIServer -Force -NotDefault -ErrorAction 'Stop' }
catch {
    Write-Error "Error connecting to VI server(s): $($VIServer -join ', ')"
    throw $_.Exception
}

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

$primaryVm = Get-VM -Name $VMName -Server $VIServer | Skip-PlaceHolder

foreach ($vm in $primaryVm) {
    if ($PSCmdlet.ShouldProcess($vm, 'create snapshot')) {
        New-Snapshot -Name "$[Snapshot Name]" -Description "$[Snapshot Description]" $[SnapMemory_Param] -VM $vm
    }
}
