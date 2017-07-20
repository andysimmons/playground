param (
    [IO.FileInfo]
    $ReportPath = "C:\temp\IT_VDISummary$(Get-Date -Format 'yyyy-MM-dd').csv"
)

function Get-VDISummary
{
    [CmdletBinding()]
    param 
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [string[]]
        $Users,

        [regex]
        $RolePattern = 'CN=(MVCoder|OfficeStaff|RehabTherapist|DefaultVDI|(MV|McCall|Elmore)?(Nurse|Physician)|MSTI(Nurse|Pharmacist)|(EVS|Transport|Win7PM-SL[BT])-U)_GG_TM',
        
        [regex]
        $DefaultRolePattern = 'CN=DefaultVDI_GG_TM',

        [regex]
        $SitePattern = 'CN=VDI-SL[BT]User_GG_CX',
    
        [regex]
        $SSOPattern  = 'CN=SSOUsers-U_GG_AP',
    
        [regex]
        $CAGPattern  = 'CN=CAGAccess_GG_CX'
    )
    
    process
    {

        foreach ($user in $Users)
        {		
            try
            {
                $adUser = (Get-ADUser -Identity $user -Properties MemberOf -ErrorAction Stop)
            }
            catch 
            { 
                $adUser = $null 
                Write-Error 'Failed to pull membership for $user! Skipping.'
            }

            if ($adUser) 
            {
                # pull role groups
                $roles = $adUser.MemberOf -match $RolePattern
                $hasRole = [bool] $roles
                $hasAssignedRole = ($roles.Count -eq 1) -and ($roles -notmatch $DefaultRolePattern)

                # check StoreFront site affinity (verify they're in exactly one group)
                $sfSite = @($adUser.MemberOf -match $SitePattern)
                if ($sfSite.Count -ne 1) 
                { 
                    $sfSite = "NONE" 
                }
                else 
                {
                    # Set the site to either "SLB" or "SLT" for clarity in the report (just grabs it from the group name)
                    $sfSite = $sfSite[0].Substring(7,3)
                }
            
                # infer basic VDI functionality
                $hasVDI = (($sfSite -ne 'NONE') -and $hasRole)

                # check single sign on
                $hasSSO = [bool] ($adUser.MemberOf -match $SSOPattern)

                # check Citrix Acccess Gateway
                $hasExternalAccess = [bool] ($adUser.MemberOf -match $CAGPattern)

                # Return a custom object describing the user and their VDI access
                [pscustomobject] [ordered] @{
                    User              = $adUser.Name
                    SamAccountName    = $adUser.SamAccountName
                    HasVDI            = $hasVDI
                    PreferredSite     = $sfSite
                    HasAssignedRole   = $hasAssignedRole              
                    HasSSO            = $hasSSO
                    HasExternalAccess = $hasExternalAccess
                    RoleGroups        = $roles
                }
            }
        
            else 
            {
                $errText = "AD_LOOKUP_ERR"

                [pscustomobject] [ordered] @{
                    User              = $user
                    SamAccountName    = $errText
                    HasVDI            = $errText
                    PreferredSite     = $errText
                    HasAssignedRole   = $errText
                    HasSSO            = $errText
                    HasExternalAccess = $errText
                    RoleGroups        = $errText
                }
            }
        }
    }
}

<#
.SYNOPSIS
    Returns all users from a group (searching recursively)

.PARAMETER GroupName
    Name of the AD group to search.

.PARAMETER CurrDepth
    Don't use this when invoking. This tracks the current depth since this function
    calls itself to dive into nested groups.

.PARAMETER MaxDepth
    Maximum recursion depth.
#>
function Get-RecursiveGroupUser
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, HelpMessage = 'AD Group name to search')]
        [string]
        $GroupName,
        
        [int]
        $CurrDepth = 0,
        
        [int]
        $MaxDepth = 5
    )		
    
    # Avoiding an infinite recursion loop...
    if ($CurrDepth -le $MaxDepth)
    {
    
        # Create collection of AD objects with direct membership in $GroupName
        $members = @((Get-ADGroup $GroupName -Properties Members).Members | .{ process {Get-ADObject -Identity $_ } })
        
        # Loop through those objects, returning users as they're found, and recursively searching groups
        # for additional users if those are found.
        foreach ($member in $members)
        {
            if ($member.ObjectClass -eq 'user')
            {
                $member
            }
            elseif ($member.ObjectClass -eq 'group')
            {
                Get-RecursiveGroupUser -GroupName $member.Name -CurrDepth $($CurrDepth + 1)
            }
            else
            {
                Write-Warning "Unexpected object class $($member.ObjectClass). Skipping $member."
            }
        }
    }
    else
    {
        Write-Warning "Max recursion depth limit ($MaxDepth) reached! Abandoning recursion on $member."
    }
}

function ConvertTo-FlatObject
{
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [object[]] $InputObject,
        
        [String] 
        $Delimiter = "`n"   
    )
  
    process 
    {
        $InputObject | ForEach-Object {

            $flatObject = New-Object PSObject

            # Loop through each property on the input object
            foreach ($property in $_.PSObject.Properties)
            {
                # If it's a collection, join everything into a string.
                if ($property.TypeNameOfValue -match '\[\]$')
                {
                    $flatValue = $property.Value -Join $Delimiter
                }
                else { $flatValue = $property.Value }

                $addMemberParams = @{
                    InputObject = $flatObject
                    MemberType  = 'NoteProperty'
                    Name        = $property.Name
                    Value       = $flatValue
                }
                Add-Member @addMemberParams
            }

            $flatObject
        }
    }
}

$itUsers = Get-RecursiveGroupUser -GroupName 'IT_GG_TM'
$vdiSummary = $itUsers | Get-VDISummary | ConvertTo-FlatObject
$vdiSummary | Export-Csv -Path $ReportPath -NoTypeInformation
Invoke-Item $ReportPath
