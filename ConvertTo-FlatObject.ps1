<#
.NOTES
    Created on:   6/26/2017
    Created by:   Andy Simmons
    Organization: St. Luke's Health System
    Filename:     ConvertTo-FlatObject.ps1

.SYNOPSIS
    Joins multi-valued properties on an object.

.DESCRIPTION
    Concatenates an object's multi-valued properties into strings, making it easier to
    export to "flat" formats (e.g. CSV) with less information loss.

    Note: Some information may still be lost on elements of a collection that are
    collections themselves, and/or elements with a generic .ToString() implementation.

.PARAMETER InputObject
    Specifies an object to be flattened.

.PARAMETER Delimiter
    Specifies one or more characters placed between the concatenated strings.

.EXAMPLE
    Get-Service | ConvertTo-FlatObject | Export-Csv -NoTypeInformation -Path C:\services.csv

    Pulls a list of local services, joins multi-valued properties (e.g. DependentServices,
    ServicesDependedOn), and exports the information to a CSV file.
#>
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
