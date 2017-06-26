<#
.SYNOPSIS
    Joins multi-valued properties on an object.

.DESCRIPTION
    Joins multi-valued properties on an object, making it easier to
    export to "flat" formats (e.g. CSV) without losing information.

.PARAMETER InputObject
    Specifies an object to be flattened.

.PARAMETER Delimiter
    Specifies one or more characters placed between the concatenated strings.

.EXAMPLE
    Get-ADUser dsmith -Properties * | Join-MultiValuedProperties | Export-Csv -NoType -Path C:\dsmith.csv
#>
function Join-MultiValuedProperties
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

			$csvFriendlyObject = New-Object PSObject

			# Loop through each property on the input object
			foreach ($property in $_.PSObject.Properties)
			{
                $propName  = $property.Name
                $propValue = $property.Value

				# If it has a value, see if it's a collection, and if it is, join
				# the elements together in a single string.
				if ($propValue -and $propValue.GetType().GetInterface('ICollection'))
				{
					$propValue = $propValue -Join $Delimiter
				}

				$addMemberParams = @{
					InputObject = $csvFriendlyObject
					MemberType  = 'NoteProperty'
					Name        = $propName
					Value       = $propValue
				}
				Add-Member @addMemberParams
			}

			$csvFriendlyObject
		}
	}
}
