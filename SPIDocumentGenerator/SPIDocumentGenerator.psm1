
function New-SPInventoryDocument
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$InputFilePath,
		[Parameter(Mandatory = $true)]
		[string]$OutputFilePath
	)
	
	$Word = New-Object -ComObject Word.Application
	$Word.Visible = $true
	$Document = $Word.Documents.Add()
	$Selection = $Word.Selection
	
	$object = Import-Clixml -Path $InputFilePath
	
	foreach ($property in $object.PSObject.Properties)
	{
		New-WordHeading -Heading $property.Value.Title
		
        # bug...
        $columns = $null
        $rows = $null

		if ($property.Value.Payload.GetType().Name -eq 'PSCustomObject')
		{
			$columns = 2
			$rows = ($property.Value.Payload.PSObject.Properties | Measure-Object).Count
			New-WordTable -WordObject $Word -Object $property.Value.Payload -Columns $columns -Rows $rows -AsList
		}
		elseif ($property.Value.Payload.GetType().Name -eq 'ArrayList')
		{
			$columns = ($property.Value.Payload[0].PSObject.Properties | Measure-Object).Count
			$rows = $property.Value.Payload.Count + 1
			New-WordTable -WordObject $Word -Object $property.Value.Payload -Columns $columns -Rows $rows -AsTable
		}
		
	}
}

Export-ModuleMember -Function New-SPInventoryDocument

Function New-WordTable
{
	[cmdletbinding(
				   DefaultParameterSetName = 'Table'
				   )]
	Param (
		[parameter()]
		[object]$WordObject,
		[parameter()]
		[object]$Object,
		[parameter()]
		[int]$Columns,
		[parameter()]
		[int]$Rows,
		[parameter(ParameterSetName = 'Table')]
		[switch]$AsTable,
		[parameter(ParameterSetName = 'List')]
		[switch]$AsList,
		[parameter()]
		[string]$TableStyle,
		[parameter()]
		[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]$TableBehavior = 'wdWord9TableBehavior',
		[parameter()]
		[Microsoft.Office.Interop.Word.WdAutoFitBehavior]$AutoFitBehavior = 'wdAutoFitWindow'
	)
	#Specifying 0 index ensures we get accurate data from a single object
	$Properties = $Object[0].psobject.properties.name
	$Range = @($WordObject.Paragraphs)[-1].Range
	$Table = $WordObject.Selection.Tables.add(
		$WordObject.Selection.Range, $Rows, $Columns, $TableBehavior, $AutoFitBehavior)
	Switch ($PSCmdlet.ParameterSetName)
	{
		'Table' {
			If (-NOT $PSBoundParameters.ContainsKey('TableStyle'))
			{
				$Table.Style = "Medium Shading 1 - Accent 1"
			}
			$c = 1
			$r = 1
			#Build header
			$Properties | ForEach-Object {
				Write-Verbose "Adding $($_)"
				$Table.cell($r, $c).range.Bold = 1
				$Table.cell($r, $c).range.text = $_
				$c++
			}
			$c = 1
			#Add Data
			For ($i = 0; $i -lt (($Object | Measure-Object).Count); $i++)
			{
				$Properties | ForEach-Object {
					$Table.cell(($i + 2), $c).range.Bold = 0
					$Table.cell(($i + 2), $c).range.text = $Object[$i].$_
					$c++
				}
				$c = 1
			}
		}
		'List' {
			If (-NOT $PSBoundParameters.ContainsKey('TableStyle'))
			{
				$Table.Style = "Light Shading - Accent 1"
			}
			$c = 1
			$r = 1
			$Properties | ForEach-Object {
				$Table.cell($r, $c).range.Bold = 1
				$Table.cell($r, $c).range.text = $_
				$c++
				$Table.cell($r, $c).range.Bold = 0
				$Table.cell($r, $c).range.text = $Object.$_
				$c--
				$r++
			}
		}
		
	}
	$WordObject.Selection.Start = $Document.Content.End
	$WordObject.Selection.TypeParagraph()
}

function New-WordHeading
{
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$Heading
	)
	
	$Selection.Style = "Heading 1"
	$Selection.TypeText($Heading)
	$Selection.TypeParagraph()
}
