function Join-XLRange {
[OutputType([XLRange])]
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true, ParameterSetName = "Range")]
    [XLRange]$Range,

    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true, ParameterSetName = "SheetAndRC")]
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true, ParameterSetName = "SheetAndName")]
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true, ParameterSetName = "SheetAndAddress")]    
    [XLSheet]$Sheet,
    
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true, ParameterSetName = "FileAndName")]
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true, ParameterSetName = "FileAndAddress")]
    [XLFile]$File,

    [Parameter(Mandatory = $true, Position = 1, ParameterSetName = "SheetAndRC")]
    [Alias("Row")]
    [int]$FromRow,

    [Parameter(ParameterSetName = "SheetAndRC")]
    [int]$ToRow = $FromRow,

    [Parameter(Mandatory = $true, Position = 2, ParameterSetName = "SheetAndRC")]
    [Alias("Column")]
    [int]$FromColumn,

    [Parameter(ParameterSetName = "SheetAndRC")]
    [int]$ToColumn = $FromColumn,
    
    [Parameter(ParameterSetName = "SheetAndName")]
    [XLScope]$Scope = [XLScope]::Any,

    [Parameter(Mandatory = $true, ParameterSetName = "FileAndName")]
    [Parameter(Mandatory = $true, ParameterSetName = "SheetAndName")]
    [string]$Name,
    
    [Parameter(Mandatory = $true, ParameterSetName = "FileAndAddress")]
    [Parameter(Mandatory = $true, ParameterSetName = "SheetAndAddress")]
    [string]$Address,
    
    [Switch]$PassThru = $false
)  
begin{
}
process {
    [void]$PSBoundParameters.Remove('PassThru')
    $res = Resolve-XLRange @PSBoundParameters

    [OfficeOpenXml.ExcelRange]$excelRange = $res.Range.Range
    
    foreach ($table in $excelRange.Worksheet.Tables) {
        if (($excelRange.Start.Column -ge $table.Address.Start.Column -and $excelRange.Start.Column -le $table.Address.End.Column) -or
            ($excelRange.Start.Row -ge $table.Address.Start.Row -and $excelRange.Start.Row -le $table.Address.End.Row) -or 
           (($excelRange.End.Column -ge $table.Address.Start.Column -and $excelRange.End.Column -le $table.Address.End.Column) -or
            ($excelRange.End.Row -ge $table.Address.Start.Row -and $excelRange.End.Row -le $table.Address.End.Row))) {
            
            throw "Range overlaps with existing table."
        }        
    }
    $excelRange.Merge = $true
    
    if ($PassThru.IsPresent) {
        $PSCmdlet.WriteObject($res.InputObject, $false);
    } else {
        $PSCmdlet.WriteObject($res.Range, $false);
    }
}
end{}
}
