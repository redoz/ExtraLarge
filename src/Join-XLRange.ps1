function Join-XLRange {
[OutputType([XLRange], ParameterSetName = "Range")]
[OutputType([XLSheet], ParameterSetName = "Sheet")]
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true, ParameterSetName = "Range")]
    [XLRange]$Range,
    
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true, ParameterSetName = "Sheet")]
    [XLSheet]$Sheet,
    
    [Parameter(Mandatory = $true, Position = 1, ParameterSetName = "Sheet")]
    [Alias("Row")]
    [int]$FromRow,
    
    [Parameter(Mandatory = $true, ParameterSetName = "Sheet")]
    [int]$ToRow,
    
    [Parameter(Mandatory = $true, Position = 2, ParameterSetName = "Sheet")]
    [Alias("Column")]
    [int]$FromColumn,
    
    [Parameter(Mandatory = $true, ParameterSetName = "Sheet")]
    [int]$ToColumn,
    
    [Switch]$PassThru = $false
)  
begin{
    [OfficeOpenXml.ExcelRange]$excelRange = $null 
    if ($PSCmdlet.ParameterSetName -eq "Sheet") {
        $excelRange  = $Sheet.Worksheet.Cells.Item($FromRow, $FromColumn, $ToRow, $ToColumn)  
    } else {
        $excelRange = $Range.Range
    }
    
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
        if ($PSCmdlet.ParameterSetName -eq "Sheet") {
            $PSCmdlet.WriteObject($Sheet, $false);
        } else {
            $PSCmdlet.WriteObject($Range, $false);
        }
    }
}
end{}
}