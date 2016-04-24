# TODO this should probably be implemented as a specialization of Set-XLRange
function Set-XLValue {
[OutputType([XLRange])]
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true, ParameterSetName = "Range")]
    [XLRange]$Range,
    
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true, ParameterSetName = "Sheet")]
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true, ParameterSetName = "Named")]    
    [XLSheet]$Sheet,
    
    [Parameter(Mandatory = $true, Position = 1, ParameterSetName = "Sheet")]
    [Alias("Row")]
    [int]$FromRow,
    
    [Parameter(ParameterSetName = "Sheet")]
    [int]$ToRow = $FromRow,
    
    [Parameter(Mandatory = $true, Position = 2, ParameterSetName = "Sheet")]
    [Alias("Column")]
    [int]$FromColumn,
    
    [Parameter(ParameterSetName = "Sheet")]
    [int]$ToColumn = $FromColumn,
    
    [Parameter(Mandatory = $true, ParameterSetName = "Named")]
    [string]$Name,
    
    [Parameter(Mandatory = $true)]
    [object]$Value,
    
    [XLNumberFormat]$NumberFormat = [XLNumberFormat]::General,
    
    [Switch]$PassThru = $false
)  
begin{
}
process {
    if ($PSCmdlet.ParameterSetName -eq 'Named') {
        $excelRange = $Sheet.Worksheet.Workbook.Names[$Name]
        if ($excelRange -eq $null) {
            $excelRange = $Sheet.Worksheet.Names[$Name]
        }
        
        if ($excelRange -eq $null) {
            throw "Name does not exist: $Name"
        }
        $Range = [XLRange]::new($Sheet.Owner, $excelRange)
        
    } elseif ($PSCmdlet.ParameterSetName -eq 'Sheet') {
        [OfficeOpenXml.ExcelRange]$excelRange = $Sheet.Worksheet.Cells.Item($FromRow, $FromColumn, $ToRow, $ToColumn)
        $Range = [XLRange]::new($Sheet.Owner, $excelRange)
    }

    $Range.Range.Value = $Value
    
    $cellFmt = $Range.Range.Style.Numberformat;
    # this could very well be wrong
    switch ([XLNumberFormat]$NumberFormat) {
        "Text" { $cellFmt.Format = "Text" }
        "Date" { $cellFmt.Format = "yyyy-mm-dd" }
        "General" { $cellFmt.Format = "General" }
        "Percent" { $cellFmt.Format = "0.0%"; }
        "DateTime" { $cellFmt.Format = "yyyy-mm-dd h:mm.ss" }
        "Time" { $cellFmt.Format = "h:mm.ss" }
    }
    
    $Range.HasHeaders = $false
    
    if ($PassThru.IsPresent -and $Sheet -ne $null) {
        $PSCmdlet.WriteObject($Sheet, $false)        
    } else {
        $PSCmdlet.WriteObject($Range, $false)
    }
}
end{}
}
