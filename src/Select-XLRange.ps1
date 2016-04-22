function Select-XLRange {
[OutputType([XLRange])]
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true)]
    [XLSheet]$Sheet,
    
    [Parameter(Mandatory = $true, Position = 1)]
    [Alias("Row")]
    [int]$FromRow,
    
    [int]$ToRow = $FromRow,
    
    [Parameter(Mandatory = $true, Position = 2)]
    [Alias("Column")]
    [int]$FromColumn,
    
    [int]$ToColumn = $FromColumn,
    
    [string[]]$Header = $null,
    
    [Switch]$HasHeader = $false
)  
begin{
    [OfficeOpenXml.ExcelRange]$range = $Sheet.Worksheet.Cells.Item($FromRow, $FromColumn, $ToRow, $ToColumn)
    
    if ($Header -ne $null -and $Header.Length -ne ($range.Columns)) {
        throw "Header contains $($Header.Length) elements but the selection is for $($range.Columns) columns"
    }

    $xlRange = [XLRange]::new($Sheet.Owner, $range)
    # TODO move these to ctor so they're not publically writeable
    $xlRange.Header = $Header
    $xlRange.HasHeader = $HasHeader.IsPresent
    $PSCmdlet.WriteObject($xlRange, $false)
    
}
end{}
}