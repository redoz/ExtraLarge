function Select-XLRange {
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
    
    [string[]]$Headers = $null,
    
    [Switch]$HasHeaders = $false
)  
begin{
}
process {
    [void]$PSBoundParameters.Remove('Headers')
    [void]$PSBoundParameters.Remove('HasHeaders')    
    $res = Resolve-XLRange @PSBoundParameters

    [OfficeOpenXml.ExcelRange]$excelRange = $res.Range.Range    
    
    if ($Headers -ne $null -and $Headers.Length -ne $excelRange.Columns) {
        throw "Header contains $($Header.Length) elements but the selection is for $($excelRange.Columns) columns"
    }

    $xlRange = [XLRange]::new($Sheet.Owner, $excelRange)
    # TODO move these to ctor so they're not publically writeable
    $xlRange.Headers = $Headers
    $xlRange.HasHeaders = $HasHeaders.IsPresent
    $PSCmdlet.WriteObject($xlRange, $false)
    
}
end{}
}
