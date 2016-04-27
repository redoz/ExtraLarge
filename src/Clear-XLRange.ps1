function Clear-XLRange {
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
    
    $res.Range.Range.Clear()
    
    if ($PassThru.IsPresent) {
        $PSCmdlet.WriteObject($res.InputObject, $false);
    } else {
        $PSCmdlet.WriteObject($res.Range, $false);
    }
}
end{}
}
