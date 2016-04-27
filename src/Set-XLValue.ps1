# this should probably be implemented as a specialization of Set-XLRange
function Set-XLValue {
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

    [Parameter(Mandatory = $true)]
    [object]$Value,

    [ValidateScript({$_ -eq $null -or $_ -is [string] -or $_ -is [XLNumberFormat]})]
    $NumberFormat = $null,

    [Switch]$PassThru = $false
)
begin{}
process {
    [void]$PSBoundParameters.Remove('Value')
    [void]$PSBoundParameters.Remove('NumberFormat')    
    [void]$PSBoundParameters.Remove('PassThru')
    $res = Resolve-XLRange @PSBoundParameters
    $Range = $res.Range
    
    $Range.Range.Value = $Value

    if ($NumberFormat -ne $null) {
        $cellFmt = $Range.Range.Style.Numberformat
        
        if ($NumberFormat -is [string]) {
            $cellFmt.Format = $NumberFormat
        } else {
            # this could very well be wrong
            switch ($NumberFormat.ToString()) {
                "Text" { $cellFmt.Format = "Text" }
                "Date" { $cellFmt.Format = "yyyy-mm-dd" }
                "General" { $cellFmt.Format = "General" }
                "Percent" { $cellFmt.Format = "0.0%"; }
                "DateTime" { $cellFmt.Format = "yyyy-mm-dd h:mm.ss" }
                "Time" { $cellFmt.Format = "h:mm.ss" }
                default { $cellFmt.Format = $NumberFormat }
            }
        }
    }

    $Range.HasHeaders = $false

    if ($PassThru.IsPresent) {
        $PSCmdlet.WriteObject($res.InputObject, $false)
    } else {
        $PSCmdlet.WriteObject($Range, $false)
    }
}
end{}
}
