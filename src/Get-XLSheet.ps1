Function Get-XLSheet {
[CmdletBinding()]
[OutputType([XLSheet])]
param(
    [Parameter(ParameterSetName = "FileAndName", Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
    [Parameter(ParameterSetName = "FileAndIndex", Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
    [Parameter(ParameterSetName = "File", Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
    [XLFile]$File,
    [Parameter(ParameterSetName = "PathAndName", Mandatory = $true, Position = 0)]
    [Parameter(ParameterSetName = "PathAndIndex", Mandatory = $true, Position = 0)]
    [Parameter(ParameterSetName = "Path", Mandatory = $true, Position = 0)]
    [string]$Path,
    [Parameter(ParameterSetName = "FileAndName", Mandatory = $true, Position = 1)]
    [Parameter(ParameterSetName = "PathAndName", Mandatory = $true, Position = 1)]
    [string]$Name,
    [Parameter(ParameterSetName = "FileAndIndex", Mandatory = $true, Position = 1)]
    [Parameter(ParameterSetName = "PathAndIndex", Mandatory = $true, Position = 1)]
    [int]$Index
    
)
begin{}
process{
    if ($PSCmdlet.ParameterSetName.StartsWith("Path")) {
        $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path);
        if (-not (Test-Path -LiteralPath $resolvedPath)) {
            throw "Path not found: '$Path'";
        }
        $File = [OfficeOpenXml.ExcelPackage]::new($resolvedPath);
    }

    if ($PSCmdlet.ParameterSetName.EndsWith("Name")) {
        $worksheet = $File.Package.Workbook.Worksheets[$Name];
        
        if ($worksheet -eq $null) {
            throw "Sheet '$Name' not found."
        }
        
        [XLSheet]::new($File.Package, $worksheet)
    } elseif ($PSCmdlet.ParameterSetName.EndsWith("Index")) {
        $worksheet = $File.Package.Workbook.Worksheets.Item($Index)
        if ($worksheet -ne $null) {
            [XLSheet]::new($File.Package, $worksheet)
        }
    } else {
        $File.Package.Workbook.Worksheets | Foreach-Object -Process { [XLSheet]::new($File.Package, $_) }
    }
}
end {
}
}

