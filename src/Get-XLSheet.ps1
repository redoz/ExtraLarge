Function Get-XLSheet {
[CmdletBinding()]
[OutputType([XLSheet])]
param(
    [Parameter(ParameterSetName = "FileAndName", Mandatory = $true, ValueFromPipeline = $true)]
    [Parameter(ParameterSetName = "FileAndIndex", Mandatory = $true, ValueFromPipeline = $true)]
    [Parameter(ParameterSetName = "File", Mandatory = $true, ValueFromPipeline = $true)]
    [XLFile]$File,
    [Parameter(ParameterSetName = "PathAndName", Mandatory = $true)]
    [Parameter(ParameterSetName = "PathAndIndex", Mandatory = $true)]
    [Parameter(ParameterSetName = "Path", Mandatory = $true)]
    [string]$Path,
    [Parameter(ParameterSetName = "FileAndName", Mandatory = $true)]
    [Parameter(ParameterSetName = "PathAndName", Mandatory = $true)]
    [string]$Name,
    [Parameter(ParameterSetName = "FileAndIndex", Mandatory = $true)]
    [Parameter(ParameterSetName = "PathAndIndex", Mandatory = $true)]
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
                        
        [XLSheet]::new($File.Package, $worksheet)
    } else {
        $File.Package.Workbook.Worksheets | Foreach-Object -Process { [XLSheet]::new($File.Package, $_) }
    }
}
end {
}
}

