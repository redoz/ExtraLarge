Function Get-XLFile {
[CmdletBinding()]
[OutputType([XLFile])]
param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [string]$Path
)
begin{}
process{
    $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path);
    if (-not (Test-Path -LiteralPath $resolvedPath)) {
        throw "File not found: '$Path'";
    }
    $Package = [OfficeOpenXml.ExcelPackage]::new($resolvedPath);
    
    Write-Output -InputObject ([XLFile]::new($Package))
}
end {
}
}
