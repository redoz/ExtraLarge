Function Add-XLSheet {
[CmdletBinding()]
[OutputType([object])]
param(
    [Parameter(ParameterSetName = "Package", Mandatory = $true, ValueFromPipeline = $true)]
    [OfficeOpenXml.ExcelPackage]$Package,
    [Parameter(ParameterSetName = "Path", Mandatory = $true)]
    [string]$Path,
    [string]$Name,
    [string]$Hidden = $false,
    [switch]$Force = $false,
    [switch]$PassThru = $false,
    [scriptblock]$With = $null
)
begin{}
process{
    if ($PSCmdlet.ParameterSetName -eq "Path") {
        $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path);
        if (-not (Test-Path -LiteralPath $resolvedPath)) {
            throw "Path not found: '$Path'";
        }
        $Package = [OfficeOpenXml.ExcelPackage]::new($resolvedPath);
    }

    if ($Package.Workbook.Worksheets[$Name] -ne $null) {
        if ($Force.IsPresent) {
            Write-Verbose -Message "Deleting worksheet: '${Name}'"
            $Package.Workbook.Worksheets.Delete($Name);
        }
    }

    $worksheet = $Package.Workbook.Worksheets.Add($Name);

    if ($With -ne $null) {
        $worksheet | ForEach-Object -Process $With
    }

    if ($PassThru.IsPresent) {
        Write-Output -InputObject $Package
    } else {
        Write-Output -InputObject $worksheet
    }
}
end {
}
}
