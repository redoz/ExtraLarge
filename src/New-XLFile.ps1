function New-XLFile {
[OutputType([XLFile])]
param(
    [Parameter(Position = 0, Mandatory=$true)]
    [string]$Path,
    [switch]$NoDefaultSheet = $false,
    [switch]$PassThru = $false,
    [switch]$Force = $false,
    # TODO make -With work
    [scriptblock]$With = $null
)
begin {
    $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    if (Test-Path -Path $resolvedPath -PathType Leaf) {
        if ($Force.IsPresent) {
            Remove-Item -Path $resolvedPath -Force
        } else {
            throw "File exists: '${resolvedPath}', use -Force to overwrite.";
        }
    }
    $package = [OfficeOpenXml.ExcelPackage]::new($resolvedPath);
    
    if ($PassThru.IsPresent) {
        return [XLFile]::new($package);
    }
}
process {
}
end {

    if ($package.Workbook.Worksheets.Count -eq 0) {
        if (-not $NoDefaultSheet.IsPresent) {
            Write-Verbose -Message "Creating default worksheet: 'Default'"
            [void]$package.Workbook.Worksheets.Add("Default");
        }
    } else { 
        $package.Save();
    }
}
}