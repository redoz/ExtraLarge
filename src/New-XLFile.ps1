function New-XLFile {
[OutputType([XLFile])]
param(
    [Parameter(Position = 0, Mandatory=$true, ValueFromPipeline = $true)]
    [string]$Path,
    [switch]$NoDefaultSheet = $false,
    [switch]$PassThru = $false,
    [switch]$Force = $false,
    [scriptblock]$With = $null
)
begin {
    $createdFiles = [System.Collections.Generic.List[XLFile]]::new()
}
process {
    $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    if (Test-Path -Path $resolvedPath -PathType Leaf) {
        if ($Force.IsPresent) {
            Remove-Item -Path $resolvedPath -Force
        } else {
            throw "File exists: '${resolvedPath}', use -Force to overwrite.";
        }
    }
    
    try {
        $package = [OfficeOpenXml.ExcelPackage]::new($resolvedPath);
    } catch {
        throw;
    }
    
    $xlFile = [XLFile]::new($package)
    $createdFiles.Add($xlFile)
    
    if ($With -ne $null) {
        $xlFile | ForEach-Object -Process $With
    }
        
    if ($PassThru.IsPresent) {
        $PSCmdlet.WriteObject($xlFile)
    }    
}
end {
    foreach ($file in $createdFiles) {
        if ($file.Package.Workbook.Worksheets.Count -eq 0) {
            if (-not $NoDefaultSheet.IsPresent) {
                Write-Verbose -Message "Creating default worksheet: 'Default'"
                [void]$file.Package.Workbook.Worksheets.Add("Default");
                $file.Save();
            }
        } else { 
            $file.Save();
        }
    }
}
}