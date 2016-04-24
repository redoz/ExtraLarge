Function Add-XLSheet {
[CmdletBinding()]
[OutputType([XLSheet])]
param(
    [Alias("Package")]
    [Parameter(ParameterSetName = "File", Mandatory = $true, ValueFromPipeline = $true)]
    [XLFile]$File,
    [Parameter(ParameterSetName = "Path", Mandatory = $true)]
    [string]$Path,
    [string]$Name,
    [string]$Hidden = $false,
    [switch]$Force = $false,
    [switch]$PassThru = $false,
    [scriptblock]$With = $null,
    [switch]$Save = $false
)
begin{}
process{
    $package = $null
    if ($PSCmdlet.ParameterSetName -eq "Path") {
        $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path);
        if (-not (Test-Path -LiteralPath $resolvedPath)) {
            throw "Path not found: '$Path'";
        }
        $package = [OfficeOpenXml.ExcelPackage]::new($resolvedPath)
    } else {
        $package = $File.Package
    }

    if ($package.Workbook.Worksheets[$Name] -ne $null) {
        if ($Force.IsPresent) {
            Write-Verbose -Message "Deleting worksheet: '${Name}'"
            $package.Workbook.Worksheets.Delete($Name)
        }
    }

    $worksheet = [XLSheet]::new($package, $package.Workbook.Worksheets.Add($Name))
    
    if ($Save.IsPresent) {
        $package.Save()
    }

    if ($With -ne $null) {
        $worksheet | ForEach-Object -Process $With
    }

    if ($PassThru.IsPresent) {
        if ($File -eq $null) {
            $File = [XLFile]::new($package)
        }
        
        Write-Output -InputObject $File
    } else {
        Write-Output -InputObject $worksheet
    }
}
end {
}
}
