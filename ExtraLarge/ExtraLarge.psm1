Set-StrictMode -Version 5
Add-Type -Path "$($PSScriptRoot)\EPPlus.dll"

function New-XLFile {
param(
    [Parameter(Mandatory=$true)]
    [string]$Path,
    [Parameter(Mandatory=$true)]
    [object[]]$Sheets,
    [switch]$PassThru = $false
)

    $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    $package = [OfficeOpenXml.ExcelPackage]::new($resolvedPath);
    
    foreach ($sheet in $Sheets) {
        Add-XLSheet -Package $package -Name $sheet
        $package.Workbook.Worksheets.Add($sheet.Name);
    }
    

    
    $package.Save();
    
    if ($PassThru.IsPresent) {
        return $package;
    }
}

function New-XLSheet {
param(
    [string]$Name = "Sheet",
    [switch]$Hidden = $false,
    [object[]]$Tables = @()
)
    New-Object -TypeName PSObject -Property @{
        Name = $Name
        Hidden = [bool]$Hidden
        Tables = $Tables
    }
}

Class XLSheet {
    XLSheet([string]$name) {
        $this.Name = $name;
    }
    XLSheet([OfficeOpenXml.ExcelWorksheet]$sheet) {
        $this.Name = $sheet.Name;
        $this.Worksheet = $sheet;
    }

    [OfficeOpenXml.ExcelWorksheet]$Worksheet
    [string]$Name
}

Function Add-XLSheet {
[OutputType([XLSheet])]
param(
    [Parameter(ParameterSetName = "PackageAndName", Mandatory = $true)]
    [Parameter(ParameterSetName = "PackageAndSheet", Mandatory = $true)]
    [OfficeOpenXml.ExcelPackage]$Package,
    [Parameter(ParameterSetName = "PathAndName", Mandatory = $true)]
    [Parameter(ParameterSetName = "PathAndSheet", Mandatory = $true)]
    [string]$Path,
    [Parameter(ParameterSetName = "PackageAndName", Mandatory = $true)]
    [Parameter(ParameterSetName = "PathAndName", Mandatory = $true)]
    [string]$Name,
    [Parameter(ParameterSetName = "PackageAndSheet", Mandatory = $true)]
    [Parameter(ParameterSetName = "PathAndSheet", Mandatory = $true)]
    [XLSheet]$Sheet,
    [switch]$PassThru = $false
)
    
    if ($PSCmdlet.ParameterSetName.StartsWith("Path")) {
        $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path);
        if (!Test-Path -LiteralPath $resolvedPath) {
            throw "Path not found: '$Path'";
        }
        $Package = [OfficeOpenXml.ExcelPackage]::new($resolvedPath);

    }

    if ($PSCmdlet.ParameterSetName.EndsWith("Name")) {
        $worksheet = $Package.Workbook.Worksheets.Add($Name);
        $Sheet = [XLSheet]::new($worksheet);
    } else {
        $worksheet = $Package.Workbook.Worksheets.Add($Sheet.Name);
        $Sheet.Worksheet = $worksheet;
    }


    if ($PassThru.IsPresent) {
        Write-Output -InputObject $Sheet
    }
}

function New-XLTable {
param(
    [Parameter(Mandatory = $true)]
    [object]$Data,
    [string]$Row = 1,
    [string]$Column = 1,
    [Switch]$Transpose = $false
)  

}

Export-ModuleMember -Function New-XLFile
Export-ModuleMember -Function New-XLSheet
Export-ModuleMember -Function New-XLTable

              