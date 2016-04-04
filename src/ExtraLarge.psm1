Set-StrictMode -Version 5
Add-Type -Path "$($PSScriptRoot)\EPPlus.dll"

function New-XLFile {
param(
    [Parameter(Position = 0, Mandatory=$true)]
    [string]$Path,
    [switch]$PassThru = $false,
    [switch]$Force = $false,
    # TODO make -With work
    [scriptblock]$With = $null
)
begin {
    $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    if (Test-Path -Path $resolvedPath) {
        if ($Force.IsPresent) {
            Remove-Item -Path $resolvedPath -Force
        } else {
            throw "File exists: '${resolvedPath}', use -Force to overwrite.";
        }
    }
    $package = [OfficeOpenXml.ExcelPackage]::new($resolvedPath);
    
    if ($PassThru.IsPresent) {
        return $package;
    }
}
process {
}
end {

    if ($package.Workbook.Worksheets.Count -eq 0) {
        Write-Verbose -Message "Creating default worksheet: 'Default'"
        [void]$package.Workbook.Worksheets.Add("Default");
    }

    $package.Save();
}
 
}



Function Add-XLSheet {
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
    # TODO make -With work
    [scriptblock]$With = $null
)
begin{}
process{
    if ($PSCmdlet.ParameterSetName -eq "Path") {
        $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path);
        if (!Test-Path -LiteralPath $resolvedPath) {
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

    if ($PassThru.IsPresent) {
        Write-Output -InputObject $Package
    } else {
        Write-Output -InputObject $worksheet
    }
}
end {}
}

function Get-Value($Datum, $ColumnDefinition) {
   $value = $Datum | ForEach-Object -Process $ColumnDefinition.Expression;
   if ($value -eq $null -and $ColumnDefinition.ContainsKey('Default')) {
        $value = $ColumnDefinition.Default;
   }
   $value;
}

function Get-Columns($Datum , $ColumnDefinitions) {
    # normalize columns
    if ($ColumnDefinitions -eq $null) {
        $ColumnDefinitions = Get-Member -InputObject $Datum -MemberType Properties | 
            ForEach-Object -Process { @{Name = $_.Name; Property = $_.Name; Type = [Type]$_.TypeName} }
    }

    foreach ($col in $ColumnDefinitions) {
        if (-not $col.ContainsKey('Expression')) {
            $propertyName = $col.Property;
            $col.Expression = { $_.$propertyName }.GetNewClosure();
        }

        if (-not $col.ContainsKey('Default')) {
            $col.Default = $null;
        }

        if (-not $col.ContainsKey('Type')) {
            $value = Get-Value -Datum $Datum -Column $col;
            if ($value -ne $null) {
                $col.Type = $value.GetType();
            } else {
                $col.Type = $null;
            }
        }

        if ($col['NumberFormat'] -ne $null) {
            $col.NumberFormat = [XLNumberFormat]$col.NumberFormat;
        } else {
            $col.NumberFormat = switch ($col.Type) {
                    {$_ -eq [String]}        {[XLNumberFormat]::Text}
                    {$_ -eq [DateTime]}      {[XLNumberFormat]::DateTime}
                    default                  {[XLNumberFormat]::General}
                };
        }
    }

    $ColumnDefinitions;
}

. $PSScriptRoot\Add-XLTable.ps1
. $PSScriptRoot\Add-XLChart.ps1

Export-ModuleMember -Function New-XLFile
Export-ModuleMember -Function Add-XLSheet
Export-ModuleMember -Function Add-XLTable
Export-ModuleMember -Function Add-XLChart
Export-ModuleMember -Function Add-XLChartSeries

              