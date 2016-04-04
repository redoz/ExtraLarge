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

function Get-Columns($Datum , $ColumnDefinitions) {
    if ($ColumnDefinitions -eq $null) {
        $ColumnDefinitions = Get-Member -InputObject $Datum -MemberType Properties | ForEach-Object -Process { @{Name = $_.Name; Property = $_.Name; Type = [Type]$_.TypeName} }
    }

    foreach ($col in $ColumnDefinitions) {
        if (-not $col.ContainsKey('Expression')) {
            $propertyName = $col.Property;
            $col.Expression = { $_.$propertyName }.GetNewClosure();
        }

        if (-not $col.ContainsKey('Type')) {
            
        }
    }

    $ColumnDefinitions;
}

function Add-XLTable {
param(
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true)]
    [OfficeOpenXml.ExcelWorksheet]$Sheet,
    [Parameter(Mandatory = $true)]
    [string]$Name,
    [Parameter(Mandatory = $true)]
    [object]$Data,
    [object[]]$Columns = $null,
    [int]$Row = 0,
    [int]$Column = 0,
    [Switch]$Transpose = $false,
    [switch]$PassThru = $false
)  
begin{
    #validate some input
    if ($Columns -ne $null) {
        $Columns = $Columns | ForEach-Object -Process {
                switch ($_) {
                    {$_ -is [string]} { @{Name = $_; Property = $_}; break; }
                    {$_ -is [Dictionary]} { 
                        if ($col.Name -eq $null) {
                            if ($col.Property -ne $null) {
                                $col.Name = $col.Property;        
                            } else {
                                throw "Name or Property is required for column definition";
                            }
                        }

                        if ($col.Property -eq $null) {
                            if ($col.Expression -eq $null) {
                                throw "Property or Expression is requierd for column definitions";
                            }
                        }
                        $_;
                        break;
                    }
                    default {throw "Invalid column definition: " + $_;}
                }
            };
        }


    # extract tabular data 
    $rows = [System.Collections.Generic.List[object[]]]::new(); 

    if ($Data -is [PSObject]) {

        $Columns = Get-Columns -Datum $Data -ColumnDefinitions $Columns

        $headerRow = @($Columns.Name);
        
        # this is a hack, and probably slow
        $dataRow = @($Columns | ForEach-Object -Process { $Data | ForEach-Object -Process $_.Expression })

        $rows.Add($headerRow);
        $rows.Add($dataRow);
        
    } elseif ($Data -is [System.Collections.IDictionary]) {

        foreach ($kvp in $Data.GetEnumerator()) {
            $rows.Add(@($kvp.Name, $kvp.Value))
        }
    } else {
        [bool]$firstIteration = $true;
        foreach ($datum in $Data) {
            if ($firstIteration) {
                $Columns = Get-Columns -Datum $datum -ColumnDefinitions $Columns
                # add header row
                $rows.Add($Columns.Name);
                $firstIteration = $false;
            }

            $rows.Add(@($Columns | ForEach-Object -Process { $datum | ForEach-Object -Process $_.Expression }));
        }
    }

}
process{


    if ($Transpose.IsPresent) {
        throw "Not implemented yet";
        #transpose data
    }

    # find empty location in sheet that can accomodate data
    [int]$tableHeight = $rows.Count;
    [int]$tableWidth = $rows[0].Count;

    if ($Row -eq 0 -and $Column -eq 0) {
        if ($Sheet.Dimension -ne $null) {
            $Row = $Sheet.Dimension.End.Row + 2;
            $Column = $Sheet.Dimension.Start.Column;
        } else {
            $Row = 2;
            $Column = 2;
        }
    } elseif ($Row -eq 0) {
        if ($Sheet.Dimension -ne $null) {
            $Row = $Sheet.Dimension.End.Row + 2;
        } else {
            $Row = 2;
        }
    } elseif ($Column -eq 0) {
        if ($Sheet.Dimension -ne $null) {
            $Column = $Sheet.Dimension.End.Column + 2;        
        } else {
            $Column = 2;
        }
    }

    # write data into sheet
    [int]$currentRow = $Row;
    foreach ($dataRow in $rows) {
        [int]$currentColumn = $Column;
        foreach ($value in $dataRow) {
            $cell = $Sheet.Cells[$currentRow, $currentColumn];
            # TODO add this as part of the column definition
            $cell.Style.Numberformat.Format = "General";
            # TODO if the column is numeric type conver this to an actual number
            $cell.Value = $value;
            
            $currentColumn++;
        }
        $currentRow++;
    }

    # create table
    $tableRange = [OfficeOpenXml.ExcelRange]::GetAddress($Row, $Column, $Row + $tableHeight - 1, $Column + $tableWidth - 1);
    $table = $Sheet.Tables.Add($tableRange, $Name)

    if ($PassThru.IsPresent) {
        Write-Output -InputObject $Sheet
    }
}
end{}
}

Export-ModuleMember -Function New-XLFile
Export-ModuleMember -Function Add-XLSheet
Export-ModuleMember -Function Add-XLTable

              