
[Regex]$Script:AddressRegex = [Regex]::new(@"
^(?:(?<sheet>[^!]+)!)?
     (?:((?<r1c1>R\d+C\d+(?::R\d+C\d+)?))|
        (?<a1>[A-Z]+\d+(?::[A-Z]+\d+)?))$
"@, [System.Text.RegularExpressions.RegexOptions]::Compiled -bor [System.Text.RegularExpressions.RegexOptions]::IgnorePatternWhitespace  -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

function Resolve-XLRange {
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true, ParameterSetName = "Range")]
    [XLRange]$Range,

    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true, ParameterSetName = "SheetAndRC")]
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true, ParameterSetName = "SheetAndName")]
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true, ParameterSetName = "SheetAndAddress")]    
    [XLSheet]$Sheet,
    
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true, ParameterSetName = "FileAndName")]
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true, ParameterSetName = "FileAndAddress")]
    [XLFile]$File,

    [Parameter(Mandatory = $true, Position = 1, ParameterSetName = "SheetAndRC")]
    [Alias("Row")]
    [int]$FromRow,

    [Parameter(ParameterSetName = "SheetAndRC")]
    [int]$ToRow = $FromRow,

    [Parameter(Mandatory = $true, Position = 2, ParameterSetName = "SheetAndRC")]
    [Alias("Column")]
    [int]$FromColumn,

    [Parameter(ParameterSetName = "SheetAndRC")]
    [int]$ToColumn = $FromColumn,
    
    [Parameter(ParameterSetName = "SheetAndName")]
    [XLScope]$Scope = [XLScope]::Any,

    [Parameter(Mandatory = $true, ParameterSetName = "FileAndName")]
    [Parameter(Mandatory = $true, ParameterSetName = "SheetAndName")]
    [string]$Name,
    
    [Parameter(Mandatory = $true, ParameterSetName = "FileAndAddress")]
    [Parameter(Mandatory = $true, ParameterSetName = "SheetAndAddress")]
    [string]$Address
)
    $inputObject = $null
    [XLRange]$xlRange = $null
    if ($PSCmdlet.ParameterSetName -eq 'Range') {
        $xlRange = $Range
        $inputObject = $Range
    } elseif ($PSCmdlet.ParameterSetName -eq 'SheetAndRC') {
        $xlRange = [XLRange]::new($Sheet.Owner, $Sheet.Worksheet.Cells.Item($FromRow, $FromColumn, $ToRow, $ToColumn))
        $inputObject = $Sheet
    } elseif ($PSCmdlet.ParameterSetName -eq 'SheetAndName') {
        $namedRange = $null
        if ($Scope -band [XLScope]::Sheet -and $Sheet.Worksheet.Names.ContainsKey($Name)) {
            $namedRange = $Sheet.Worksheet.Names[$Name]
        }
        
        if ($namedRange -eq $null -and $Scope -band [XLScope]::File -and $Sheet.Worksheet.Workbook.Names.ContainsKey($Name)) {
            $namedRange = $Sheet.Worksheet.Workbook.Names[$Name]
        }
        if ($namedRange -eq $null) {
            throw "Could not resolve range named '$Name'"
        }
        $xlRange = [XLRange]::new($Sheet.Owner, $namedRange)
        $inputObject = $Sheet
    } elseif ($PSCmdlet.ParameterSetName -eq 'FileAndName') {
        $namedRange = $null
        if ($File.Package.Workbook.Names.ContainsKey($Name)) {
            $namedRange = $File.Package.Workbook.Names[$Name]
        } else {
            throw "Could not resolve range named '$Name'"
        }
        $xlRange = [XLRange]::new($File.Package, $namedRange)
        $inputObject = $File
    } elseif ($PSCmdlet.ParameterSetName.EndsWith('Address')) {
        $match = $Script:AddressRegex.Match($Address)
        if (-not $match.Success) {
            throw "Invalid address: '$Address'"
        }
                
        [OfficeOpenXml.ExcelWorksheet]$targetSheet = $null
        if ($match.Groups['sheet'].Success) {
            $sheetName = $match.Groups['sheet'].Value
            $workbook = $null
            if ($PSCmdlet.ParameterSetName.StartsWith('File')) {
                $workbook = $File.Package.Workbook
            } elseif ($PSCmdlet.ParameterSetName.StartsWith('Sheet')) {
                $workbook = $Sheet.Worksheet.Workbook
            }
            $targetSheet = $workbook.Worksheets[$sheetName]
            if ($targetSheet -eq $null) {
                throw "Sheet not found: '$sheetName'"
            }
        } else {
            if ($PSCmdlet.ParameterSetName.StartsWith('File')) {
                throw "No sheet specified in address: '$Address'"
            } elseif ($PSCmdlet.ParameterSetName.StartsWith('Sheet')) {
                $targetSheet = $Sheet.Worksheet
            }                
        }
        
        $a1 = $null
        if ($match.Groups['r1c1'].Success) {
            $a1 = [OfficeOpenXml.ExcelCellBase]::TranslateFromR1C1($match.Groups['r1c1'].Value, 0, 0)
            if ($a1 -match "#REF") {
                throw "Invalid address: '$($match.Groups['r1c1'].Value)'"
            }
        } else {
            $a1 = $match.Groups['a1'].Value
        }
        
        $owner = $null
        if ($PSCmdlet.ParameterSetName.StartsWith('File')) {
            $owner = $File.Package
            $inputObject = $File
        } elseif ($PSCmdlet.ParameterSetName.StartsWith('Sheet')) {
            $owner = $Sheet.Owner
            $inputObject = $Sheet
        }
        
        if (-not [OfficeOpenXml.FormulaParsing.ExcelUtilities.ExcelAddressUtil]::IsValidAddress($a1)) {
            throw "Invalid address: '$a1'"
        }
        
        # couldn't find a better way to validate
        $excelAddress = [OfficeOpenXml.ExcelCellAddress]::new($a1)
        if ($excelAddress.Row -eq 0 -or $excelAddress.Column -eq 0) {
            throw "Invalid address: '$a1'"
        }
        
        $excelRange = $targetSheet.Cells[$a1]
        
        $xlRange = [XLRange]::new($owner, $excelRange)
    }
    [pscustomobject][ordered]@{
        Range = $xlRange
        InputObject = $inputObject
    }
}    