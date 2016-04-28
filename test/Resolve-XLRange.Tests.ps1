. .\test\Setup.ps1
. .\src\Resolve-XLRange.ps1

Describe "Resolve-XLRange/Range" {
    $path = Join-Path $TestDrive 'test.xlsx'
    Copy-Item -Path .\test\data\WithNamedRange.xlsx -Destination $path
    $xlFile = Get-XLFile -Path $path 

    Context "With valid range" {
        $excelRange = $xlFile.Package.Workbook.Worksheets.Item(1).Cells[3,4,13,14]
        $xlRange = [XLRange]::new($xlFile.Package, $excelRange)
        $res = Resolve-XLRange -Range $xlRange
        
        It "Range should be an [XLRange]" {
            $res.Range -is [XLRange] | Should Be $true
        }
        
        It "Returned range should have same address as input" {
            $expectedRange = $xlRange.Range
            $actualRange = $res.Range.Range
            
            $expectedRange.Start.Address | Should Be $actualRange.Start.Address
            $expectedRange.End.Address | Should Be $actualRange.End.Address
        }
        
        It "InputObject should same as -Range parameter" {
            [Object]::ReferenceEquals($res.InputObject, $xlRange) | Should be $true
        }
    }
   
}

Describe "Resolve-XLRange/SheetAndName" {
    $path = Join-Path $TestDrive 'test.xlsx'
    Copy-Item -Path .\test\data\WithNamedRange.xlsx -Destination $path
    $xlFile = Get-XLFile -Path $path 

    Context "With valid name scoped to same sheet" {
        $xlSheet = Get-XLSheet -File $xlFile -Name "Sheet1"
        
        $res = Resolve-XLRange -Sheet $xlSheet -Name Sheet1Scope -Scope 'Sheet' 
        
        It "Range should be an [XLRange]" {
            $res.Range -is [XLRange] | Should Be $true
        }
        
        It "Returned range should match that of named range" {
            $res.Range.Range.Start.Address | Should Be "B3"
            $res.Range.Range.Columns | Should Be 1
            $res.Range.Range.Rows | Should Be 1
        }
        
        It "InputObject should same as -Sheet parameter" {
            [Object]::ReferenceEquals($res.InputObject, $xlSheet) | Should be $true
        }
    }
    
    Context "With valid name scoped to different sheet" {
        $xlSheet = Get-XLSheet -File $xlFile -Name "Sheet1"
        
        { Resolve-XLRange -Sheet $xlSheet -Name Sheet2Scope -Scope 'Sheet' } | Should Throw "Could not resolve range named 'Sheet2Scope'"
    }
    
    Context "With valid name scoped to file" {
        $xlSheet = Get-XLSheet -File $xlFile -Name "Sheet1"
        
        $res = Resolve-XLRange -Sheet $xlSheet -Name 'FileScope' -Scope File
        
        It "Range should be an [XLRange]" {
            $res.Range -is [XLRange] | Should Be $true
        }
        
        It "Returned range should match that of named range" {
            $res.Range.Range.Start.Address | Should Be "B2"
            $res.Range.Range.Columns | Should Be 1
            $res.Range.Range.Rows | Should Be 1
        }
        
        It "Returned range should belong to 'Sheet2'" {
            $res.Range.Range.Worksheet.Name | Should Be 'Sheet2'
        }
        
        It "InputObject should same as -Sheet parameter" {
            [Object]::ReferenceEquals($res.InputObject, $xlSheet) | Should be $true
        }
    }
    
    Context "With invalid name scoped to sheet" {
        $xlSheet = Get-XLSheet -File $xlFile -Name "Sheet1"
        
        It "Should throw an exception" {
            { Resolve-XLRange -Sheet $xlSheet -Name XYZ -Scope 'Sheet' } | Should Throw "Could not resolve range named 'XYZ'"
        }
    }
    
    Context "With invalid name scoped to file" {
        $xlSheet = Get-XLSheet -File $xlFile -Name "Sheet1"
        
        It "Should throw an exception" {
            { Resolve-XLRange -Sheet $xlSheet -Name XYZ -Scope File } | Should Throw "Could not resolve range named 'XYZ'"
        }
    }
}

Describe "Resolve-XLRange/SheetAndRC" {
}

Describe "Resolve-XLRange/FileAndName" {
    $path = Join-Path $TestDrive 'test.xlsx'
    Copy-Item -Path .\test\data\WithNamedRange.xlsx -Destination $path
    $xlFile = Get-XLFile -Path $path 
    
    Context "With valid name scoped to sheet" {
        It "Should throw an exception" {
            { Resolve-XLRange -File $xlFile -Name Sheet2Scope } | Should Throw "Could not resolve range named 'Sheet2Scope'"
        }
    }
    
    Context "With valid name scoped to file" {
        $res = Resolve-XLRange -File $xlFile -Name 'FileScope'
        
        It "Range should be an [XLRange]" {
            $res.Range -is [XLRange] | Should Be $true
        }
        
        It "Returned range should match that of named range" {
            $res.Range.Range.Start.Address | Should Be "B2"
            $res.Range.Range.Columns | Should Be 1
            $res.Range.Range.Rows | Should Be 1
        }
        
        It "Returned range should belong to 'Sheet2'" {
            $res.Range.Range.Worksheet.Name | Should Be 'Sheet2'
        }
        
        It "InputObject should same as -File parameter" {
            [Object]::ReferenceEquals($res.InputObject, $xlFile) | Should be $true
        }
    }
    
    Context "With invalid name" {      
        It "Should throw an exception" {
            { Resolve-XLRange -File $xlFile -Name XYZ } | Should Throw "Could not resolve range named 'XYZ'"
        }
    }
}

Describe "Resolve-XLRange/SheetAndAddress" {
}

Describe "Resolve-XLRange/FileAndAddress" {
}
