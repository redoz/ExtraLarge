BeforeAll { 
    . $PSScriptRoot\Setup.ps1
    . $PSScriptRoot\..\src\Resolve-XLRange.ps1
}

Describe "Resolve-XLRange/Range" {
    BeforeAll {
        $path = Join-Path $TestDrive 'test.xlsx'
        Copy-Item -Path $PSScriptRoot\data\WithNamedRange.xlsx -Destination $path
        $xlFile = Get-XLFile -Path $path 
    }

    Context "With valid range" {
        BeforeAll {
            $excelRange = $xlFile.Package.Workbook.Worksheets.Item(1).Cells[3,4,13,14]
            $xlRange = [XLRange]::new($xlFile.Package, $excelRange)
            $res = Resolve-XLRange -Range $xlRange
        }
        
        It "Range should be an [XLRange]" {
            $res.Range -is [XLRange] | Should -Be $true
        }
        
        It "Returned range should have same address as input" {
            $expectedRange = $xlRange.Range
            $actualRange = $res.Range.Range
            
            $expectedRange.Start.Address | Should -Be $actualRange.Start.Address
            $expectedRange.End.Address | Should -Be $actualRange.End.Address
        }
        
        It "InputObject should same as -Range parameter" {
            [Object]::ReferenceEquals($res.InputObject, $xlRange) | Should -be $true
        }
    }
   
}

Describe "Resolve-XLRange/SheetAndName" {
    BeforeAll {
        $path = Join-Path $TestDrive 'test.xlsx'
        Copy-Item -Path $PSScriptRoot\data\WithNamedRange.xlsx -Destination $path
        $xlFile = Get-XLFile -Path $path 
    }

    Context "With valid name scoped to same sheet" {
        BeforeAll {
            $xlSheet = Get-XLSheet -File $xlFile -Name "Sheet1"
            
            $res = Resolve-XLRange -Sheet $xlSheet -Name Sheet1Scope -Scope 'Sheet' 
        }
        
        It "Range should be an [XLRange]" {
            $res.Range -is [XLRange] | Should -Be $true
        }
        
        It "Returned range should match that of named range" {
            $res.Range.Range.Start.Address | Should -Be "B3"
            $res.Range.Range.Columns | Should -Be 1
            $res.Range.Range.Rows | Should -Be 1
        }
        
        It "InputObject should same as -Sheet parameter" {
            [Object]::ReferenceEquals($res.InputObject, $xlSheet) | Should -be $true
        }
    }
    
    Context "With valid name scoped to different sheet" {
        BeforeAll {
            $xlSheet = Get-XLSheet -File $xlFile -Name "Sheet1"
        }
        
        It "Should throw an exception" {
            { Resolve-XLRange -Sheet $xlSheet -Name Sheet2Scope -Scope 'Sheet' } | Should -Throw "Could not resolve range named 'Sheet2Scope'"
        }
    }
    
    Context "With valid name scoped to file" {
        BeforeAll {
            $xlSheet = Get-XLSheet -File $xlFile -Name "Sheet1"
            
            $res = Resolve-XLRange -Sheet $xlSheet -Name 'FileScope' -Scope File
        }
        
        It "Range should be an [XLRange]" {
            $res.Range -is [XLRange] | Should -Be $true
        }
        
        It "Returned range should match that of named range" {
            $res.Range.Range.Start.Address | Should -Be "B2"
            $res.Range.Range.Columns | Should -Be 1
            $res.Range.Range.Rows | Should -Be 1
        }
        
        It "Returned range should belong to 'Sheet2'" {
            $res.Range.Range.Worksheet.Name | Should -Be 'Sheet2'
        }
        
        It "InputObject should same as -Sheet parameter" {
            [Object]::ReferenceEquals($res.InputObject, $xlSheet) | Should -be $true
        }
    }
    
    Context "With invalid name scoped to sheet" {
        BeforeAll {
            $xlSheet = Get-XLSheet -File $xlFile -Name "Sheet1"
        }
        
        It "Should throw an exception" {
            { Resolve-XLRange -Sheet $xlSheet -Name XYZ -Scope 'Sheet' } | Should -Throw "Could not resolve range named 'XYZ'"
        }
    }
    
    Context "With invalid name scoped to file" {
        BeforeAll {
            $xlSheet = Get-XLSheet -File $xlFile -Name "Sheet1"
        }
        
        It "Should throw an exception" {
            { Resolve-XLRange -Sheet $xlSheet -Name XYZ -Scope File } | Should -Throw "Could not resolve range named 'XYZ'"
        }
    }
}

Describe "Resolve-XLRange/SheetAndRC" {
    BeforeAll {
        $path = Join-Path $TestDrive 'test.xlsx'
        Copy-Item -Path $PSScriptRoot\data\WithNamedRange.xlsx -Destination $path
        $xlFile = Get-XLFile -Path $path 
        $xlSheet = Get-XLSheet -File $xlFile -Name "Sheet1"
    }

    Context "With valid -Row and -Column" {
        BeforeAll {
            $res = Resolve-XLRange -Sheet $xlSheet -Row 2 -Column 2 
        }
        
        It "Range should be an [XLRange]" {
            $res.Range -is [XLRange] | Should -Be $true
        }
        
        It "Returned range should contain exactly 1 row and 1 column" {
            $res.Range.Range.Columns | Should -Be 1
            $res.Range.Range.Rows | Should -Be 1
        }
        
        It "Returned range should have address B2" {
            $res.Range.Range.Address | Should -be "B2"
        }
        
        It "InputObject should same as -Sheet parameter" {
            [Object]::ReferenceEquals($res.InputObject, $xlSheet) | Should -be $true
        }
    }
    
    Context "With oob -Row and -Column" {
        It "Should thrown an exception" {
             { Resolve-XLRange -Sheet $xlSheet -Row 0 -Column 0 } | Should -Throw
        }
    }    
    
    Context "With oob -Row" {
        It "Should thrown an exception" {
             { Resolve-XLRange -Sheet $xlSheet -Row 0 -Column 4 } | Should -Throw
        }
    }    
    
    Context "With oob -Column" {
        It "Should thrown an exception" {
             { Resolve-XLRange -Sheet $xlSheet -Row 4 -Column 0 } | Should -Throw
        }
    }       

    Context "With valid -FromRow -ToRow and -FromColumn -ToColumn" {
        
        BeforeAll {
            $res = Resolve-XLRange -Sheet $xlSheet -FromRow 2 -FromColumn 2 -ToRow 4 -ToColumn 4
        }
        
        It "Range should be an [XLRange]" {
            $res.Range -is [XLRange] | Should -Be $true
        }
        
        It "Returned range should contain exactly 3 rows and 3 columns" {
            $res.Range.Range.Columns | Should -Be 3
            $res.Range.Range.Rows | Should -Be 3
        }
        
        It "Returned range should have address B2:D4" {
            $res.Range.Range.Start.Address | Should -be "B2"
            $res.Range.Range.End.Address | Should -be "D4"
        }
        
        It "InputObject should same as -Sheet parameter" {
            [Object]::ReferenceEquals($res.InputObject, $xlSheet) | Should -be $true
        }
    }   
    
    Context "With swaped -FromRow -ToRow" {
        It "Should throw an exception" {
            { Resolve-XLRange -Sheet $xlSheet -FromRow 4 -FromColumn 2 -ToRow 2 -ToColumn 4 } | Should -Throw
        }
    }   
    
    Context "With swaped -FromColumn -ToColumn" {
        It "Should throw an exception" {
            { Resolve-XLRange -Sheet $xlSheet -FromRow 2 -FromColumn 4 -ToRow 4 -ToColumn 2 } | Should -Throw
        }
    }   
}

Describe "Resolve-XLRange/FileAndName" {
    BeforeAll {
        $path = Join-Path $TestDrive 'test.xlsx'
        Copy-Item -Path $PSScriptRoot\data\WithNamedRange.xlsx -Destination $path
        $xlFile = Get-XLFile -Path $path 
    }
    
    Context "With valid name scoped to sheet" {
        It "Should throw an exception" {
            { Resolve-XLRange -File $xlFile -Name Sheet2Scope } | Should -Throw "Could not resolve range named 'Sheet2Scope'"
        }
    }
    
    Context "With valid name scoped to file" {
        BeforeAll {
            $res = Resolve-XLRange -File $xlFile -Name 'FileScope'
        }
        
        It "Range should be an [XLRange]" {
            $res.Range -is [XLRange] | Should -Be $true
        }
        
        It "Returned range should match that of named range" {
            $res.Range.Range.Start.Address | Should -Be "B2"
            $res.Range.Range.Columns | Should -Be 1
            $res.Range.Range.Rows | Should -Be 1
        }
        
        It "Returned range should belong to 'Sheet2'" {
            $res.Range.Range.Worksheet.Name | Should -Be 'Sheet2'
        }
        
        It "InputObject should same as -File parameter" {
            [Object]::ReferenceEquals($res.InputObject, $xlFile) | Should -be $true
        }
    }
    
    Context "With invalid name" {      
        It "Should throw an exception" {
            { Resolve-XLRange -File $xlFile -Name XYZ } | Should -Throw "Could not resolve range named 'XYZ'"
        }
    }
}

Describe "Resolve-XLRange/SheetAndAddress" {
    BeforeAll {
        $path = Join-Path $TestDrive 'test.xlsx'
        Copy-Item -Path $PSScriptRoot\data\WithNamedRange.xlsx -Destination $path
        $xlFile = Get-XLFile -Path $path 
        $xlSheet = Get-XLSheet -File $xlFile -Index 1
    }
    
    Context "A1 address without sheet information" {
        BeforeAll {
            $res = Resolve-XLRange -Sheet $xlSheet -Address "A1"
            $xlRange = $res.Range
        }
        
        It "Should not return null" {
            $xlRange | Should -Not -Be $null
        }
        
        It "Range width and height should equal 1" {
            $xlRange.Range.Columns | Should -Be 1
            $xlRange.Range.Rows | Should -Be 1
        }
        
        It "Range Row and Column should equal 1" {
            $xlRange.Range.Start.Row | Should -Be 1
            $xlRange.Range.Start.Column | Should -Be 1
            $xlRange.Range.End.Row | Should -Be 1
            $xlRange.Range.End.Column | Should -Be 1
        }
    }
    
    Context "R1C1 address without sheet information" {
        BeforeAll {
            $res = Resolve-XLRange -Sheet $xlSheet -Address "R1C1"
            $xlRange = $res.Range
        }
        
        It "Should not return null" {
            $xlRange | Should -Not -Be $null
        }
        
        It "Range width and height should equal 1" {
            $xlRange.Range.Columns | Should -Be 1
            $xlRange.Range.Rows | Should -Be 1
        }
        
        It "Range Row and Column should equal 1" {
            $xlRange.Range.Start.Row | Should -Be 1
            $xlRange.Range.Start.Column | Should -Be 1
            $xlRange.Range.End.Row | Should -Be 1
            $xlRange.Range.End.Column | Should -Be 1
        }
    }
    
    Context "Invalid A1 address with sheet information" {
        It "Should throw an exception" {
            { Resolve-XLRange -Sheet $xlSheet -Address "Sheet1!A0" } | Should -Throw "Invalid address: 'A0'"
        }
    } 
    
    Context "Invalid R1C1 address with sheet information" {
        It "Should throw an exception" {
            { Resolve-XLRange -Sheet $xlSheet -Address "Sheet1!R0C0" } | Should -Throw "Invalid address: 'R0C0'"
        }
    }    
    
    Context "Bogus address" {
        It "Should throw an exception" {
            { Resolve-XLRange -Sheet $xlSheet -Address "XYZ" } | Should -Throw "Invalid address: 'XYZ'"
        }
    } 
    
    Context "Non-existing sheet" {
        It "Should throw an exception" {
            { Resolve-XLRange -Sheet $xlSheet -Address "Sheet3!A1" } | Should -Throw "Sheet not found: 'Sheet3'"
        }
    } 
    
    Context "Valid sheet and single cell A1 address" {
        BeforeAll {
            $res = Resolve-XLRange -Sheet $xlSheet -Address "Sheet2!A1" 
            $xlRange = $res.Range
        }
        It "Should not return null" {
            $xlRange | Should -Not -Be $null
        }
        
        It "Range width and height should equal 1" {
            $xlRange.Range.Columns | Should -Be 1
            $xlRange.Range.Rows | Should -Be 1
        }
        
        It "Range Row and Column should equal 1" {
            $xlRange.Range.Start.Row | Should -Be 1
            $xlRange.Range.Start.Column | Should -Be 1
            $xlRange.Range.End.Row | Should -Be 1
            $xlRange.Range.End.Column | Should -Be 1
        }
    } 
    
    Context "Valid sheet and single cell R1C1 address" {
        BeforeAll {
            $res = Resolve-XLRange -Sheet $xlSheet -Address "Sheet2!R1C1"
            $xlRange = $res.Range
        }
        It "Should not return null" {
            $xlRange | Should -Not -Be $null
        }
        
        It "The sheet should match address" {
            $xlRange.Range.Worksheet.Name | Should -Be "Sheet2"
        }
        
        It "Range width and height should equal 1" {
            $xlRange.Range.Columns | Should -Be 1
            $xlRange.Range.Rows | Should -Be 1
        }
        
        It "Range Row and Column should equal 1" {
            $xlRange.Range.Start.Row | Should -Be 1
            $xlRange.Range.Start.Column | Should -Be 1
            $xlRange.Range.End.Row | Should -Be 1
            $xlRange.Range.End.Column | Should -Be 1
        }
    }
    
    Context "Valid sheet and single cell A1 address" {
        BeforeAll {
            $res = Resolve-XLRange -Sheet $xlSheet -Address "Sheet2!A1"
            $xlRange = $res.Range
        }
        It "Should not return null" {
            $xlRange | Should -Not -Be $null
        }

        It "The sheet should match address" {
            $xlRange.Range.Worksheet.Name | Should -Be "Sheet2"
        }        
        
        It "Range width and height should equal 1" {
            $xlRange.Range.Columns | Should -Be 1
            $xlRange.Range.Rows | Should -Be 1
        }
        
        It "Range Row and Column should equal 1" {
            $xlRange.Range.Start.Row | Should -Be 1
            $xlRange.Range.Start.Column | Should -Be 1
            $xlRange.Range.End.Row | Should -Be 1
            $xlRange.Range.End.Column | Should -Be 1
        }
    } 
    
    Context "Valid sheet and single cell R1C1 address" {
        BeforeAll {
            $res = Resolve-XLRange -Sheet $xlSheet -Address "Sheet2!R1C1"
            $xlRange = $res.Range
        }
        It "Should not return null" {
            $xlRange | Should -Not -Be $null
        }
        
        It "The sheet should match address" {
            $xlRange.Range.Worksheet.Name | Should -Be "Sheet2"
        }        
        
        It "Range width and height should equal 1" {
            $xlRange.Range.Columns | Should -Be 1
            $xlRange.Range.Rows | Should -Be 1
        }
        
        It "Range Row and Column should equal 1" {
            $xlRange.Range.Start.Row | Should -Be 1
            $xlRange.Range.Start.Column | Should -Be 1
            $xlRange.Range.End.Row | Should -Be 1
            $xlRange.Range.End.Column | Should -Be 1
        }
    }    
    
    Context "Valid sheet and A1 address range" {
        BeforeAll {
            $res = Resolve-XLRange -Sheet $xlSheet -Address "Sheet2!A1:C3"
            $xlRange = $res.Range
        }
        It "Should not return null" {
            $xlRange | Should -Not -Be $null
        }
        
        It "The sheet should match address" {
            $xlRange.Range.Worksheet.Name | Should -Be "Sheet2"
        }
        
        It "Range width and height should equal 3" {
            $xlRange.Range.Columns | Should -Be 3
            $xlRange.Range.Rows | Should -Be 3
        }
        
        It "Start address Row and Column should equal 1" {
            $xlRange.Range.Start.Row | Should -Be 1
            $xlRange.Range.Start.Column | Should -Be 1
        }
        
        It "End address Row and Column should equal 3" {
            $xlRange.Range.End.Row | Should -Be 3
            $xlRange.Range.End.Column | Should -Be 3
        }
    } 
    
    Context "Valid sheet and R1C1 address range" {
        BeforeAll {
            $res = Resolve-XLRange -Sheet $xlSheet -Address "Sheet2!R1C1:R3C3"
            $xlRange = $res.Range
        }
        It "Should not return null" {
            $xlRange | Should -Not -Be $null
        }
        
        It "The sheet should match address" {
            $xlRange.Range.Worksheet.Name | Should -Be "Sheet2"
        }
        
        It "Range width and height should equal 3" {
            $xlRange.Range.Columns | Should -Be 3
            $xlRange.Range.Rows | Should -Be 3
        }
        
        It "Start address Row and Column should equal 1" {
            $xlRange.Range.Start.Row | Should -Be 1
            $xlRange.Range.Start.Column | Should -Be 1
        }
        
        It "End address Row and Column should equal 3" {
            $xlRange.Range.End.Row | Should -Be 3
            $xlRange.Range.End.Column | Should -Be 3
        }
    }   
}

Describe "Resolve-XLRange/FileAndAddress" {
    BeforeAll {
        $path = Join-Path $TestDrive 'test.xlsx'
        Copy-Item -Path $PSScriptRoot\data\WithNamedRange.xlsx -Destination $path
        $xlFile = Get-XLFile -Path $path 
    }
    
    Context "A1 address without sheet information" {
        It "Should throw an exception" {
            { Resolve-XLRange -File $xlFile -Address "A1" } | Should -Throw "No sheet specified in address: 'A1'"
        }
    }
    
    Context "R1C1 address without sheet information" {
        It "Should throw an exception" {
            { Resolve-XLRange -File $xlFile -Address "R1C1" } | Should -Throw "No sheet specified in address: 'R1C1'"
        }
    }
    
    Context "Invalid A1 address with sheet information" {
        It "Should throw an exception" {
            { Resolve-XLRange -File $xlFile -Address "Sheet1!A0" } | Should -Throw "Invalid address: 'A0'"
        }
    } 
    
    Context "Invalid R1C1 address with sheet information" {
        It "Should throw an exception" {
            { Resolve-XLRange -File $xlFile -Address "Sheet1!R0C0" } | Should -Throw "Invalid address: 'R0C0'"
        }
    }    
    
    Context "Bogus address" {
        It "Should throw an exception" {
            { Resolve-XLRange -File $xlFile -Address "XYZ" } | Should -Throw "Invalid address: 'XYZ'"
        }
    } 
    
    Context "Non-existing sheet" {
        It "Should throw an exception" {
            { Resolve-XLRange -File $xlFile -Address "Sheet3!A1" } | Should -Throw "Sheet not found: 'Sheet3'"
        }
    } 
    
    Context "Valid sheet and single cell A1 address" {
        BeforeAll {
            $res = Resolve-XLRange -File $xlFile -Address "Sheet1!A1"
            $xlRange = $res.Range
        }
        It "Should not return null" {
            $xlRange | Should -Not -Be $null
        }
        
        It "Range width and height should equal 1" {
            $xlRange.Range.Columns | Should -Be 1
            $xlRange.Range.Rows | Should -Be 1
        }
        
        It "Range Row and Column should equal 1" {
            $xlRange.Range.Start.Row | Should -Be 1
            $xlRange.Range.Start.Column | Should -Be 1
            $xlRange.Range.End.Row | Should -Be 1
            $xlRange.Range.End.Column | Should -Be 1
        }
    } 
    
    Context "Valid sheet and single cell R1C1 address" {
        BeforeAll {
            $res = Resolve-XLRange -File $xlFile -Address "Sheet1!R1C1"
            $xlRange = $res.Range
        }
        It "Should not return null" {
            $xlRange | Should -Not -Be $null
        }
        
        It "Range width and height should equal 1" {
            $xlRange.Range.Columns | Should -Be 1
            $xlRange.Range.Rows | Should -Be 1
        }
        
        It "Range Row and Column should equal 1" {
            $xlRange.Range.Start.Row | Should -Be 1
            $xlRange.Range.Start.Column | Should -Be 1
            $xlRange.Range.End.Row | Should -Be 1
            $xlRange.Range.End.Column | Should -Be 1
        }
    }
    
    Context "Valid sheet and single cell A1 address" {
        BeforeAll {
            $res = Resolve-XLRange -File $xlFile -Address "Sheet1!A1"
            $xlRange = $res.Range
        }
        It "Should not return null" {
            $xlRange | Should -Not -Be $null
        }
        
        It "Range width and height should equal 1" {
            $xlRange.Range.Columns | Should -Be 1
            $xlRange.Range.Rows | Should -Be 1
        }
        
        It "Range Row and Column should equal 1" {
            $xlRange.Range.Start.Row | Should -Be 1
            $xlRange.Range.Start.Column | Should -Be 1
            $xlRange.Range.End.Row | Should -Be 1
            $xlRange.Range.End.Column | Should -Be 1
        }
    } 
    
    Context "Valid sheet and single cell R1C1 address" {
        BeforeAll {
            $res = Resolve-XLRange -File $xlFile -Address "Sheet1!R1C1"
            $xlRange = $res.Range
        }

        It "Should not return null" {
            $xlRange | Should -Not -Be $null
        }
        
        It "Range width and height should equal 1" {
            $xlRange.Range.Columns | Should -Be 1
            $xlRange.Range.Rows | Should -Be 1
        }
        
        It "Range Row and Column should equal 1" {
            $xlRange.Range.Start.Row | Should -Be 1
            $xlRange.Range.Start.Column | Should -Be 1
            $xlRange.Range.End.Row | Should -Be 1
            $xlRange.Range.End.Column | Should -Be 1
        }
    }    
    
    Context "Valid sheet and A1 address range" {
        BeforeAll {
            $res = Resolve-XLRange -File $xlFile -Address "Sheet1!A1:C3"
            $xlRange = $res.Range
        }
        
        It "Should not return null" {
            $xlRange | Should -Not -Be $null
        }
        
        It "Range width and height should equal 3" {
            $xlRange.Range.Columns | Should -Be 3
            $xlRange.Range.Rows | Should -Be 3
        }
        
        It "Start address Row and Column should equal 1" {
            $xlRange.Range.Start.Row | Should -Be 1
            $xlRange.Range.Start.Column | Should -Be 1
        }
        
        It "End address Row and Column should equal 3" {
            $xlRange.Range.End.Row | Should -Be 3
            $xlRange.Range.End.Column | Should -Be 3
        }
    } 
    
    Context "Valid sheet and R1C1 address range" {
        BeforeAll {
            $res = Resolve-XLRange -File $xlFile -Address "Sheet1!R1C1:R3C3"
            $xlRange = $res.Range
        }
        It "Should not return null" {
            $xlRange | Should -Not -Be $null
        }
        
        It "Range width and height should equal 3" {
            $xlRange.Range.Columns | Should -Be 3
            $xlRange.Range.Rows | Should -Be 3
        }
        
        It "Start address Row and Column should equal 1" {
            $xlRange.Range.Start.Row | Should -Be 1
            $xlRange.Range.Start.Column | Should -Be 1
        }
        
        It "End address Row and Column should equal 3" {
            $xlRange.Range.End.Row | Should -Be 3
            $xlRange.Range.End.Column | Should -Be 3
        }
    }   
}    
