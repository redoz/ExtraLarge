BeforeAll { 
    . $PSScriptRoot\Setup.ps1
}

Describe "Get-XLSheet/FileAndName" {
    Context "With single sheet" {
        BeforeAll {
            $path = Get-TestPath
            $file = New-XLFile -Path $path -PassThru -NoDefaultSheet
            $sheet = Add-XLSheet -Name "Sheet 1" -File $file
        }
        
        It "Should return existing sheet" {
            $result = Get-XLSheet -File $file -Name "Sheet 1"
            $result | Should -Not -Be $null
            $result.Owner | Should -Be $file.Package
            $result.Worksheet | Should -be $sheet.Worksheet
        }
        
        It "Should throw for non-existing sheet" {
            { Get-XLSheet -File $file -Name "Sheet 2" } | Should -Throw
        }
    }
}

Describe "Get-XLSheet/FileAndIndex" {
    Context "With single sheet" {
        BeforeAll {
            $path = Get-TestPath
            $file = New-XLFile -Path $path -PassThru -NoDefaultSheet
            $sheet = Add-XLSheet -Name "Sheet 1" -File $file
        }
        
        It "Should return existing sheet" {
            $result = Get-XLSheet -File $file -Index 0
            $result | Should -Not -Be $null
            $result.Owner | Should -Be $file.Package
            $result.Worksheet | Should -be $sheet.Worksheet
        }
        
        It "Should throw for index out of bounds" {
            { Get-XLSheet -File $file -Index 2 } | Should -Throw
        }
    }
}

Describe "Get-XLSheet/File" {
    Context "With no sheets" {
        BeforeAll {
            $path = Get-TestPath
            $file = New-XLFile -Path $path -PassThru -NoDefaultSheet
        }
        
        It "Should return nothing" {
            Get-XLSheet -File $file | Should -Be $null
        }
    }

    Context "With single sheet" {
        BeforeAll {
            $path = Get-TestPath
            $file = New-XLFile -Path $path -PassThru -NoDefaultSheet
            $sheet = Add-XLSheet -Name "Sheet 1" -File $file
        }
        
        It "Should return existing sheet" {
            $result = Get-XLSheet -File $file
            $result | Should -Not -Be $null
            $result.Owner | Should -Be $file.Package
            $result.Worksheet | Should -be $sheet.Worksheet
        }
    } 
    Context "With multiple sheets" {
        BeforeAll {
            $path = Get-TestPath
            $file = New-XLFile -Path $path -PassThru -NoDefaultSheet
            $sheet1 = Add-XLSheet -Name "Sheet 1" -File $file
            $sheet2 = Add-XLSheet -Name "Sheet 2" -File $file        
        }
        
        It "Should return existing sheet" {
            $result = Get-XLSheet -File $file
            $result | Should -Not -Be $null
            $result.Count | Should -Be 2
            $result[0].Owner | Should -Be $file.Package
            $result[0].Worksheet | Should -be $sheet1.Worksheet
            $result[1].Owner | Should -Be $file.Package
            $result[1].Worksheet | Should -be $sheet2.Worksheet            
        }
    }    
}

Describe "Get-XLSheet/PathAndName" {
    Context "With single sheet" {
        BeforeAll {
            $path = Get-TestPath
            $file = New-XLFile -Path $path -PassThru -NoDefaultSheet
            $sheet = Add-XLSheet -Name "Sheet 1" -File $file
            Save-XLFile -File $file
        }
        
        It "Should return existing sheet" {
            $result = Get-XLSheet -Path $path -Name "Sheet 1"
            $result | Should -Not -Be $null
            $result.Owner | Should -Not -Be $null
            $result.Worksheet | Should -Not -Be $null
        }
        
        It "Should throw for non-existing sheet" {
            { Get-XLSheet -Path $path -Name "Sheet 2" } | Should -Throw
        }
    }
}

Describe "Get-XLSheet/PathAndIndex" {
    Context "With single sheet" {
        BeforeAll {
            $path = Get-TestPath
            $file = New-XLFile -Path $path -PassThru -NoDefaultSheet
            $sheet = Add-XLSheet -Name "Sheet 1" -File $file
            Save-XLFile -File $file
        }
        
        It "Should return existing sheet" {
            $result = Get-XLSheet -Path $path -Index 0
            $result | Should -Not -Be $null
            $result.Owner | Should -Not -Be $null
            $result.Worksheet | Should -Not -Be $null
        }
        
        It "Should throw for index out of bounds" {
            { Get-XLSheet -Path $path -Index 2 } | Should -Throw
        }
    }
}

Describe "Get-XLSheet/Path" {
    Context "With non-existing file" {
        It "Should throw" {
            { Get-XLSheet -Path "Test.xlsx" } | Should -Throw
        }
    }
    Context "With single sheet" {
        BeforeAll {
            $path = Get-TestPath
            $file = New-XLFile -Path $path -PassThru -NoDefaultSheet
            $sheet = Add-XLSheet -Name "Sheet 1" -File $file
            Save-XLFile -File $file
        }
                
        It "Should return existing sheet" {
            $result = Get-XLSheet -Path $path
            $result | Should -Not -Be $null
            $result.Owner | Should -Not -Be $null
            $result.Worksheet | Should -Not -Be $null
        }
    } 
    Context "With multiple sheets" {
        BeforeAll {
            $path = Get-TestPath
            $file = New-XLFile -Path $path -PassThru -NoDefaultSheet
            $sheet1 = Add-XLSheet -Name "Sheet 1" -File $file
            $sheet2 = Add-XLSheet -Name "Sheet 2" -File $file
            Save-XLFile -File $file
        }
        
        It "Should return existing sheet" {
            $result = Get-XLSheet -Path $path
            $result | Should -Not -Be $null
            $result.Count | Should -Be 2
            $result[0].Owner | Should -Not -Be $null
            $result[0].Worksheet | Should -Not -Be $null
            $result[1].Owner | Should -Not -Be $null
            $result[1].Worksheet | Should -Not -Be $null
        }
    }    
}


