BeforeAll { 
    . $PSScriptRoot\Setup.ps1
}

Describe "Get-XLFile" {
    Context "File exists" {
        BeforeAll {
            $path = Get-TestPath
            New-XLFile -Path $path
        }
        It "Should load the file" {
            $file = Get-XLFile -Path $path
            $file | Should -Not -Be $null
            $file.Package | Should -Not -Be $null
        }
    }
    
    Context "File does not exist" {
        BeforeAll {
            $path = Get-TestPath
        }
        It "Should throw" {
            { Get-XLFile -Path $path } | Should -Throw
        }
    }
}

