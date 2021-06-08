BeforeAll { 
    . $PSScriptRoot\Setup.ps1
}

Describe "New-XLFile" {
    Context "There are no worksheets added" {
        It "creates a default worksheet" {
            [string]$path = Get-TestPath
            New-XLFile -Path $path -Verbose | Should -BeNullOrEmpty
            $path | Should -Exist
        }
    }
    Context "PassThru is set" {
        It "should return an [XLFile]" {
            [string]$path = Get-TestPath
            New-XLFile -Path $path -PassThru | %{$_ -is [XLFile]} | Should -Be $true
            $path | Should -Exist
        }
    }
    
    Context "PassThru is not set" {
        It "should return nothing" {
            [string]$path = Get-TestPath
            New-XLFile -Path $path -PassThru:$false | Should -BeNullOrEmpty
            $path | Should -Exist
        }
    }
    
    Context "File exists" {
        It "should throw without -Force" {
            [string]$path = Get-TestPath
            Set-Content -Path $path -Value "" -Force
            $path | Should -Exist
            { New-XLFile -Path $path } | Should -Throw
            $path | Should -Exist
        }
        It "should overwrite with -Force" {
            [string]$path = Get-TestPath
            Set-Content -Path $path -Value @() -AsByteStream -Force
            $path | Should -Exist
            (Get-Item -Path $path).Length -eq 0 | Should -Be $true
            New-XLFile -Path $path -Force | Should -BeNullOrEmpty
            (Get-Item -Path $path).Length -gt 0 | Should -Be $true
        }
    }
}

