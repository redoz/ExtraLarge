Import-Module -Name "$PSScriptRoot\src\ExtraLarge.psm1" -Force

$raw = @"
A,B,C,Date
2,1,1,2016-03-29
5,10,1,2016-03-29
"@

$data = ConvertFrom-Csv -InputObject $raw
Remove-Item C:\temp\out3.xlsx -Force

New-XLFile -Path c:\temp\out3.xlsx -PassThru |
    Add-XLSheet -Name 'Sheet 1' |
        Add-XLTable -Name Table1 -Data $data -Columns A,B,C,Date -PassThru |
        Add-XLTable -Name Table2 -Data $data[0] -Columns A,B,C,Date  


return;


