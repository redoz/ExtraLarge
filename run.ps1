Import-Module -Name c:\dev\XExcel\ExtraLarge\ExtraLarge.psm1

$raw = @"
A,B,C,Date
2,1,1,2016-03-29
5,10,1,2016-03-29
"@

$data = ConvertFrom-Csv -InputObject $raw


New-XLFile -Path c:\temp\out.xlsx `
           -Sheets @(
                  New-XLSheet -Name "Sheet 1"`
                              -Tables @(
                                  New-XLTable -Row 4 -Column 5 -Data $data
                                  )
              )