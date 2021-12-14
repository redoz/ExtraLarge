# ExtraLarge
[![Build status](https://ci.appveyor.com/api/projects/status/mujcrkjo8hyjpy3r/branch/master?svg=true)](https://ci.appveyor.com/project/redoz/extralarge/branch/master)
[![Coverage Status](https://coveralls.io/repos/github/redoz/ExtraLarge/badge.svg?branch=master)](https://coveralls.io/github/redoz/ExtraLarge?branch=master)

*This project is now archived, the licensing change in epplus means there is little incentive to maintain this further.*

Create Excel files using PowerShell, inspired by https://github.com/dfinke/ImportExcel, made possible by http://epplus.codeplex.com/

## Sample
```PowerShell
Import-Module -Name ExtraLarge

$raw = @"
A,B,C,Date
2,1,0.5,2016-03-29
5,10,,2016-03-30
"@

$data = ConvertFrom-Csv -InputObject $raw

New-XLFile -Path out.xlsx -PassThru |
    Add-XLSheet -Name 'Sheet 1' |
        Add-XLTable -Name Table1 -Data $data -Columns @{Name='A';Type=[int]},@{Name='B';Type=[int]},@{Name='C';Type=[float];NumberFormat='Percent'},@{Name='Date';Type=[DateTime]} -PassThru |
        Add-XLTable -Name Table2 -Data $data -Columns  @{Name='A';Type=[int]},@{Name='B';Type=[int]},@{Name='C';Type=[float];Default=30},@{Name='D';Type=[float];Default=99},@{Name='Date';Type=[DateTime];NumberFormat=[XLNumberFormat]::Date} -PassThru |
        Add-XLChart -Header "Chart 1" -Type "Line" -Column 6 -XSeries "Table2[Date]" -With { $_ | 
                                                                Add-XLChartSeries -YSeries "Table2[A]" -PassThru | 
                                                                Add-XLChartSeries -YSeries "Table2[B]" -Type AreaStacked -PassThru |
                                                                Add-XLChartSeries -YSeries "Table2[C]" -Type AreaStacked
                                                            }
```

## Todo
* Add more tests
* Add parameter sets per chart type to Add-XLChart to enable type safe options
* Add documentation
* Add samples
* Add Set-XLRange
* Add Set-XLValue/Formula
* Add Show/Hide-XLRange/Row/Column
* Add pivot table/chart support
* Add Get-* functions
* Find shorter/cleaner syntax for Add-XLTable -Columns parameter
* ...
