function Add-XLChart {
param(
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true)]
    [OfficeOpenXml.ExcelWorksheet]$Sheet,
    [string]$Header,
    [OfficeOpenXml.Drawing.Chart.eChartType]$Type,
    [int]$Row = 1,
    [int]$Column = 1,
    [int]$Width = 800,
    [int]$Height = 480,
    [Switch]$PassThru = $false,
    [Scriptblock]$With = $null

);
begin {}
process {
    $chart = $Sheet.Drawings.AddChart("Chart" + [Guid]::NewGuid().ToString('n'), $Type);
    $chart.Title.Text = $Header
    $chart.SetPosition($Row, 0, $Column, 0);
    $chart.SetSize($Width, $Height);

    if ($With -ne $null) {
        $null = $chart | ForEach-Object -Process $With
    }
    if ($PassThru.IsPresent) {
        $Sheet;
    } elseif ($With -eq $null) {
        $chart;
    }
}
end {}
}

#TODO add parameter sets so Type can be typed but still optional
function Add-XLChartSeries {
param(
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true)]
    [OfficeOpenXml.Drawing.Chart.ExcelChart]$Chart,
    [object]$Header = $null,
    [Parameter(Mandatory = $true)]
    [string]$XSeries,
    [Parameter(Mandatory = $true)]
    [string]$YSeries,
    [object]$Type = $null,
    [switch]$PassThru = $false
)
begin{}
process{
    if ($Type -eq $null) {
        $series = $Chart.Series.Add($YSeries, $XSeries);
    } else {
        $subChart = $Chart.PlotArea.ChartTypes | Where-Object -FilterScript {$_.ChartType -eq [OfficeOpenXml.Drawing.Chart.eChartType]$Type}
        if ($subChart -eq $null) {
            $subChart = $Chart.PlotArea.ChartTypes.Add([OfficeOpenXml.Drawing.Chart.eChartType]$Type);
        }
        $series = $subChart.Series.Add($YSeries, $XSeries);
    }
    if ($Header -ne $null) {
        $series.Header = [string]$Header;
    }
    if ($PassThru.IsPresent) {
        $Chart
    }
}
end{}
}