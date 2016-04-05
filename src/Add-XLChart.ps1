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
    [System.Collections.IDictionary]$Options = @{},
    [Switch]$PassThru = $false,
    [Scriptblock]$With = $null

);
begin {}
process {
    $chart = $Sheet.Drawings.AddChart("Chart" + [Guid]::NewGuid().ToString('n'), $Type);
    $chart.Title.Text = $Header
    $chart.SetPosition($Row, 0, $Column, 0);
    $chart.SetSize($Width, $Height);

    # TODO this kinda sucks
    if ([bool]$Options['ShowPercent']) {
        $chart.DataLabel.ShowPercent= $true;
    }

    if ([bool]$Options['ShowValue']) {
        $chart.DataLabel.ShowValue = $true;
    }

    if ([bool]$Options['NoLegend']) {
        $chart.Legend.Remove();
    }

    if ([bool]$Options['HideYAxis']) {
        $chart.YAxis.Deleted = $true;
        $chart.YAxis.MajorTickMark = [OfficeOpenXml.Drawing.Chart.eAxisTickMark]::None;
        $chart.YAxis.MinorTickMark = [OfficeOpenXml.Drawing.Chart.eAxisTickMark]::None;

        # TODO this deletes all majorGridLines which I'm pretty sure is not correct
        $chartXml = $chart.ChartXml;
        $nsuri = $chartXml.DocumentElement.NamespaceURI;
        $nsm = [System.Xml.XmlNamespaceManager]::new($chartXml.NameTable);
        $nsm.AddNamespace("c", $nsuri);

        $gridLines = $chartXml.SelectNodes('//c:majorGridlines', $nsm)
        $null = $gridLines | ForEach-Object -Process {$_.ParentNode.RemoveChild($_);}

    }
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