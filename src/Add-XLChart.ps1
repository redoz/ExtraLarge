function Add-XLChart {
[OutputType([XLChart])]
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true)]
    [XLSheet]$Sheet,
    [string]$Header,
    [OfficeOpenXml.Drawing.Chart.eChartType]$Type,
    [string]$XSeries,
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
    $worksheet = $Sheet.Worksheet
    $chart = $worksheet.Drawings.AddChart("Chart" + [Guid]::NewGuid().ToString('n'), $Type)
    $chart.Title.Text = $Header
    $chart.SetPosition($Row, 0, $Column, 0)
    $chart.SetSize($Width, $Height)

    # TODO this kinda sucks
    if ([bool]$Options['ShowPercent']) {
        $chart.DataLabel.ShowPercent= $true
    }

    if ([bool]$Options['ShowValue']) {
        $chart.DataLabel.ShowValue = $true
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
    
    if (-not [bool]$Options['XAxisBetweenTicks'] -and $chart.YAxis -is [OfficeOpenXml.Drawing.Chart.ExcelChartAxis]) {
        # TODO not sure why this has to be set on the YAxis, makes little sense
        $chart.YAxis.CrossBetween = [OfficeOpenXml.Drawing.Chart.eCrossBetween]::MidCat
    }
    
    $xlChart = [XLChart]::new($Sheet.Owner, $chart)
    
    if (-not [string]::IsNullOrEmpty($XSeries)) {
        $xlChart.XSeries = $XSeries
    }
    
    if ($With -ne $null) {
        $null = $xlChart | ForEach-Object -Process $With
    }
    if ($PassThru.IsPresent) {
        $Sheet;
    } elseif ($With -eq $null) {
        $xlChart;
    }
}
end {}
}

#TODO add parameter sets so Type can be typed but still optional
function Add-XLChartSeries {
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true)]
    [XLChart]$Chart,
    [object]$Header = $null,
    [string]$XSeries,
    [Parameter(Mandatory = $true)]
    [string]$YSeries,
    [object]$Type = $null,
    [switch]$PassThru = $false
)
begin{}
process{
    $excelChart = $Chart.Chart;
    
    if ($XSeries -eq '') {
        $XSeries = $Chart.XSeries;
    }
    
    if ($XSeries -eq '') {
        throw "XSeries was not provided and the chart doesn't have one specified."
    }
    
    if ($Type -eq $null) { 
        $series = $excelChart.Series.Add($YSeries, $XSeries);
    } else {
        $subChart = $excelChart.PlotArea.ChartTypes | Where-Object -FilterScript {$_.ChartType -eq [OfficeOpenXml.Drawing.Chart.eChartType]$Type}
        if ($subChart -eq $null) {
            $subChart = $excelChart.PlotArea.ChartTypes.Add([OfficeOpenXml.Drawing.Chart.eChartType]$Type);
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
