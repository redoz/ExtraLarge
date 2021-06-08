public class XLChart : XLBase {
    public XLChart(OfficeOpenXml.ExcelPackage owner, OfficeOpenXml.Drawing.Chart.ExcelChart chart) : base(owner) {
        this.Chart = chart;
    }
    public OfficeOpenXml.Drawing.Chart.ExcelChart Chart {get; private set;}
    
    public string XSeries {get; set;}
    
    public static implicit operator XLChart(OfficeOpenXml.Drawing.Chart.ExcelChart chart) {
        return new XLChart(null, chart);
    }
}