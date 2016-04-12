public abstract class XLBase {
    protected XLBase(OfficeOpenXml.ExcelPackage owner) {
        this.Owner = owner;
    }
    public bool HasOwner { get { return this.Owner != null; }}
    public OfficeOpenXml.ExcelPackage Owner {get; private set;}
}

public class XLFile
{
    public XLFile(OfficeOpenXml.ExcelPackage package) {
        this.Package = package;
    }
    public OfficeOpenXml.ExcelPackage Package {get; private set;}
    
    public void Save() {
        this.Package.Save();
    }
    
    public static implicit operator XLFile(OfficeOpenXml.ExcelPackage package) {
        return new XLFile(package);
    }
}

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

public class XLSheet : XLBase {
    public XLSheet(OfficeOpenXml.ExcelPackage owner, OfficeOpenXml.ExcelWorksheet worksheet) : base(owner) {
        this.Worksheet = worksheet;
    }
    
    public OfficeOpenXml.ExcelWorksheet Worksheet {get; private set;}
    
    public static implicit operator XLSheet(OfficeOpenXml.ExcelWorksheet sheet) {
        return new XLSheet(null, sheet);
    }
}