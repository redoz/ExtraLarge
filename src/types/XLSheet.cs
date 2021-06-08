public class XLSheet : XLBase {
    public XLSheet(OfficeOpenXml.ExcelPackage owner, OfficeOpenXml.ExcelWorksheet worksheet) : base(owner) {
        this.Worksheet = worksheet;
    }
    
    public string Name { get { return this.Worksheet.Name; } }
    
    public OfficeOpenXml.ExcelWorksheet Worksheet {get; private set;}
    
    public static implicit operator XLSheet(OfficeOpenXml.ExcelWorksheet sheet) {
        return new XLSheet(null, sheet);
    }
}