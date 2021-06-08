using System.IO;

public class XLFile
{
    public XLFile(OfficeOpenXml.ExcelPackage package) {
        this.Package = package;
    }
    public OfficeOpenXml.ExcelPackage Package {get; private set;}
    
    public void Save() {
        this.Package.Save();
    }

    public void SaveAs(FileInfo fileInfo)
    {
        this.Package.SaveAs(fileInfo);
    }
    
    public static implicit operator XLFile(OfficeOpenXml.ExcelPackage package) {
        return new XLFile(package);
    }
}