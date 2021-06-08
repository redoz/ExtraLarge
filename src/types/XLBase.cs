public abstract class XLBase {
    protected XLBase(OfficeOpenXml.ExcelPackage owner) {
        this.Owner = owner;
    }
    public bool HasOwner { get { return this.Owner != null; }}
    public OfficeOpenXml.ExcelPackage Owner {get; private set;}
}