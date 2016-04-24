using System.Management.Automation;
using System.Collections.Generic;
using System.Collections;
using System;
using System.Text.RegularExpressions;

public enum XLNumberFormat {
    Text,
    Date,
    General,
    Percent,
    DateTime,
    Time
}

public enum XLTotalsFunction {
    Average=101,
    Count=102,
    CountA=103,
    Max=104,
    Min=105,
    Product=106,
    Stdev=107,
    StdevP=108,
    Sum=109,
    Var=110,
    VarP=111
}

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
    
    public string Name { get { return this.Worksheet.Name; } }
    
    public OfficeOpenXml.ExcelWorksheet Worksheet {get; private set;}
    
    public static implicit operator XLSheet(OfficeOpenXml.ExcelWorksheet sheet) {
        return new XLSheet(null, sheet);
    }
}

public class XLRange : XLBase, IEnumerable<PSObject> {
    
    private static readonly Regex _dateTimeFormatMatch = new Regex("[ymdhs]|AM/PM", System.Text.RegularExpressions.RegexOptions.Compiled);
    
    public XLRange(OfficeOpenXml.ExcelPackage owner, OfficeOpenXml.ExcelRangeBase range) : base(owner) {
        this.Range = range;
    }
    
    public string Name { 
        get {
            if (this.Range is OfficeOpenXml.ExcelNamedRange) {
                return ((OfficeOpenXml.ExcelNamedRange)this.Range).Name;
            } else {
                return null;
            }
        }
    }
         
    
    public string Address { get { return this.Range.FullAddress; } }
    
    public OfficeOpenXml.ExcelRangeBase Range {get; private set;}
    
    public string[] Headers {get; set;}
    
    public bool HasHeaders { get;set; }
    
    public IEnumerator<PSObject> GetEnumerator() {
        // TODO this is for "Data" only, should have properties indicating what format was requested
        // TODO include Transpose property
        int rowOffset = this.Range.Start.Row;
        int columnOffset = this.Range.Start.Column;

        string[] columns;
        if (this.Headers != null)
            columns = this.Headers;
        else if (this.HasHeaders) {
            columns = new string[this.Range.Columns];
            for (int i = 0; i < columns.Length; i++)
                columns[i] = this.Range.Worksheet.Cells[rowOffset, columnOffset + i].Text;
        }
        else {
            columns = new string[this.Range.Columns];
            for (int i = 0; i < columns.Length; i++)
                columns[i] = OfficeOpenXml.ExcelCellAddress.GetColumnLetter(columnOffset + i);
        }
        
        for (int rowNum = this.HasHeaders ? 1 : 0 ; rowNum < this.Range.Rows; rowNum++)
        {
            PSObject row = new PSObject();
            
            for (int colNum = 0; colNum < columns.Length; colNum++)
            {
                var cell = this.Range.Worksheet.Cells[rowOffset + rowNum, columnOffset + colNum];
                
                // this is pretty horrible, but doesn't seem to be a better way
                object cellValue;
                if (cell != null) {
                    if (cell.Style.Numberformat.BuildIn) {
                        switch (cell.Style.Numberformat.NumFmtID) {
                            case 14:
                            case 15:
                            case 16:
                            case 17:
                            case 18:
                            case 19:
                            case 20:
                            case 21:
                            case 22:
                            case 45:
                            case 46:
                            case 47:
                            case 27:
                            case 30:
                            case 36:
                            case 50:
                            case 57:
                               if (cell.Value is double)
                                    cellValue = DateTime.FromOADate((double)cell.Value);
                                else
                                    cellValue = cell.Value;
                                break;
                            default:
                                cellValue = cell.Value;
                                break;
                        }   
                    } else if (cell.Value is double && _dateTimeFormatMatch.IsMatch(cell.Style.Numberformat.Format)) {
                        cellValue = DateTime.FromOADate((double)cell.Value);
                    } else
                        cellValue = cell.Value;    
                } else {
                    cellValue = null;
                }
                
                row.Members.Add(new PSNoteProperty(columns[colNum], cellValue));
            }
            
            yield return row;
        }
    }
    
    IEnumerator System.Collections.IEnumerable.GetEnumerator() {
        return this.GetEnumerator();
    }
    
    public static implicit operator XLRange(OfficeOpenXml.ExcelRange range) {
        return new XLRange(null, range);
    }
}