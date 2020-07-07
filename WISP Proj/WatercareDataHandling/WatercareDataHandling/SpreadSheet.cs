using System;
using System.Collections.Generic;
using WatercareDataHandling;

public class SpreadSheet
{
    public SpreadSheetRow HeaderRow { get; set; }
    public List<SpreadSheetRow> Rows { get; set; }

    public double BatchTotal { get; set; }
}
public class SpreadSheetRow
{

   
    private readonly IList<Object> _values;
    public SpreadSheetRow(IList<Object> values)
    {
        _values = values;
    }
    DocumentUtils dat = new DocumentUtils();
    public string InvoiceTransactionAmount => _getValue(0);

    public string WasteWaterFixedCharges => _getValue(1);
    public  string PropertyAddress => _getValue(2);
     
    public string SupplierNumber => _getValue(3);
    
    public string DueDate => _getValue(4);
    public string ThisReading => _getValue(5);
    public string LastReading => _getValue(6);
    public string FileName => _getValue(8);
    public string importDate => _getValue(9);

    public double tenantCharges;
    public String PaidByOwner => _getValue(7);
    public bool OwnerPaid;

    private string _getValue(int columnIndex)
    {
        try
        {
            var s = _values[columnIndex].ToString();
            return s;
        }
        catch (Exception ex)
        {
            return String.Empty;
        }
    }
}