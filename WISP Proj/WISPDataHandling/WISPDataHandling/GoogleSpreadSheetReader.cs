using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;

public class GoogleSpreadSheetReader
{
    private readonly SheetsService _sheetService;
    public GoogleSpreadSheetReader(GoogleService googleService)
    {
        _sheetService = googleService.GetSheetsService();
        
    }

    public SpreadSheet GetSpreadSheet(string spreadSheetId, string range)
    {
        SpreadsheetsResource.ValuesResource.GetRequest request = _sheetService.Spreadsheets.Values.Get(spreadSheetId, range);

        ValueRange response = request.Execute();
        IList<IList<Object>> values = response.Values;
        var rows = new List<SpreadSheetRow>();

        //Used for Batch total method found in GoogleApiCalls - Batch total found on J1 in spreadsheet
       if(values.Count == 1)
        {
            SpreadSheetRow row;
            row = new SpreadSheetRow(values[0]);
            rows.Add(row);

        }

        for (int i = 1; i < values.Count; i++)
        {
            SpreadSheetRow row;

 
                 row = new SpreadSheetRow(values[i]); 
                 rows.Add(row);


        }
        var headerRow = new SpreadSheetRow(values[0]);
        var spreadSheet = new SpreadSheet();
        spreadSheet.HeaderRow = headerRow;
        spreadSheet.Rows = new List<SpreadSheetRow>();
        spreadSheet.Rows.AddRange(rows);
        return spreadSheet;
    }

    




}