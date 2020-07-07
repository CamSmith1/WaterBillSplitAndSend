using Google.Apis.Sheets.v4;
using Newtonsoft.Json;
using System;
using WatercareDataHandling;

public class Program
{
    static void Main(string[] args)
    {

        ExcelFileHandler ExcelHandler = new ExcelFileHandler();
        //GenerateSpreadSheetData pulls data from the google sheet in Json format, creates a new excel spreadsheet, saves file into activesheets directory and archives files that preexist in the activesheets dir
        ExcelHandler.GenerateSpreadsheetData();
    }




}

