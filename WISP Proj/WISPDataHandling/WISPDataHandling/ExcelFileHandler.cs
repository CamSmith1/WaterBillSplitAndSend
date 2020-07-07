using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace WatercareDataHandling
{
    public class ExcelFileHandler
    {




        //Regex for string manipulation to remove white space
        static readonly Regex trimmer = new Regex(@"\s\s+");


        //Create an excel spreadsheet and populate it with data from Google Sheets
        //File path C\WatercareExcel
        //Files named WatercareInvoices_Month,day
        public void GenerateSpreadsheetData()
        {
            //************Directory Setup**************\\
            //Sets up directories if they do not exist
            DocumentUtils.createDirectoriesIfNotExist();

            //Archive files in Active folder 
            DocumentUtils.ArchiveActiveFiles();
            //*****************************************\\

            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Worksheet1");

                //Excel spreadsheet headers
                List<string[]> headerRow = new List<string[]>()
                    {
                      new string[] { "Transaction Amount", "Wastewater Fixed Charges","Tenant Charges", "Property Address", "Supplier Number", "Due Date" , "This Reading" , "Last Reading" ,"Paid By Owner", "Filename","Invoice Import Time" }
                    };

                string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";
                var worksheet = excel.Workbook.Worksheets["Worksheet1"];
                worksheet.Cells[headerRange].LoadFromArrays(headerRow);

                //Populate rows
                GoogleAPICalls gApi = new GoogleAPICalls();
                SpreadSheet sheet = gApi.GetSpreadsheetData();


                List<string[]> rows = new List<string[]>();

                double batchTotal = 0;

                //Extract fields from an individual row of the spreadsheet object feed it into string array and save array into an array list to be fed into excel file
                foreach (var row in sheet.Rows)
                {
                    double tenantCharges = Convert.ToDouble(row.InvoiceTransactionAmount) - Convert.ToDouble(row.WasteWaterFixedCharges);
                    double transAmount = Convert.ToDouble(row.InvoiceTransactionAmount);

                    bool ownerPaid = false;
                    //If Owner has paid the waterbill, only charge tenant
                    if (row.PaidByOwner != "Total amount to pay")
                    {
                        ownerPaid = true;
                        transAmount = 0.00;
                    }
                    string propertyAddress = DocumentUtils.formatAddress(row.PropertyAddress);

                    //Handle properties which can throw errors and need to be hardcoded. 
                    if(row.SupplierNumber == "5049839-01"){propertyAddress = "18A Heath";}
                    if (row.SupplierNumber == "5393154-01"){propertyAddress = "34A Raphoe";}

                    //Exceptions added 25/02/19
                    if (row.SupplierNumber == "5367534-02"){propertyAddress = "2/6 Sandgate";}
                    if (row.SupplierNumber == "4005685-03") {propertyAddress = "4/68 Jellicoe"; }
                    if (row.SupplierNumber == "5040359-01") { propertyAddress = "14/81 Shakespeare"; }
                    if (row.SupplierNumber == "5090510-02") { propertyAddress = "4/278A Birkdale"; }
                    if (row.PropertyAddress.Contains("station")) { propertyAddress = "62a Station"; }
                    if (row.PropertyAddress.Contains("Loloma")) { propertyAddress = "33 Loloma"; }
                    if (row.PropertyAddress.Contains("Canal")) { propertyAddress = "90a Canal"; }



                    //Verbena needs to be processed manually as its waterbil handles 1/25A and 2/25A in 1 single invoice
                    if (!propertyAddress.Contains("Verbena") || !propertyAddress.Contains("coronation"))//Doto create a collection of hardcoded propety names to be excluded
                    {

                        string[] rowData = new string[]
                        {  
                        transAmount.ToString(),
                        row.WasteWaterFixedCharges,
                        tenantCharges.ToString(),
                        propertyAddress,
                        row.SupplierNumber,
                        row.DueDate,
                        trimmer.Replace(row.ThisReading," "),
                        trimmer.Replace(row.LastReading," "),
                        ownerPaid.ToString(),
                        row.FileName,
                        row.importDate
                        };

                        if (!propertyAddress.Contains("Home")) { rowData[2] = "0"; }//If home street property set the tenant charge to 0 as owner pays all

                        batchTotal += transAmount;
                        rows.Add(rowData);
                    }
                 
                }

                //Load batchTotal Amount
                worksheet.Cells["L1"].LoadFromText(batchTotal.ToString());
                //Load data from row data into spreadsheeet starting cell 2 column 1
                worksheet.Cells[2, 1].LoadFromArrays(rows);



                //Todo if filename already exists as in if the process has been ran more than one time for the day handle the naming convention - 06/01/19
                FileInfo excelFile = new FileInfo(@"C:\WatercareInvoiceProcessing\ActiveSpreadsheet\WaterCareInvoice " + ".xlsx");
                excel.SaveAs(excelFile);
            }



        }


    }
}
