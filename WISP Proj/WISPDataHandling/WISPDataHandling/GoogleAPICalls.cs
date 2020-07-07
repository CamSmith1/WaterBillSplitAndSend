using System;
using System.Collections.Generic;
using System.Text;
using Google.Apis.Sheets.v4;
using Newtonsoft.Json;

namespace WatercareDataHandling
{
   public class GoogleAPICalls
    {
        public GoogleAPICalls()
        {

        }
       public  void printWatercareParsedData()
        {

            Console.WriteLine("Pulling data from Google Sheet");
            var googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\GoogleSecret\\GoogleSecret.json";
            var applicationName = "Warercare Invoice Parsed Data";
            string[] scopes = { SheetsService.Scope.SpreadsheetsReadonly };

            var googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

            var spreadSheetId = "1mKAlHmbH4Cfx_noGlmOIGbmhbKbjbVNxhmn-dQ-Z3x8";
            var range = "$A:$J";

            var reader = new GoogleSpreadSheetReader(googleService);

            var spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);
            foreach(var x in spreadSheet.Rows)
            {
                

            }

            Console.WriteLine(JsonConvert.SerializeObject(spreadSheet.Rows));


          
            Console.Read();

        }

        public SpreadSheet GetSpreadsheetData()
        {
       
            var googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\GoogleSecret\\GoogleSecret.json";
            var applicationName = "Warercare Invoice Parsed Data";
            string[] scopes = { SheetsService.Scope.SpreadsheetsReadonly };

            var googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

            var spreadSheetId = "1mKAlHmbH4Cfx_noGlmOIGbmhbKbjbVNxhmn-dQ-Z3x8";
            var range = "$A:$J";

            var reader = new GoogleSpreadSheetReader(googleService);

            return reader.GetSpreadSheet(spreadSheetId, range);
            


        }





    }
}
