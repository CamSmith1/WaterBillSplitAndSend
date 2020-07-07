using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace WatercareDataHandling
{
    public class DocumentUtils
    {

        static readonly string activeSpreadsheetPath = @"C:\\WatercareInvoiceProcessing\ActiveSpreadsheet\";
        static readonly string archiveSpreadsheetPath = @"C:\\WatercareInvoiceProcessing\SpreadsheetArchive\";
        static readonly string mainDirectory = @"C:\\WatercareInvoiceProcessing";

        //Takes a property address and returns the first 2 words from the string
        public static string formatAddress(string address)
        {
            var WordsArray = address.Split();
            string addressSubString = WordsArray[0] + ' ' + WordsArray[1];
            return addressSubString;

        }

        //If the correct directories do not initially exist create them
        public static void createDirectoriesIfNotExist()
        {

            //Create Main work directory if not exists
            if (!Directory.Exists(mainDirectory))
            {
                Directory.CreateDirectory(mainDirectory);
                Console.WriteLine("Directory " + mainDirectory + " has been created");
            }
            //Create Active Spreadsheet Path if not exists
            if (!Directory.Exists(activeSpreadsheetPath))
            {
                Directory.CreateDirectory(activeSpreadsheetPath);
                Console.WriteLine("Directory " + activeSpreadsheetPath + " has been created");
            }

            //Create Archive Spreadsheet Path is not exists
            if (!Directory.Exists(archiveSpreadsheetPath))
            {
                Directory.CreateDirectory(archiveSpreadsheetPath);
                Console.WriteLine("Directory " + archiveSpreadsheetPath + " has been created");
            }
        }


        //Moves all files from Active directory to Archive

        public static void ArchiveActiveFiles()
        {
            DirectoryInfo dir = new DirectoryInfo(activeSpreadsheetPath);
            FileInfo[] files = dir.GetFiles();

            foreach(FileInfo file in files)
            {
                if(File.Exists(archiveSpreadsheetPath + file.Name))
                {
                    File.Delete(archiveSpreadsheetPath + file.Name);
                }
               // file.MoveTo(archiveSpreadsheetPath + file.Name);
                File.Move(file.FullName, archiveSpreadsheetPath + "WaterCareInvoices " + DateTime.Now.ToString("yyyy - dd - M--HH - mm - ss") + ".xlsx");
               


                //WaterCareInvoices " + DateTime.Now.ToString("yyyy - dd - M--HH - mm - ss") + ".xlsx"
            }
        }









    }
}
