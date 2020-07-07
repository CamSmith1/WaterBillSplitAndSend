using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using System.Data;
using WatercareDataHandling;
using MahApps.Metro.Controls;
using System.Diagnostics;
using System.Threading;

namespace GUIWISPDataHandling
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        public MainWindow()
        {
            InitializeComponent();
        }


        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {

            
                        //Generate a new Excel Spreadsheet populated with most current data from Google Sheets
                        ExcelFileHandler ExcelHandler = new ExcelFileHandler();
                        ExcelHandler.GenerateSpreadsheetData();

                        string FileDirectory = @"C:\WatercareInvoiceProcessing\ActiveSpreadsheet\";

                        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                        //Static File From Base Path...........
                        Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(FileDirectory + "WaterCareInvoice .xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                        ////Dynamic File Using Uploader...........
                        //Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(txtFilePath.Text.ToString(), 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(1); ;
                        Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;

                        string strCellData = "";
                        string douCellData;
                        int rowCnt = 0;
                        int colCnt = 0;

                        DataTable dt = new DataTable();
                        for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                        {
                            string strColumn = "";
                            strColumn = Convert.ToString((excelRange.Cells[1, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2);
                            dt.Columns.Add(strColumn, typeof(string));
                        }

                        for (rowCnt = 2; rowCnt <= excelRange.Rows.Count; rowCnt++)
                        {
                            string strData = "";
                            for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                            {
                                try
                                {
                                    strCellData = (string)(excelRange.Cells[rowCnt, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                                    strData += strCellData + "|";
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex);
                                    douCellData = (excelRange.Cells[rowCnt, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2.ToString() ;
                                    strData += douCellData.ToString() + "|";
                                }
                            }
                            strData = strData.Remove(strData.Length - 1, 1);
                            dt.Rows.Add(strData.Split('|'));
                        }

                        dtGrid.Visibility = Visibility.Visible;
                        dtGrid.ItemsSource = dt.DefaultView;

                        excelBook.Close(true, null, null);
                        excelApp.Quit();

                        //set ui element visability
                        btnProcess.Visibility = Visibility.Visible;
            
        }


        private void btnClick_ProcessData(object sender, RoutedEventArgs e)
        {
            //Execute Bot to input data into QPMS
            QPMSBot bot = new QPMSBot();
            bot.startBot();         
        }

        private void btnClick_GoToGoogleSheet(object sender, RoutedEventArgs e)
        {
            string sMessageBoxText = "Any changes you make to the Google Sheet containing invoice batch data, will require you to click the Generate Data button again to update the data before processing";
            string sCaption = "QPMS WISP";

            MessageBoxButton btnMessageBox = MessageBoxButton.YesNo;
            MessageBoxImage icnMessageBox = MessageBoxImage.Warning;

            MessageBoxResult rsltMessageBox = MessageBox.Show(sMessageBoxText, sCaption, btnMessageBox, icnMessageBox);

            switch (rsltMessageBox)
            {
                case MessageBoxResult.Yes:
                    Process.Start("chrome.exe", "https://docs.google.com/spreadsheets/d/1mKAlHmbH4Cfx_noGlmOIGbmhbKbjbVNxhmn-dQ-Z3x8/edit#gid=0");
                    break;
            }
        }
        private void btnClick_AddInvoice(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Outlook.Application oApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook._MailItem oMailItem = (Microsoft.Office.Interop.Outlook._MailItem)oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
            oMailItem.To = "tzecfmuwvbrm@in.docparser.com";
            oMailItem.Body = "Attach 1 or many waterbill pdf invoices to this email, wait 1 - 2 minutes and click the generate data button in QPMS WISP to view and process the data ";
            
            oMailItem.Display(true); 
        }

        private void btnClick_DisplayHelpMessage(object sender, RoutedEventArgs e)
        {

            string sMessageBoxText = "How To Use QPMS WISP \n\n" +
                "1) Click Generate Data button to pull the latest processed Watercare Invoice data \n\n" +
                "2) If no data is displayed or you want to add invoices to process click the Add Invoices button, which will open an outlook email, attach the Watercare Invoices.pdf that you want processed and send the email \n\n" +
                "3) If you have just send an email with attached invoices, click the Generate Data button again to pull the updated data. \n\n" +
                "4) When you have confirmed that the data is correct and ready to be procssed into a batch click the button Process Batch \n\n" +
                "5) Process Batch will cause a UIPath popup to display, when it opens click the start button for QPMSWaterCareBot to begin processing and do not click on the screen until the batch has been processed"
                ;
            string sCaption = "Help";

            MessageBoxButton btnMessageBox = MessageBoxButton.OK;
            MessageBoxImage icnMessageBox = MessageBoxImage.Warning;

            MessageBoxResult rsltMessageBox = MessageBox.Show(sMessageBoxText, sCaption, btnMessageBox, icnMessageBox);

            switch (rsltMessageBox)
            {
                case MessageBoxResult.OK:
                   //Do nothing box will close
                    break;
            }
        }



    }
}

