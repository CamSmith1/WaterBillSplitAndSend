using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace GUIWISPDataHandling
{
    class QPMSBot
    {

        IWebDriver driver;
        String ChromeDriverDirectory = "";
        IWebElement userNameInput, pwdInput, loginBtn, addNewBatchBtn ;

        //Global Excel Spreadsheet stuff
        Microsoft.Office.Interop.Excel.Application appObj;
        Workbook objBook;
        Worksheet objWorkSheet;
        Range range;

        public void startBot()
        {
            startBrowser();
            login("*******", "******");
            navigateToTransactions();
            createNewBatch();
            //Inputs all the details to create a new batch
            InputTransactionDetails();
            //Logic to input transactions data from spreadsheet
            InputTransactions();



        }

        //Open google chrome and navigate to QPMS
        public void startBrowser()
        {
            string workingDirectory = Environment.CurrentDirectory;
            string projectDirectory = Directory.GetParent(workingDirectory).Parent.FullName;
            ChromeDriverDirectory = projectDirectory + @"\assets\";
            driver = new ChromeDriver(ChromeDriverDirectory);
            driver.Url = "https://www.quinovic.co.nz/QPMSNet/Login/Login.aspx";

        }


        public void InputTransactions()
        {
            selectSupplier();
            Range usedRange = objWorkSheet.UsedRange;
            bool skipFirstRow = false;
            foreach(Range row in usedRange.Rows)
            {
                if(skipFirstRow == false)
                {
                    skipFirstRow = true;
                }
                else
                {
                    String[] rowData = new string[row.Columns.Count];
                    for (int i = 0; i < row.Columns.Count; i++)
                    {
                        rowData[i] = Convert.ToString(row.Cells[1, i + 1].Value2);
                    }
                    Console.WriteLine(rowData);
                    string transactionAmount = rowData[0];
                    string wastewaterFixedCharges = rowData[1];
                    string propAddress = rowData[3];
                    string supplierNo = rowData[4];
                    string thisReading = rowData[6];
                    string lastReading = rowData[7];
                    string paidByOwner = rowData[8];

                    //Add batch
                    processSingleBatch(propAddress, transactionAmount, supplierNo, wastewaterFixedCharges, thisReading, lastReading, paidByOwner);
                }
            }

        }


        //Login to QPMS when given a username and password
        public void login(string username, string password)
        {
            string unameXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table[@id='LoginTable']/tbody/tr[1]/td[2]/input[@id='ctl00_ContentPlaceHolder1_txtUserName']";
            string pwdXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table[@id='LoginTable']/tbody/tr[2]/td[2]/input[@id='ctl00_ContentPlaceHolder1_txtPassword']";
            string loginBtnXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table[@id='LoginTable']/tbody/tr[3]/td[2]/input[@id='ctl00_ContentPlaceHolder1_cmdLogin']";
            userNameInput = driver.FindElement(By.XPath(unameXpath));
            pwdInput = driver.FindElement(By.XPath(pwdXpath));
            loginBtn = driver.FindElement(By.XPath(loginBtnXpath));
            userNameInput.SendKeys(username);
            pwdInput.SendKeys(password);
            loginBtn.Click();  
        }

        public void navigateToTransactions()
        {
            driver.Url = "https://www.quinovic.co.nz/QPMSNet/Transactions/MaintainBatches.aspx";
        }

        public void createNewBatch()
        {
            string AddNewBtnXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table/tbody/tr/td[1]/table[2]/tbody/tr/td/input[@id='ctl00_ContentPlaceHolder1_TransactionBatchesListEdit1_MaintainBatchesList_btnAddNew']";
            addNewBatchBtn = driver.FindElement(By.XPath(AddNewBtnXpath));
            addNewBatchBtn.Click();
        }
        //Setup transaction details page for the batch. This method finishes at the Input transactions data page
        public void InputTransactionDetails()
        {

            string TransTypeIDXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table/tbody/tr/td[2]/table/tbody/tr[4]/td[2]/select[@id='ctl00_ContentPlaceHolder1_TransactionBatchesListEdit1_MaintainBatchesEdit_cmbTransTypeId']";
            string BatchTotalXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table/tbody/tr/td[2]/table/tbody/tr[5]/td[2]/input[@id='ctl00_ContentPlaceHolder1_TransactionBatchesListEdit1_MaintainBatchesEdit_txtBatchTotal']";
            string BatchDescriptionXPath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table/tbody/tr/td[2]/table/tbody/tr[6]/td[2]/input[@id='ctl00_ContentPlaceHolder1_TransactionBatchesListEdit1_MaintainBatchesEdit_txtBatchDescription']";
            string addBatchBtnXPath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table/tbody/tr/td[2]/table/tbody/tr[15]/td/input[@id='ctl00_ContentPlaceHolder1_TransactionBatchesListEdit1_MaintainBatchesEdit_btnAdd']";
            string transactionsXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table/tbody/tr/td[1]/table[@id='ctl00_ContentPlaceHolder1_TransactionBatchesListEdit1_MaintainBatchesList_dataGrid']/tbody/tr[@class='datagrid-itemstyle']/td[9]/a[@id='ctl00_ContentPlaceHolder1_TransactionBatchesListEdit1_MaintainBatchesList_dataGrid_ctl03_lnkViewTransactions']";
          
            //Select Water Rates dropdown for Transaction Type
            IWebElement element = driver.FindElement(By.XPath(TransTypeIDXpath));
            SelectElement select = new SelectElement(element);
            select.SelectByText("Water Rates");

            //Input Transaction Amount
            IWebElement transactionAmount = driver.FindElement(By.XPath(BatchTotalXpath));
            transactionAmount.Clear();
            transactionAmount.SendKeys(getTransactionAmmount());

            //Input Batch Description
            IWebElement batchDescriptionInput = driver.FindElement(By.XPath(BatchDescriptionXPath));
            batchDescriptionInput.SendKeys("Water Rates");

            //Add batch
            IWebElement addBatchBtn = driver.FindElement(By.XPath(addBatchBtnXPath));
            addBatchBtn.Click();

            //Click the transactions button to take the screen to transaction input page
            IWebElement transactionsBtn = driver.FindElement(By.XPath(transactionsXpath));
            transactionsBtn.Click();

        }

        //Input data for transactions

     

        //*****************************************Input Data for Transactions***********************************************


        //When given all the parameters from single row in spreadsheet process a single batch
        public void processSingleBatch(string address, string transAmt,string supplierInvNo , string FixedCharges, string thisReading,  string LastReading, string paidByOwner )
        {
            string TransAmtXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table/tbody/tr/td[2]/table/tbody/tr[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_trTransAmt']/td[2]/input[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_txtGrossTransAmt']";
            string commentsXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table/tbody/tr/td[2]/table/tbody/tr[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_trComments']/td[2]/textarea[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_txtComments']";
            string TenantChargeXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table/tbody/tr/td[2]/table/tbody/tr[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_trTenantCharge']/td[2]/input[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_txtTenantWaterUsage']";
            string validateBtnXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table/tbody/tr/td[2]/table/tbody/tr[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_trButtons']/td/input[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_btnValidate']";
            string addBtnXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table/tbody/tr/td[2]/table/tbody/tr[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_trButtons']/td/input[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_btnAdd']";
            string supplierInvNoXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table/tbody/tr/td[2]/table/tbody/tr[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_trSupplierInvNo']/td[2]/input[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_txtSupplierInvNo']";

            string tenancyIDXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table/tbody/tr/td[2]/table/tbody/tr[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_trTenancyId']/td[2]/select[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_cmbTenancyId']";
            


            IWebElement transAmtInput = driver.FindElement(By.XPath(TransAmtXpath));
            transAmtInput.Clear();
            transAmtInput.SendKeys(transAmt);

            bool paidByOwnerBoolean = bool.Parse(paidByOwner);
            IWebElement tenantChargesInput = driver.FindElement(By.XPath(TenantChargeXpath));
            //If not paid by owner tenant charges is transaction amount minus fixed charges
            if(!paidByOwnerBoolean)
            {
                double tenantChargeNo = double.Parse(transAmt) - double.Parse(FixedCharges);
                tenantChargesInput.Clear();
                tenantChargesInput.SendKeys(tenantChargeNo.ToString());
            }

            IWebElement supplierInvNoInput = driver.FindElement(By.XPath(supplierInvNoXpath));
            supplierInvNoInput.Clear();
            supplierInvNoInput.SendKeys(supplierInvNo);

            //IWebElement tenancyDropDown = driver.FindElement(By.XPath(tenancyIDXpath));
            //tenancyDropDown.SendKeys(Keys.Down);
  

            //Set the NS number for the address
            searchAddressInput(address);

            // Select the last value in the dropdown
            SelectElement tenancyDropDownList = new SelectElement(driver.FindElement(By.XPath(tenancyIDXpath)));
            int x = tenancyDropDownList.Options.Count;
            if (x > 0)
                tenancyDropDownList.SelectByIndex(x - 1);

            //Input Comments
            IWebElement comments = driver.FindElement(By.XPath(commentsXpath));
            string CommentText = "Water rates to: "+ returnDate(thisReading) + System.Environment.NewLine + "This reading:- "+ getReading(thisReading)+ System.Environment.NewLine + "Last Reading:- "+ getReading(LastReading) + System.Environment.NewLine + "Units:- "+ getUnitsUsed(thisReading, LastReading)+ " Units";
            comments.Clear();
            comments.SendKeys(CommentText);

            //Add transaction
            IWebElement validateBtn = driver.FindElement(By.XPath(validateBtnXpath));
            validateBtn.Click();
            IWebElement addBtn = driver.FindElement(By.XPath(addBtnXpath));
            addBtn.Click();
        }

        //*********************************String Utils************************************************
        //string manipulation utils getting comments

        //When given a Reading regex to return the date
        public string returnDate(string text)
        {
            Regex regex = new Regex(@"\d{2}-\w{3}-\d{2}");
            Match match = regex.Match(text);

            if(match.Success)
            {
                return match.Groups[0].Value;
            }
            return null;
        }

        //Substring reading to get value after the date
        public string getReading(string text)
        {
            String date = returnDate(text);
            return text.Substring(text.IndexOf(date) + date.Length);
        }

        //When given last reading and this reading return the difference in units used
        public double getUnitsUsed(string thisReading, string lastReading)
        {
            return getUnits(thisReading) - getUnits(lastReading);
        }

        public double getUnits(string reading)
        {
            string s = getReading(reading);
            string number = s.Split(' ')[1]; ;
            return double.Parse(number);
        }

        //*********************************************************************************






        //************************************UI Utils*******************************
        //Select Supplier
        public void selectSupplier()
        {
            string searchSupplierXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table/tbody/tr/td[2]/table/tbody/tr[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_trSupplierId']/td[2]/input[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_btnLookupSupplier']";
            string searchSupplierNameXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table/tbody/tr/td[2]/div/span[3]/div[@class='ss']/table[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_LookupSupplier1_dlgLookup-table']/tbody/tr[@class='dialog_body_row']/td[@class='dialog_body_cell']/div/div[@class='dialog_content']/div[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_LookupSupplier1_dlgLookup_Panel1']/input[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_LookupSupplier1_dlgLookup_txtSupplierName']";
            string searchBtnXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table/tbody/tr/td[2]/div/span[3]/div[@class='ss']/table[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_LookupSupplier1_dlgLookup-table']/tbody/tr[@class='dialog_body_row']/td[@class='dialog_body_cell']/div/div[@class='dialog_content']/div[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_LookupSupplier1_dlgLookup_Panel1']/input[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_LookupSupplier1_dlgLookup_btnSearch']";
            string okBtnXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table/tbody/tr/td[2]/div/span[3]/div[@class='ss']/table[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_LookupSupplier1_dlgLookup-table']/tbody/tr[@class='dialog_body_row']/td[@class='dialog_body_cell']/div/div[@class='dialog_content']/div[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_LookupSupplier1_dlgLookup_Panel2']/input[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_LookupSupplier1_dlgLookup_btnDialogOK']";

            IWebElement searchSupplierBtn = driver.FindElement(By.XPath(searchSupplierXpath));
            searchSupplierBtn.Click();
            IWebElement searchSupplierSearchBox = driver.FindElement(By.XPath(searchSupplierNameXpath));
            searchSupplierSearchBox.SendKeys("Watercare");
            IWebElement searchBtn = driver.FindElement(By.XPath(searchBtnXpath));
            searchBtn.Click();
            IWebElement okBtn = driver.FindElement(By.XPath(okBtnXpath));
            okBtn.Click();
        }

        public void searchAddressInput(string address)
        {
            string searchBtnXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table/tbody/tr/td[2]/table/tbody/tr[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_trPropertyId']/td[2]/span/input[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_txtPropertyId_btn']";
            string inputBoxXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table/tbody/tr/td[2]/table/tbody/tr[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_trPropertyId']/td[2]/span/div[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_txtPropertyId_up1']/span/div[@class='ss']/table[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_txtPropertyId_lp1_dlgLookup-table']/tbody/tr[@class='dialog_body_row']/td[@class='dialog_body_cell']/div/div[@class='dialog_content']/div[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_txtPropertyId_lp1_dlgLookup_Panel1']/input[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_txtPropertyId_lp1_dlgLookup_txtPropertyAddress']";
            string addrSearchBtnXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table/tbody/tr/td[2]/table/tbody/tr[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_trPropertyId']/td[2]/span/div[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_txtPropertyId_up1']/span/div[@class='ss']/table[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_txtPropertyId_lp1_dlgLookup-table']/tbody/tr[@class='dialog_body_row']/td[@class='dialog_body_cell']/div/div[@class='dialog_content']/div[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_txtPropertyId_lp1_dlgLookup_Panel1']/input[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_txtPropertyId_lp1_dlgLookup_btnSearch']";
            string okBtnXpath = "/html/body/form[@id='aspnetForm']/div[@class='master_outer_div']/div[@class='round_bottomleft']/div[@class='round_bottomright']/div[@class='round_topleft']/div[@class='round_topright']/div/div[@class='master_inner_div']/div[@class='Content_Div']/table/tbody/tr/td[2]/table/tbody/tr[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_trPropertyId']/td[2]/span/div[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_txtPropertyId_up1']/span/div[@class='ss']/table[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_txtPropertyId_lp1_dlgLookup-table']/tbody/tr[@class='dialog_body_row']/td[@class='dialog_body_cell']/div/div[@class='dialog_content']/div[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_txtPropertyId_lp1_dlgLookup_Panel2']/input[@id='ctl00_ContentPlaceHolder1_MaintainBatchesTransactionsListEdit1_MaintainBatchesTransactionsEdit1_txtPropertyId_lp1_dlgLookup_btnDialogOK']";


            IWebElement searchBtn = driver.FindElement(By.XPath(searchBtnXpath));
            searchBtn.Click();

            IWebElement addrInputBox = driver.FindElement(By.XPath(inputBoxXpath));
            addrInputBox.Clear();
            addrInputBox.SendKeys(address);

            IWebElement addrSearchBtn = driver.FindElement(By.XPath(addrSearchBtnXpath));
            addrSearchBtn.Click();

            IWebElement okBtn = driver.FindElement(By.XPath(okBtnXpath));
            okBtn.Click();
        }

        //Get transaction Amount from Excel spreadsheet
        public string getTransactionAmmount()
        {

            //Directory of the data spreadsheet
            string FileDirectory = @"C:\WatercareInvoiceProcessing\ActiveSpreadsheet\WaterCareInvoice .xlsx";
            appObj = new Microsoft.Office.Interop.Excel.Application();
            objBook = appObj.Workbooks.Open(FileDirectory, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            objWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)objBook.Worksheets.get_Item(1);

            range = (Microsoft.Office.Interop.Excel.Range)objWorkSheet.Cells[1, 12];
            double amt = range.Value;
            return amt.ToString();
        }


        public void closeBrowser()
        {
            driver.Close();
        }

        //*********************************************************************************

    }
}
