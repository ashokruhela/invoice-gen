using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace InvoiceGenerator
{
    internal class ExcelDataProvider
    {
        string filePath = string.Empty;
        public event EventHandler<string> UpdateProgress;
        Dictionary<int, string> columns = new Dictionary<int, string>();
        DataTable dtInvoice = new DataTable();
        public bool OpenExcelFile { get; set; }

        private void RaiseUpdateProgress(string currentValue)
        {
            if (UpdateProgress != null)
            {
                UpdateProgress(this, currentValue);
            }
        }

        public int GetMaxRows(string excelFilePath)
        {
            int rows = 0;
            Excel.Application excelApp = null;
            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                Excel.Workbook currentWorkBook = excelApp.Workbooks.Open(excelFilePath);
                Excel.Worksheet currentSheet = currentWorkBook.Sheets[1];
                rows = currentSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                Marshal.ReleaseComObject(currentWorkBook);
                Marshal.ReleaseComObject(currentSheet);
            }
            finally
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
            return rows;
        }

        public Task<DataTable> GetExcelDataAsync(string excelFilePath)
        {
            return Task.Run<DataTable>(() =>
            {

                DataTable dtExcelData = new DataTable("Excel Data");
                columns.Clear();
                Excel.Application excelApp = null;
                try
                {
                    excelApp = new Excel.Application();
                    excelApp.Visible = false;
                    Excel.Workbook currentWorkBook = excelApp.Workbooks.Open(excelFilePath);
                    Excel.Worksheet currentSheet = currentWorkBook.Sheets[1];
                    int lastRow = currentSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    Excel.Range last = currentSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    int lastColumn = currentSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    System.Array columnValues = (System.Array)currentSheet.get_Range("A1", last).Cells.Value;
                    for (int i = 0; i < lastColumn; i++)
                    {
                        DataColumn column = new DataColumn();
                        object val = columnValues.GetValue(1, i + 1);
                        column.Caption = val != null ? columnValues.GetValue(1, i + 1).ToString().Trim() : string.Empty;
                        dtExcelData.Columns.Add(column);
                        //Fill columns information
                        columns.Add(i, column.Caption);
                    }
                    //add extra columns to keep track which invoice to exclude
                    DataColumn skip = new DataColumn();
                    skip.Caption = Constants.Skip;
                    skip.DefaultValue = "NO";
                    columns.Add(columns.Count, skip.Caption);
                    dtExcelData.Columns.Add(skip);

                    if (!columns.ContainsValue(Constants.ExcluceInvoice))
                    {
                        DataColumn exclude = new DataColumn();
                        exclude.Caption = Constants.ExcluceInvoice;
                        exclude.DefaultValue = "NO";
                        columns.Add(columns.Count, exclude.Caption);
                        dtExcelData.Columns.Add(exclude);
                    }

                    for (int index = 2; index <= lastRow; index++)
                    {

                        DataRow newRow = dtExcelData.NewRow();
                        for (int i = 1; i <= lastColumn; i++)
                        {
                            object cellValue = columnValues.GetValue(index, i);
                            newRow[i - 1] = cellValue == null ? null : cellValue.ToString();
                        }
                        dtExcelData.Rows.Add(newRow);
                        string customerName = GetColumnValue(newRow, Constants.CustomerName);
                        if (customerName.Length > 0)
                            RaiseUpdateProgress(string.Format("Loading data for - {0}", customerName));
                    }
                    //delete rows to be excluded
                    List<DataRow> rowsToDelete = new List<DataRow>();
                    foreach (DataRow row in dtExcelData.Rows)
                    {
                        if (GetColumnValue(row, Constants.ExcluceInvoice).ToUpper() == "YES")
                            rowsToDelete.Add(row);
                    }
                    //delete excluded rows
                    foreach (DataRow row in rowsToDelete)
                    {
                        dtExcelData.Rows.Remove(row);
                    }
                }
                finally
                {
                    if (excelApp != null)
                    {
                        excelApp.Quit();
                    }
                }
                return dtExcelData;

            });


        }

        private string GetColumnValue(DataRow row, string columnCaption)
        {
            object cellValue = string.Empty;

            if (columns.ContainsValue(columnCaption))
            {
                var columnInfo = columns.FirstOrDefault(x => x.Value.ToUpper() == columnCaption.ToUpper());
                cellValue = row[columnInfo.Key];
            }

            return cellValue == null ? string.Empty : cellValue.ToString();
        }

        private Dictionary<string, string> GetColumnValue(DataRow row, List<string> columnCaptions)
        {
            Dictionary<string, string> columnValues = new Dictionary<string, string>();

            object cellValue = string.Empty;
            foreach (string columnCaption in columnCaptions)
            {
                if (columns.ContainsValue(columnCaption))
                {
                    var columnInfo = columns.FirstOrDefault(x => x.Value.ToUpper() == columnCaption.ToUpper());
                    cellValue = row[columnInfo.Key];
                    columnValues.Add(columnCaption, cellValue == null ? string.Empty : cellValue.ToString());
                }
            }

            return columnValues;
        }

        private void SetColumnValue(DataRow row, string columnCaption, string cellValue)
        {
            if (columns.ContainsValue(columnCaption))
            {
                var columnInfo = columns.FirstOrDefault(x => x.Value.ToUpper() == columnCaption.ToUpper());
                row[columnInfo.Key] = cellValue;
            }
        }

        public Task<bool> GenerateInvoice(DataTable dtTable)
        {
            return Task.Run(() =>
            {
                bool invoiceGenerated = true;
                Excel.Application excelApp = null;
                List<string> invoices = new List<string>();
                try
                {
                    ApplyFilterCondition(dtTable);
                    string tempValue = string.Empty;
                    var headerColor = Excel.XlRgbColor.rgbDeepSkyBlue;

                    foreach (DataRow row in dtTable.Rows)
                    {
                        string custName = GetColumnValue(row, Constants.CustomerName);
                        if (GetColumnValue(row, Constants.Skip).ToUpper() == "YES" || custName.Length == 0)
                            continue;


                        string outputFilename = GetFileName(custName);
                        excelApp = new Excel.Application();
                        excelApp.Visible = false;
                        object misValue = System.Reflection.Missing.Value;
                        string supportEmail = @"info@mineemart.com";
                        string companyTinNumber = string.Empty;
                        string website = "www.mineemart.com";

                        Excel.Workbook newWorkBook = excelApp.Workbooks.Add(misValue);
                        Excel.Worksheet newWorkSheet = (Excel.Worksheet)newWorkBook.Sheets.get_Item(1);
                        //fixed information
                        Excel.Range range = newWorkSheet.get_Range("A1", "K40");
                        range.Interior.Color = Excel.XlRgbColor.rgbWhite;
                        Marshal.FinalReleaseComObject(range);

                        #region Set Width and height
                        newWorkSheet.Columns[1].ColumnWidth = 43.29;
                        newWorkSheet.Columns[2].ColumnWidth = .50;
                        newWorkSheet.Columns[3].ColumnWidth = 1.5;
                        newWorkSheet.Columns[4].ColumnWidth = 2;
                        newWorkSheet.Columns[5].ColumnWidth = 6;
                        newWorkSheet.Columns[6].ColumnWidth = 13;
                        newWorkSheet.Columns[7].ColumnWidth = 5;
                        newWorkSheet.Columns[8].ColumnWidth = 7;
                        newWorkSheet.Columns[9].ColumnWidth = 19.14;
                        newWorkSheet.Columns[10].ColumnWidth = 15;

                        newWorkSheet.Rows[1].RowHeight = 26;
                        newWorkSheet.Rows[5].RowHeight = 25;
                        newWorkSheet.Rows[6].RowHeight = 25;
                        newWorkSheet.Rows[7].RowHeight = 25;
                        newWorkSheet.Rows[8].RowHeight = 25;
                        newWorkSheet.Rows[9].RowHeight = 25;
                        newWorkSheet.Rows[10].RowHeight = 25;
                        newWorkSheet.Rows[11].RowHeight = 25;
                        #endregion

                        #region  Header
                        string rangeValue = "Retail Invoice";
                        range = SetRangeParams(ref newWorkSheet, "A1", "J1", rangeValue: rangeValue, merge: true, bold: true, center: true, borderAround: true, release: false);
                        range.Cells.Font.Size = 20;
                        //range.Characters[1, rangeValue.Length].Font.FontStyle = "bold";
                        range.Interior.Color = headerColor;
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        range = newWorkSheet.get_Range("A4", "J4");
                        range.Interior.Color = headerColor;
                        Marshal.FinalReleaseComObject(range);

                        range = newWorkSheet.get_Range("A14", "J15");
                        range.Interior.Color = headerColor;
                        Marshal.FinalReleaseComObject(range);

                        //Border around all data
                        SetRangeParams(ref newWorkSheet, "A1", "J24", false, string.Empty, true, false, true);

                        //customer id
                        rangeValue = "CUSTOMER ID : " + GetColumnValue(row, Constants.CustomerID);
                        range = SetRangeParams(ref newWorkSheet, "A2", "A3", true, rangeValue, true, release: false, center: true);
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Characters[1, 13].Font.FontStyle = "bold";
                        range.Characters[1, 13].Font.Size = 14;
                        Marshal.FinalReleaseComObject(range);

                        //order id 
                        rangeValue = "ORDER ID : " + GetColumnValue(row, Constants.OrderID);
                        range = SetRangeParams(ref newWorkSheet, "B2", "H3", true, rangeValue, true, release: false, center: true);
                        range.Characters[1, 10].Font.FontStyle = "bold";
                        range.Characters[1, 10].Font.Size = 14;
                        Marshal.FinalReleaseComObject(range);

                        //invoice number
                        rangeValue = "INVOICE NO : " + GetColumnValue(row, Constants.InvoiceNo);
                        range = SetRangeParams(ref newWorkSheet, "I2", "J3", true, rangeValue, true, release: false, center: true);
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Characters[1, 13].Font.FontStyle = "bold";
                        range.Characters[1, 13].Font.Size = 14;
                        Marshal.FinalReleaseComObject(range);

                        //seller detils
                        string tinSupport = string.Format("Company's TIN/VAT No. :- {0}\n{1}", Constants.TinNumber, supportEmail);
                        string sellerAddress = string.Format("{0}\n\nNew Delhi, India\nTelephone: {1}\n\n{2}", 
                            Constants.CompanyName, Constants.Phone, tinSupport);
                        SetRangeParams(ref newWorkSheet, "A4", "D4", true, "SELLER DETAILS", true, false, true, bold: true);
                        range = SetRangeParams(ref newWorkSheet, "A5", "D11", true, sellerAddress, true, release: false);
                        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        range.Characters[1, Constants.CompanyName.Length].Font.FontStyle = "bold";
                        range.Characters[1, Constants.CompanyName.Length].Font.Size = 14;
                        int index = sellerAddress.IndexOf(tinSupport);
                        if(index > 0)
                        {
                            range.Characters[index + 1, tinSupport.Length].Font.Color = Excel.XlRgbColor.rgbDeepSkyBlue;
                            range.Characters[index + 1, tinSupport.Length].Font.Underline = true;
                        }

                        Marshal.FinalReleaseComObject(range);

                        //buyer detils
                        SetRangeParams(ref newWorkSheet, "E4", "J4", true, "BUYER", true, false, true, bold: true);
                        SetRangeParams(ref newWorkSheet, "E5", "J11", true, GetBillingAddress(row), true, false, true);

                        //invoice date
                        tempValue = GetColumnValue(row, Constants.OrderDate);
                        DateTime orderDate = DateTime.Now;
                        DateTime.TryParse(tempValue, out orderDate);
                        string title = "INVOICE DATE : ";
                        rangeValue = title + orderDate.ToString("dd/MM/yyyy");
                        range = SetRangeParams(ref newWorkSheet, "A12", "A13", true, rangeValue, release: false);
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        range.Characters[1, title.Length].Font.FontStyle = "bold";
                        Marshal.FinalReleaseComObject(range);

                        //courier name
                        title = "COURIER NAME: ";
                        rangeValue = title + GetColumnValue(row, Constants.CourierName);
                        range = SetRangeParams(ref newWorkSheet, "B12", "G13", true, rangeValue, release: false, center: true);
                        range.Characters[1, title.Length].Font.FontStyle = "bold";
                        Marshal.FinalReleaseComObject(range);

                        //Tracking (AWB) No.
                        title = "TRAKING (AWB) NO. : ";
                        rangeValue = title + GetColumnValue(row, Constants.TrackingNumber);
                        range = SetRangeParams(ref newWorkSheet, "H12", "J13", true, rangeValue, release: false);
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        range.Characters[1, title.Length].Font.FontStyle = "bold";
                        Marshal.FinalReleaseComObject(range);

                        //item details
                        SetRangeParams(ref newWorkSheet, "A14", "F15", true, "ITEM DETAILS", true, false, true, bold: true, center: true);
                        SetRangeParams(ref newWorkSheet, "A16", "F22", true, GetProducts(row), true, false, true);

                        //QTY
                        SetRangeParams(ref newWorkSheet, "G14", "G15", true, "QTY", true, false, true, bold: true, center: true);
                        string qty = GetColumnValue(row, Constants.QTY);
                        rangeValue = string.IsNullOrEmpty(qty) ? "1" : qty;
                        SetRangeParams(ref newWorkSheet, "G16", "G22", true, rangeValue, true, release: true, center: true);

                        //Gross Value
                        SetRangeParams(ref newWorkSheet, "H14", "I15", true, "Gross Value", true, false, true, bold: true, center: true);
                        string tempprice = GetColumnValue(row, Constants.OrderValue);
                        bool hasPaisa = !tempprice.Contains(".");
                        string price = !hasPaisa ? tempprice : string.Format("{0}.00", tempprice);
                        range = SetRangeParams(ref newWorkSheet, "H16", "I22", true, price, true, release: false, center: true);
                        range.NumberFormat = "0.00";
                        Marshal.FinalReleaseComObject(range);

                        //Net Amount
                        SetRangeParams(ref newWorkSheet, "J14", "J15", true, "Net Amount", true, false, true, bold: true, center: true);
                        range = SetRangeParams(ref newWorkSheet, "J16", "J22", true, price, true, release: false, center: true);
                        range.NumberFormat = "0.00";
                        Marshal.FinalReleaseComObject(range);

                        //empty block
                        SetRangeParams(ref newWorkSheet, "A23", "F23", true, string.Empty, true, false, true);

                        //Amount Payable:
                        SetRangeParams(ref newWorkSheet, "G23", "I23", true, "Amount Payable: ", true, false, true, bold: true);

                        //total amount
                        range = SetRangeParams(ref newWorkSheet, "J23", "J23", true, price, true, release: false, center: true);
                        range.NumberFormat = "0.00";
                        Marshal.FinalReleaseComObject(range);

                        //Amount In Words :
                        SetRangeParams(ref newWorkSheet, "A24", "A24", true, "Amount In Words : ", true, false, true, bold: true);
                        string words = HelpUtil.ToWords(Convert.ToDecimal(price));
                        SetRangeParams(ref newWorkSheet, "B24", "J24", true, words.ToUpper(), true, false, true, bold: true, center: true);

                        //Term and Conditions : 
                        SetRangeParams(ref newWorkSheet, "A26", "J26", true, "Term and Conditions : ", false, false, true, bold: true);
                        rangeValue = @"(1) Refunds will be made as per our refund policy.";
                        SetRangeParams(ref newWorkSheet, "A27", "J27", true, rangeValue, false, false, true);
                        rangeValue = @"(2) VAT/CST is applicable on above amount is: - Rs 000.00/- ";
                        SetRangeParams(ref newWorkSheet, "A28", "J28", true, rangeValue, false, false, true);
                        rangeValue = string.Format("(3) In Case of any queries, please call our customer care on: {0} or email: {1}", Constants.CustCareNumber, supportEmail);
                        SetRangeParams(ref newWorkSheet, "A29", "J29", true, rangeValue, false, false, true);
                        rangeValue = @"(4) All disputes are subject to the exclusive jurisdiction of competent courts and forums in Delhi/New Delhi only.";
                        SetRangeParams(ref newWorkSheet, "A30", "J30", true, rangeValue, false, false, true);

                        rangeValue = @"Visit us At : " + website;
                        SetRangeParams(ref newWorkSheet, "A33", "J33", true, rangeValue, false, false, true, bold: true, center: true);
                        rangeValue = @"This is a computer generated invoice. No signature required.";
                        SetRangeParams(ref newWorkSheet, "A34", "J34", true, rangeValue, false, false, true, bold: true, center: true);

                        newWorkSheet.PageSetup.Zoom = false;
                        newWorkSheet.PageSetup.FitToPagesWide = 1;
                        newWorkSheet.PageSetup.FitToPagesTall = 1;
                        newWorkBook.SaveAs(outputFilename, Excel.XlFileFormat.xlOpenXMLWorkbook,
                            System.Reflection.Missing.Value, Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                            Excel.XlSaveConflictResolution.xlUserResolution, true, Missing.Value, Missing.Value,
                            Missing.Value);
                        newWorkBook.Close(true, misValue, misValue);
                        Marshal.ReleaseComObject(newWorkSheet);
                        Marshal.ReleaseComObject(newWorkBook);
                        if (OpenExcelFile)
                        {
                            string[] temp = outputFilename.Split(new char[] { '\\' });
                            string file = temp[temp.Count() - 1];
                            string currentDate = DateTime.Today.ToString(Constants.FolderNameFormat);
                            var files = Directory.GetFiles(Path.Combine(Constants.OutputFilePath, DateTime.Today.ToString("MMMM"), currentDate), string.Format("{0}.*", file));
                            if (files.Count() > 0)
                            {
                                invoices.Add(files[0]);
                            }
                        }
                        RaiseUpdateProgress(string.Format("Generated invoice for  {0}", custName));
                    }
                }
                finally
                {
                    if(excelApp == null)
                    {
                        System.Windows.Forms.MessageBox.Show("No data to process. Either customer name could be empty, column name is invalid or skip is YES in sheet. Make sure that valid data is present.", 
                            "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    }
                    else
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                        try
                        {
                            foreach (string invoice in invoices)
                            {
                                System.Diagnostics.Process.Start(invoice);
                            }

                        }
                        catch { }
                    }
                    
                }
                return invoiceGenerated;
            });
        }

        public void ApplyFilterCondition(DataTable dtRawData)
        {
            //condition 1 - if cust id, order number, invoice no & ref no are same then 
            //there would be only one invoice and amount will be added
            //FilterCondition(dtRawData, 1);

            //condition 2 - if Only customer id & invoice no are same then the amount will be added but both invoice no has to be displayed
            FilterCondition(dtRawData, 2);
        }

        private void FilterCondition(DataTable dtRawData, int condition)
        {

            List<string> captionNames = new List<string>(new string[] { Constants.CustomerID, Constants.OrderID, Constants.InvoiceNo, Constants.Products });
            foreach (DataRow row in dtRawData.Rows)
            {
                Dictionary<string, string> mainRowValues = GetColumnValue(row, captionNames);
                string custID = mainRowValues[Constants.CustomerID];
                if (GetColumnValue(row, Constants.Skip).ToUpper() == "YES" || custID.Length == 0)
                    continue;

                string orderID = mainRowValues[Constants.OrderID];
                string invoiceNo = mainRowValues[Constants.InvoiceNo];
                //string refNo = mainRowValues[Constants.RefNo];
                string products = string.Empty;
                string newProducts = products = mainRowValues[Constants.Products];
                double totalPrice = Convert.ToDouble(GetColumnValue(row, Constants.OrderValue));
                string newOrderID = orderID;
                foreach (DataRow otherRow in dtRawData.Rows)
                {
                    if (GetColumnValue(otherRow, Constants.Skip).ToUpper() == "YES")
                        continue;
                    if (!row.Equals(otherRow))
                    {
                        Dictionary<string, string> otherRowValues = GetColumnValue(otherRow, captionNames);
                        string othercustID = otherRowValues[Constants.CustomerID];
                        string otherorderID = otherRowValues[Constants.OrderID];
                        string otherinvoiceNo = otherRowValues[Constants.InvoiceNo];
                        //string otherrefNo = otherRowValues[Constants.RefNo];
                        string otherProducts = otherRowValues[Constants.Products];

                        //if (condition == 1 && custID == othercustID && orderID == otherorderID && invoiceNo == otherinvoiceNo && refNo == otherrefNo)
                        //{
                        //    newProducts = string.Join("+", newProducts, otherProducts);
                        //    totalPrice += Convert.ToDouble(GetColumnValue(otherRow, Constants.OrderValue));
                        //    SetColumnValue(row, Constants.Products, newProducts);
                        //    SetColumnValue(otherRow, Constants.Skip, "Yes");
                        //    SetColumnValue(row, Constants.OrderValue, totalPrice.ToString());
                        //}
                        //else 
                        if (condition == 2 && custID == othercustID && invoiceNo == otherinvoiceNo)
                        {
                            newProducts = string.Join("+", newProducts, otherProducts);
                            totalPrice += Convert.ToDouble(GetColumnValue(otherRow, Constants.OrderValue));
                            // string lastDigits = otherorderID.Length >= 5 ? otherorderID.Substring(otherorderID.Length - 5, 5) : otherorderID;
                            newOrderID = string.Join("/", newOrderID, otherorderID);
                            SetColumnValue(otherRow, Constants.Skip, "Yes");
                            SetColumnValue(row, Constants.Products, newProducts);
                            SetColumnValue(row, Constants.OrderValue, totalPrice.ToString());
                            SetColumnValue(row, Constants.OrderID, newOrderID);
                        }
                    }
                }
            }
        }

        private string GetShippingAddress(DataRow row)
        {
            string gender = GetColumnValue(row, Constants.Gender);
            if (gender.ToUpper() == "F")
                gender = "Mrs.";
            else if (gender.ToUpper() == "U")
                gender = "Miss";
            else
                gender = "Mr.";
            string customerName = GetColumnValue(row, Constants.CustomerName);
            string address = GetColumnValue(row, Constants.Address);
            string city = GetColumnValue(row, Constants.City);
            string state = GetColumnValue(row, Constants.State);
            string pincode = GetColumnValue(row, Constants.Pincode);
            string emailID = GetColumnValue(row, Constants.EmailID);
            string phone = GetColumnValue(row, Constants.Phone);
            string alternateNo = GetColumnValue(row, Constants.AlternameNumber);
            if (alternateNo.Length > 0 && alternateNo.Trim() != "-")
            {
                phone = string.Join("/", phone, alternateNo);
            }
            string billingAddress = string.Format("{0} {1}\n\n{2}\n\n{3}\n{4} Pin {5}\nIndia\n{6}\n+91{7}", gender, customerName, address, city, state, pincode, emailID, phone);
            return billingAddress;
        }

        private string GetBillingAddress(DataRow row)
        {
            return GetShippingAddress(row);
        }

        private string GetInvoice(DataRow row)
        {
            string invoiceNo = string.Empty;
            invoiceNo = GetColumnValue(row, Constants.InvoiceNo);
            if (invoiceNo.Length == 0)
                invoiceNo = GetColumnValue(row, Constants.RefNo);
            return invoiceNo;
        }

        private string GetFileName(string customerName)
        {
            customerName = customerName.Replace(".", string.Empty);
            string currentDate = DateTime.Today.ToString(Constants.FolderNameFormat);
            string month = DateTime.Today.ToString("MMMM");
            string fileName = Path.Combine(Constants.OutputFilePath, month, currentDate, customerName);
            if (!Directory.Exists(Path.Combine(Constants.OutputFilePath, month, currentDate)))
            {
                Directory.CreateDirectory(Path.Combine(Constants.OutputFilePath, month, currentDate));
            }
            var files = Directory.GetFiles(Path.Combine(Constants.OutputFilePath, month, currentDate), string.Format("{0}*", customerName));
            fileName = files.Count() == 0 ? fileName : string.Format("{0}_{1}", fileName, (files.Count() + 1).ToString());

            if (File.Exists(fileName))
            {
                Random randomNumber = new Random();
                fileName = string.Format("{0}_{1}", fileName, randomNumber.Next(50, 100));
            }

            return fileName;
        }

        private string GetProducts(DataRow row)
        {
            var temp = GetColumnValue(row, Constants.Products).Split(new char[] { '+' });
            StringBuilder products = new StringBuilder();
            int productCount = temp.Count();

            foreach (string product in temp)
            {
                products.AppendLine(product.Trim());
            }
            if (productCount > 4)
            {
                products = products.Replace("\n", ", ");
                //Remove last comma
                products.Remove(products.Length - 3, 2);
            }
            return products.ToString();
        }

        private Excel.Range SetRangeParams(ref Excel.Worksheet newWorkSheet, string startRange, string endRange, 
            bool merge = true, string rangeValue = "", bool borderAround = false, 
            bool lineStyle = false, bool release = true, bool bold = false, bool center = false)
        {
            Excel.Range range = newWorkSheet.get_Range(startRange, endRange);
            range.WrapText = true;
            if(merge)
                range.Merge();
            if(borderAround)
                range.BorderAround(Type.Missing, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
            if(lineStyle)
                range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            if(!string.IsNullOrEmpty(rangeValue))
                range.Value2 = rangeValue;
            range.Font.Bold = bold;

            if(center)
            {
                range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }

            if (release)
                Marshal.FinalReleaseComObject(range);
            return range;
        }
    }
}
