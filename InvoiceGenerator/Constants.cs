using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InvoiceGenerator
{
    class Constants
    {
        public static string OutputFilePath = string.Empty;
        public static readonly string CustomerName = "Customer Name";
        public static readonly string ModeOfPayment = "Mode of Payment";
        public static readonly string OrderDate = "Order Date";
        public static readonly string InvoiceNo = "Invoice No";
        public static readonly string RefNo = "Ref. No.";
        public static readonly string CustomerID = "Customer ID";
        public static readonly string OrderID = "Order Number";
        public static readonly string Gender = "Gender";
        public static readonly string Address = "Address";
        public static readonly string City = "City";
        public static readonly string State = "State";
        public static readonly string Pincode = "Pin Code";
        public static readonly string EmailID = "Email id";
        public static readonly string Phone = "Customer Mobile";
        public static readonly string QTY = "Unit";
        public static readonly string MRP = "MRP";
        public static readonly string OrderValue = "Order Value";
        public static readonly string Products = "Products";
        public static readonly string ExcluceInvoice = "Exclude";
        public static readonly string Skip = "Skip";
        public static readonly string CourierName = "Courier Name";
        public static readonly string TrackingNumber = "Tracking (AWB) No";
        public static string ContestName = "SPL Contest";
        public static string FolderNameFormat = "dd-MM-yyyy";
        public static readonly string AlternameNumber = "Alternate Number";
        public static readonly string Shipping = "Shipping";
        public static string CustCareNumber = "";
        public static string TinNumber = "";
        public static string CompanyName = "HRD ENTERPRISES";

        static Constants()
        {
            LoadSettings();
        }

        public static void LoadSettings()
        {
            
            OutputFilePath = Properties.Settings.Default.OutputPath;
            FolderNameFormat = Properties.Settings.Default.FolderNameFormat;
            ContestName = Properties.Settings.Default.ContestName;
            CustCareNumber = Properties.Settings.Default.CustCareNo;
            TinNumber = Properties.Settings.Default.TinNumber;
            CompanyName = Properties.Settings.Default.Company;
        }
        
    }
}
