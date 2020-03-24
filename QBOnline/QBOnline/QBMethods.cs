using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using Intuit.Ipp.Core;
using Intuit.Ipp.Data;
using System.Configuration;
using Intuit.Ipp.OAuth2PlatformClient;
using Intuit.Ipp.Core.Configuration;
using Intuit.Ipp.Security;
using Intuit.Ipp.DataService;
using System.Diagnostics;
using System.Data.SqlClient;


namespace TSI.QBInterface
{
    public class QBMethods
    {


        // OAuth2 client configuration

        static string redirectURI = ConfigurationManager.AppSettings["redirectURI"];
        static string clientID = ConfigurationManager.AppSettings["clientID"];
        static string clientSecret = ConfigurationManager.AppSettings["clientSecret"];
        static string logPath = ConfigurationManager.AppSettings["logPath"];
        static string appEnvironment = ConfigurationManager.AppSettings["appEnvironment"];

        static string realmId = ConfigurationManager.AppSettings["realmId"];
        static string connectionString;
        static string authCode;
        static string idToken;
        public static IList<JsonWebKey> keys;
        public static Dictionary<string, string> dictionary = new Dictionary<string, string>();
        public static ServiceContext servicecontext;
        public static DataService services;
        public static OAuth2Client oauthClient = new OAuth2Client(clientID, clientSecret, redirectURI, appEnvironment);

        public static string AccessToken;
        public static string RefreshToken;
        private static Dictionary<string, string> CustomerIDs = new Dictionary<string, string>();
        private static Dictionary<string, Invoice> Invoices = new Dictionary<string, Invoice>();
        private static Dictionary<string, string> Items = new Dictionary<string, string>();
        private static Dictionary<string, string> Classes = new Dictionary<string, string>();
        private static Dictionary<string, string> Terms = new Dictionary<string, string>();
        private static Dictionary<string, string> Departments = new Dictionary<string, string>();
        private static Dictionary<string, string> CustomerTypes = new Dictionary<string, string>();
        private static Dictionary<string, string> Vendors = new Dictionary<string, string>();
        private static Dictionary<string, Account> Accounts = new Dictionary<string, Account>();
        //private string authorizeUrl;
        public QBMethods()
        {

            connectionString = ConfigurationManager.ConnectionStrings["JobScheduler2007"].ConnectionString;

        }

        public static void CreateService()
        {
            List<OidcScopes> scopes = new List<OidcScopes>();
            scopes.Add(OidcScopes.Accounting);
            OAuth2RequestValidator reqValidator = new OAuth2RequestValidator(AccessToken);
            servicecontext = new ServiceContext(realmId, IntuitServicesType.QBO, reqValidator);

            services = new DataService(servicecontext);
            servicecontext.IppConfiguration.AdvancedLogger.RequestAdvancedLog.EnableSerilogRequestResponseLoggingForConsole = true;
            servicecontext.IppConfiguration.AdvancedLogger.RequestAdvancedLog.EnableSerilogRequestResponseLoggingForDebug = true;
            servicecontext.IppConfiguration.AdvancedLogger.RequestAdvancedLog.EnableSerilogRequestResponseLoggingForRollingFile = true;
            servicecontext.IppConfiguration.AdvancedLogger.RequestAdvancedLog.EnableSerilogRequestResponseLoggingForTrace = true;
            servicecontext.IppConfiguration.AdvancedLogger.RequestAdvancedLog.ServiceRequestLoggingLocationForFile = logPath;//Any drive logging location
        }


        public void ConnecttoQBAuth()
        {

            try
            {
                if (!dictionary.ContainsKey("accessToken"))
                {
                    List<OidcScopes> scopes = new List<OidcScopes>();
                    scopes.Add(OidcScopes.Accounting);
                    var authorizeUrl = oauthClient.GetAuthorizationURL(scopes);
                    ProcessStartInfo sInfo = new ProcessStartInfo(authorizeUrl);
                    Process.Start(sInfo);
                }
            }
            catch (Exception ex)
            {

            }
        }



        public static void UpdateSecurityTokensInDB()
        {

            SqlConnection connection = new SqlConnection(connectionString);

            try
            {

                connection.Open();

                SqlCommand Updatecmd = new SqlCommand("Update QuickBooksSecurityTokens SET AccessToken = @AccessToken, RefreshToken = @RefreshToken, IsAuthenticate = 0 where IsAuthenticate = 1",
                    connection);

                int updCount = 0;

                Updatecmd.Parameters.Clear();
                Updatecmd.Parameters.AddWithValue("@AccessToken", AccessToken);
                Updatecmd.Parameters.AddWithValue("@RefreshToken", RefreshToken);
                updCount += Updatecmd.ExecuteNonQuery();

            }
            catch (Exception ex)
            {

            }
            finally
            {
                connection.Close();
            }


        }

        public static void UpdateSecurityToAuthenticate(int employeeID)
        {

            SqlConnection connection = new SqlConnection(connectionString);

            try
            {

                connection.Open();

                SqlCommand Updatecmd = new SqlCommand("Update QuickBooksSecurityTokens SET IsAuthenticate = 1 where employee_id = @employee_id",
                    connection);

                int updCount = 0;

                Updatecmd.Parameters.Clear();
                Updatecmd.Parameters.AddWithValue("@employee_id", employeeID);

                updCount += Updatecmd.ExecuteNonQuery();

            }
            catch (Exception ex)
            {

            }
            finally
            {
                connection.Close();
            }


        }

        public static void GetDBTokens(int employeeID)
        {
            SqlConnection connection = new SqlConnection(connectionString);

            try
            {

                connection.Open();

                SqlCommand Updatecmd = new SqlCommand("select AccessToken, RefreshToken from QuickBooksSecurityTokens where employee_id = @employee_id",
                    connection);



                Updatecmd.Parameters.Clear();
                Updatecmd.Parameters.AddWithValue("@employee_id", employeeID);

                SqlDataReader DR = Updatecmd.ExecuteReader();
                DataTable DT = new DataTable();
                DT.Load(DR);

                if (DT.Rows.Count > 0)
                {
                    AccessToken = (string)DT.Rows[0]["AccessToken"];
                    RefreshToken = (string)DT.Rows[0]["RefreshToken"];
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                connection.Close();
            }
        }


        public string CreatePayablesBill(DataTable dtBill)
        {

            if (services == null)
                CreateService();
            if (Items.Count == 0)
                GetProductAndServicesPrefs();
            if (Classes.Count == 0)
                GetClasses();
            if (Terms.Count == 0)
                GetTerms();
            if (Vendors.Count == 0)
                GetVendors();
            if (Accounts.Count == 0)
                GetAccounts();
            if (Departments.Count == 0)
                GetDepartments();

            Bill bill = new Bill();


            bill.APAccountRef = new ReferenceType()
            {
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Account),
                name = "Account Payable",
                Value = "379"
            };

      

            bill.DepartmentRef = new ReferenceType();


            bill.DepartmentRef.type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Department);
            bill.DepartmentRef.name = (string)dtBill.Rows[0]["RegionName"];
            bill.DepartmentRef.Value = Departments[(string)dtBill.Rows[0]["RegionName"]];
            

            bill.VendorRef = new ReferenceType();
            bill.VendorRef.type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Vendor);

            if (!Vendors.ContainsKey((string)dtBill.Rows[0]["VendorName"]))
                return "Vendor Bill not loaded: Vendor not in QB";

            bill.VendorRef.name = (string)dtBill.Rows[0]["VendorName"];
            bill.VendorRef.Value =Vendors[(string)dtBill.Rows[0]["VendorName"]];
            bill.DocNumber = ((string)dtBill.Rows[0]["VendorInvoiceNumber"]);
         

            bill.TxnDate = (DateTime)dtBill.Rows[0]["BillDate"];
            bill.TxnDateSpecified = true;
            bill.DueDate = (DateTime)dtBill.Rows[0]["BillDate"];
            bill.DueDateSpecified = true;
           

            Line[] lines = new Line[dtBill.Rows.Count];
            int billLineCnt = 0;
            foreach (System.Data.DataRow poItemrow in dtBill.Rows)
            {
                Line line1 = new Line();
                line1.Amount = (decimal)poItemrow["BillLineAmount"];
                line1.AmountSpecified = true;
                line1.Description = poItemrow["BillLineDescription"].ToString();
                line1.DetailType = LineDetailTypeEnum.AccountBasedExpenseLineDetail;
                line1.DetailTypeSpecified = true;
              
                AccountBasedExpenseLineDetail lineDetail = new AccountBasedExpenseLineDetail();
                lineDetail.BillableStatus = BillableStatusEnum.Billable;

               
                lineDetail.AccountRef = new ReferenceType();
                lineDetail.AccountRef.name = Accounts[(string)poItemrow["BillLineGLCode"]].Name;
                
                lineDetail.AccountRef.Value = Accounts[(string)poItemrow["BillLineGLCode"]].Id;
                lineDetail.AccountRef.type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Account);


                lineDetail.ClassRef = new ReferenceType
                {

                    type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Class),
                    name = (string)poItemrow["RegionName"],
                    Value = Classes[(string)poItemrow["RegionName"]]

                };

                line1.AnyIntuitObject = lineDetail;

                lines[billLineCnt] = line1;
                billLineCnt++;

                bill.Line = lines;
            }
            Bill resultBill = services.Add(bill) as Bill;
            return "Vendor Bill Loaded";
        }




        public string CreateReceivablesInvoice(DataTable dtInvoice)
        {

            if (services == null)
                CreateService();
            if (Items.Count == 0)
                GetProductAndServicesPrefs();
            if (Classes.Count == 0)
                GetClasses();
            if (Terms.Count == 0)
                GetTerms();
            if (Departments.Count == 0)
                GetDepartments();
            

            Invoice invoice = new Invoice();
            //invoice.TemplateRef.name = "Trinity Stairs Invoice";
            invoice.TemplateRef = new ReferenceType()
            {
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.CompanyInfo),
                //name = "Type",
                name = "Trinity Stairs Invoice"
            };
            invoice.DocNumber = Convert.ToString(dtInvoice.Rows[0]["InvoiceID"]);
            if (CustomerIDs.Count == 0)
                GetCustomerList();

            invoice.CustomerRef = new ReferenceType();
            invoice.CustomerRef.type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Customer);
            invoice.CustomerRef.name = (string)dtInvoice.Rows[0]["CustomerName"];
            if (!CustomerIDs.ContainsKey((string)dtInvoice.Rows[0]["CustomerName"]))
                return "Invoice not loaded: Customer not in QB";

            invoice.CustomerRef.Value = CustomerIDs[(string)dtInvoice.Rows[0]["CustomerName"]];



            invoice.ClassRef = new ReferenceType()
            {
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Class),
                name = (string)dtInvoice.Rows[0]["RegionName"],
                Value = Classes[(string)dtInvoice.Rows[0]["RegionName"]]
            };
            invoice.CustomField = new CustomField[2];
            if (dtInvoice.Rows[0]["SalesOrderNumber"] != DBNull.Value)
            {
                string tempPONumber = (string)dtInvoice.Rows[0]["SalesOrderNumber"];
                if (tempPONumber.Length > 25)
                    tempPONumber = tempPONumber.Remove(25);
                invoice.PONumber = tempPONumber;

                CustomField cf1 = new CustomField();
                cf1.Type = CustomFieldTypeEnum.StringType;
                cf1.Name = "P.O. Number";
                cf1.AnyIntuitObject = tempPONumber;
                cf1.DefinitionId = "1";
                invoice.CustomField[0] = cf1;
            }
            invoice.TxnDate = (DateTime)dtInvoice.Rows[0]["InvoiceDate"];
            invoice.TxnDateSpecified = true;
            invoice.DueDate = (DateTime)dtInvoice.Rows[0]["InvoiceDate"];
            invoice.DueDateSpecified = true;
            invoice.ShipDate = (DateTime)dtInvoice.Rows[0]["InvoiceDate"];
            invoice.ShipDateSpecified = true;

            invoice.DepartmentRef = new ReferenceType()

            {
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Department),
                name = (string)dtInvoice.Rows[0]["RegionName"],
                Value = Departments[(string)dtInvoice.Rows[0]["RegionName"]]
            };

            if (dtInvoice.Rows[0]["PaymentTermDescription"] != DBNull.Value)
            {
                invoice.SalesTermRef = new ReferenceType();
                invoice.SalesTermRef.name = ((string)dtInvoice.Rows[0]["PaymentTermDescription"]).TrimEnd();
                if (!Terms.ContainsKey(((string)dtInvoice.Rows[0]["PaymentTermDescription"]).TrimEnd()))
                {
                    return "Payment Terms not in Terms collection:" + ((string)dtInvoice.Rows[0]["PaymentTermDescription"]).TrimEnd();
                }
                invoice.SalesTermRef.Value = Terms[((string)dtInvoice.Rows[0]["PaymentTermDescription"]).TrimEnd()];
                invoice.SalesTermRef.type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Term);
            }


            //if (dtInvoice.Rows[0]["ScheduledDate"] != DBNull.Value)
            // invoice.Other.SetValue(((DateTime)dtInvoice.Rows[0]["ScheduledDate"]).ToShortDateString());
            invoice.ShipAddr = new PhysicalAddress()
            {
                Line1 = (string)dtInvoice.Rows[0]["AddressStreet"],
                CountrySubDivisionCode = (string)dtInvoice.Rows[0]["AddressState"],
                City = (string)dtInvoice.Rows[0]["AddressCity"],
                PostalCode = (string)dtInvoice.Rows[0]["AddressZip"]
            };

            if (dtInvoice.Rows[0]["InvoiceTax"] != DBNull.Value)
            {
                invoice.TxnTaxDetail = new TxnTaxDetail();
                invoice.TxnTaxDetail.TotalTax = (decimal)dtInvoice.Rows[0]["InvoiceTax"];
                invoice.TxnTaxDetail.TotalTaxSpecified = true;

            }



            List<Line> invLines = new List<Line>();
            int ivLineCnt = 0;
            //Add Invoice Lines
            foreach (DataRow dtRow in dtInvoice.Rows)
            {
                Line invoiceLine = new Line();
                //invoiceLine.ItemRef.FullName.SetValue((string)dtRow["Jobtype"]);
                invoiceLine.Description = ((string)dtRow["InvoiceLineText"]);
                invoiceLine.Amount = (Convert.ToDecimal(dtRow["InvoiceLineTotalAmount"]));
                invoiceLine.AmountSpecified = true;

                invoiceLine.DetailType = LineDetailTypeEnum.SalesItemLineDetail;
                invoiceLine.DetailTypeSpecified = true;
                invoiceLine.Id = (ivLineCnt + 1).ToString();
                ivLineCnt++;



                //invoiceLine.

                SalesItemLineDetail lineSalesItemLineDetail = new SalesItemLineDetail();
                lineSalesItemLineDetail.ItemRef = new ReferenceType()
                {
                    type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Item),
                    name = (string)dtRow["JobType"],
                    Value = Items[(string)dtRow["JobType"]]
                };
                lineSalesItemLineDetail.ClassRef = new ReferenceType()
                {
                    type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Class),
                    name = (string)dtRow["RegionName"],
                    Value = Classes[(string)dtRow["RegionName"]]
                };

                lineSalesItemLineDetail.Qty = 1m;
                lineSalesItemLineDetail.QtySpecified = true;

                //Causes XML Error
                //lineSalesItemLineDetail.AnyIntuitObject = Convert.ToInt64(Math.Abs(Convert.ToDecimal(dtRow["InvoiceLineTotalAmount"])));

                //Line Sales Item Line Detail - TaxCodeRef
                //For US companies, this can be 'TAX' or 'NON'
                if (dtInvoice.Rows[0]["InvoiceTax"] != DBNull.Value)
                {
                    lineSalesItemLineDetail.TaxCodeRef = new ReferenceType();
                    if ((decimal)dtRow["InvoiceTax"] > 0)
                    {
                        lineSalesItemLineDetail.TaxCodeRef.Value = "TAX";
                    }
                    else
                    {
                        lineSalesItemLineDetail.TaxCodeRef.Value = "NON";
                    }

                }

                //Line Sales Item Line Detail - ServiceDate 
                lineSalesItemLineDetail.ServiceDate = DateTime.Now.Date;
                lineSalesItemLineDetail.ServiceDateSpecified = true;
                //Assign Sales Item Line Detail to Line Item
                invoiceLine.AnyIntuitObject = lineSalesItemLineDetail;

                invLines.Add(invoiceLine);
            }

            if (dtInvoice.Rows[0]["InvoiceDiscount"] != DBNull.Value && Convert.ToDecimal(dtInvoice.Rows[0]["InvoiceDiscount"]) != 0)
            {

                Line invoiceLine = new Line();
                invoiceLine.Description = "Discount";
                invoiceLine.Amount = (Convert.ToDecimal(dtInvoice.Rows[0]["InvoiceDiscount"]));
                invoiceLine.AmountSpecified = true;
                invoiceLine.DetailType = LineDetailTypeEnum.DiscountLineDetail;
                invoiceLine.Id = (ivLineCnt + 1).ToString();
                ivLineCnt++;
                invLines.Add(invoiceLine);

            }

            if (dtInvoice.Rows[0]["InvoiceShipping"] != DBNull.Value && Convert.ToDecimal(dtInvoice.Rows[0]["InvoiceShipping"]) > 0)
            {
                Line invoiceLine = new Line()
                {
                    Amount = (decimal)dtInvoice.Rows[0]["InvoiceShipping"],
                    AmountSpecified = true,
                    DetailType = LineDetailTypeEnum.SalesItemLineDetail,
                    Description = "Shipping"
                };
                invoiceLine.DetailTypeSpecified = true;
                invoiceLine.Id = (ivLineCnt + 1).ToString();
                ivLineCnt++;
                invLines.Add(invoiceLine);
            }



            //invoice.CustomerMemo = new MemoRef()
            //{
            //    Value = sentInvAddr
            //};

            invoice.ARAccountRef = new ReferenceType()
            {
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Account),
                name = "Account Receivable",
                Value = string.Concat((string)dtInvoice.Rows[0]["RegionARGLCode"])
            };

            invoice.TotalAmt = (Convert.ToDecimal(dtInvoice.Rows[0]["JobInvoiceAmount"]));
            invoice.TotalAmtSpecified = true;

            invoice.Line = invLines.ToArray();

            Invoice resultBill = services.Add(invoice) as Invoice;

            return "Invoice Loaded";

        }

        public string CreateReceivablesCreditMemo(DataTable dtInvoice)
        {
            if (services == null)
                CreateService();
            if (Items.Count == 0)
                GetProductAndServicesPrefs();
            if (Classes.Count == 0)
                GetClasses();
            if (Terms.Count == 0)
                GetTerms();
            if (Departments.Count == 0)
                GetDepartments();

            CreditMemo creditMemo = new CreditMemo();
            //invoice.TemplateRef.name = "Trinity Stairs Invoice";
            creditMemo.TemplateRef = new ReferenceType()
            {
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.CompanyInfo),
                //name = "Type",
                name = "Trinity Stairs Credit Memo"
            };
            creditMemo.DocNumber = Convert.ToString(dtInvoice.Rows[0]["InvoiceID"]);
            if (CustomerIDs.Count == 0)
                GetCustomerList();

            creditMemo.CustomerRef = new ReferenceType();
            creditMemo.CustomerRef.type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Customer);
            creditMemo.CustomerRef.name = (string)dtInvoice.Rows[0]["CustomerName"];
            if (!CustomerIDs.ContainsKey((string)dtInvoice.Rows[0]["CustomerName"]))
                return "Credit Memo not loaded: Customer not in QB";

            creditMemo.CustomerRef.Value = CustomerIDs[(string)dtInvoice.Rows[0]["CustomerName"]];

            creditMemo.ClassRef = new ReferenceType()
            {
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Class),
                name = (string)dtInvoice.Rows[0]["RegionName"],
                Value = Classes[(string)dtInvoice.Rows[0]["RegionName"]]
            };
            creditMemo.CustomField = new CustomField[2];
            if (dtInvoice.Rows[0]["SalesOrderNumber"] != DBNull.Value)
            {
                string tempPONumber = (string)dtInvoice.Rows[0]["SalesOrderNumber"];
                if (tempPONumber.Length > 25)
                    tempPONumber = tempPONumber.Remove(25);
                creditMemo.PONumber = tempPONumber;

                CustomField cf1 = new CustomField();
                cf1.Type = CustomFieldTypeEnum.StringType;
                cf1.Name = "P.O. Number";
                cf1.AnyIntuitObject = tempPONumber;
                cf1.DefinitionId = "1";
                creditMemo.CustomField[0] = cf1;
            }

            creditMemo.TxnDate = (DateTime)dtInvoice.Rows[0]["InvoiceDate"];
            creditMemo.TxnDateSpecified = true;
            creditMemo.DueDate = (DateTime)dtInvoice.Rows[0]["InvoiceDate"];
            creditMemo.DueDateSpecified = true;
            creditMemo.ShipDate = (DateTime)dtInvoice.Rows[0]["InvoiceDate"];
            creditMemo.ShipDateSpecified = true;

            creditMemo.DepartmentRef = new ReferenceType()

            {
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Department),
                name = (string)dtInvoice.Rows[0]["RegionName"],
                Value = Departments[(string)dtInvoice.Rows[0]["RegionName"]]
            };

            if (dtInvoice.Rows[0]["PaymentTermDescription"] != DBNull.Value)
            {
                creditMemo.SalesTermRef = new ReferenceType();
                creditMemo.SalesTermRef.name = ((string)dtInvoice.Rows[0]["PaymentTermDescription"]).TrimEnd();
                if (!Terms.ContainsKey(((string)dtInvoice.Rows[0]["PaymentTermDescription"]).TrimEnd()))
                {
                    return "Payment Terms not in Terms collection:" + ((string)dtInvoice.Rows[0]["PaymentTermDescription"]).TrimEnd();
                }
                creditMemo.SalesTermRef.Value = Terms[((string)dtInvoice.Rows[0]["PaymentTermDescription"]).TrimEnd()];
                creditMemo.SalesTermRef.type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Term);
            }


            //if (dtInvoice.Rows[0]["ScheduledDate"] != DBNull.Value)
            // invoice.Other.SetValue(((DateTime)dtInvoice.Rows[0]["ScheduledDate"]).ToShortDateString());
            creditMemo.ShipAddr = new PhysicalAddress()
            {
                Line1 = (string)dtInvoice.Rows[0]["AddressStreet"],
                CountrySubDivisionCode = (string)dtInvoice.Rows[0]["AddressState"],
                City = (string)dtInvoice.Rows[0]["AddressCity"],
                PostalCode = (string)dtInvoice.Rows[0]["AddressZip"]
            };

            if (dtInvoice.Rows[0]["InvoiceTax"] != DBNull.Value)
            {
                creditMemo.TxnTaxDetail = new TxnTaxDetail();
                creditMemo.TxnTaxDetail.TotalTax = (decimal)dtInvoice.Rows[0]["InvoiceTax"];
                creditMemo.TxnTaxDetail.TotalTaxSpecified = true;

            }



            List<Line> CMLines = new List<Line>();
            int cmLineCnt = 0;
            //Add Invoice Lines
            foreach (DataRow dtRow in dtInvoice.Rows)
            {
                Line creditMemoLine = new Line();
                //creditMemoLine.ItemRef.FullName.SetValue((string)dtRow["Jobtype"]);
                creditMemoLine.Description = ((string)dtRow["InvoiceLineText"]);
                creditMemoLine.Amount = (Convert.ToDecimal(dtRow["InvoiceLineTotalAmount"]));
                creditMemoLine.AmountSpecified = true;

                creditMemoLine.DetailType = LineDetailTypeEnum.SalesItemLineDetail;
                creditMemoLine.DetailTypeSpecified = true;
                creditMemoLine.Id = (cmLineCnt + 1).ToString();
                cmLineCnt++;



                SalesItemLineDetail lineSalesItemLineDetail = new SalesItemLineDetail();
                lineSalesItemLineDetail.ItemRef = new ReferenceType()
                {
                    type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Item),
                    name = (string)dtRow["JobType"],
                    Value = Items[(string)dtRow["JobType"]]
                };
                lineSalesItemLineDetail.ClassRef = new ReferenceType()
                {
                    type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Class),
                    name = (string)dtRow["RegionName"],
                    Value = Classes[(string)dtRow["RegionName"]]
                };

                lineSalesItemLineDetail.Qty = 1m;
                lineSalesItemLineDetail.QtySpecified = true;

                //Causes XML Error
                //lineSalesItemLineDetail.AnyIntuitObject = Convert.ToInt64(Math.Abs(Convert.ToDecimal(dtRow["InvoiceLineTotalAmount"])));

                //Line Sales Item Line Detail - TaxCodeRef
                //For US companies, this can be 'TAX' or 'NON'
                if (dtInvoice.Rows[0]["InvoiceTax"] != DBNull.Value)
                {
                    lineSalesItemLineDetail.TaxCodeRef = new ReferenceType();
                    if ((decimal)dtRow["InvoiceTax"] > 0)
                    {
                        lineSalesItemLineDetail.TaxCodeRef.Value = "TAX";
                    }
                    else
                    {
                        lineSalesItemLineDetail.TaxCodeRef.Value = "NON";
                    }

                }

                //Line Sales Item Line Detail - ServiceDate 
                lineSalesItemLineDetail.ServiceDate = DateTime.Now.Date;
                lineSalesItemLineDetail.ServiceDateSpecified = true;
                //Assign Sales Item Line Detail to Line Item
                creditMemoLine.AnyIntuitObject = lineSalesItemLineDetail;

                CMLines.Add(creditMemoLine);
            }

            if (dtInvoice.Rows[0]["InvoiceDiscount"] != DBNull.Value && Convert.ToDecimal(dtInvoice.Rows[0]["InvoiceDiscount"]) != 0)
            {

                Line creditMemoLine = new Line();
                creditMemoLine.Description = "Discount";
                creditMemoLine.Amount = (Convert.ToDecimal(dtInvoice.Rows[0]["InvoiceDiscount"]));
                creditMemoLine.AmountSpecified = true;
                creditMemoLine.DetailType = LineDetailTypeEnum.DiscountLineDetail;
                creditMemoLine.Id = (cmLineCnt + 1).ToString();
                cmLineCnt++;
                CMLines.Add(creditMemoLine);

            }

            if (dtInvoice.Rows[0]["InvoiceShipping"] != DBNull.Value && Convert.ToDecimal(dtInvoice.Rows[0]["InvoiceShipping"]) > 0)
            {
                Line creditMemoLine = new Line()
                {
                    Amount = (decimal)dtInvoice.Rows[0]["InvoiceShipping"],
                    AmountSpecified = true,
                    DetailType = LineDetailTypeEnum.SalesItemLineDetail,
                    Description = "Shipping"
                };
                creditMemoLine.DetailTypeSpecified = true;
                creditMemoLine.Id = (cmLineCnt + 1).ToString();
                cmLineCnt++;
                CMLines.Add(creditMemoLine);
            }



            //invoice.CustomerMemo = new MemoRef()
            //{
            //    Value = sentInvAddr
            //};

            creditMemo.ARAccountRef = new ReferenceType()
            {
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Account),
                name = "Account Receivable",
                Value = string.Concat((string)dtInvoice.Rows[0]["RegionARGLCode"])
            };

            creditMemo.TotalAmt = (Convert.ToDecimal(dtInvoice.Rows[0]["JobInvoiceAmount"]));
            creditMemo.TotalAmtSpecified = true;

            creditMemo.Line = CMLines.ToArray();

            CreditMemo resultBill = services.Add(creditMemo) as CreditMemo;

            return "Credit Memo Loaded";
            //CreditMemo creditMemo = new CreditMemo();
            //creditMemo.DocNumber = Guid.NewGuid().ToString("N").Substring(0, 10);

            //creditMemo.CustomerRef = new ReferenceType()
            //{
            //    type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Customer),
            //    //name = "TestCustomer",
            //    Value = string.Concat("QB:", (string)dtInvoice.Rows[0]["CustomerName"])
            //};

            //if (dtInvoice.Rows[0]["SalesOrderNumber"] != DBNull.Value)
            //    creditMemo.PONumber = (string)dtInvoice.Rows[0]["SalesOrderNumber"];
            ////creditMemo.CustomerMemo = new MemoRef()
            ////{
            ////    Value = sentInvAddr
            ////};

            ////creditMemo.TotalAmt = sentInvAmount;
            //creditMemo.TotalAmtSpecified = false;

            //if (dtInvoice.Rows[0]["InvoiceTax"] != null && Convert.ToDecimal(dtInvoice.Rows[0]["InvoiceTax"]) != 0)
            //{
            //    creditMemo.TxnTaxDetail = new TxnTaxDetail();
            //    creditMemo.TxnTaxDetail.TotalTax = (decimal)dtInvoice.Rows[0]["InvoiceTax"];
            //    creditMemo.TxnTaxDetail.TotalTaxSpecified = true;
            //}

            //creditMemo.DueDate = (DateTime)dtInvoice.Rows[0]["InvoiceDate"];
            //creditMemo.DueDateSpecified = true;

            //creditMemo.ARAccountRef = new ReferenceType()
            //{
            //    type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Account),
            //    name = "Account Receivable",
            //    Value = "QB:37"
            //};

            //creditMemo.ClassRef = new ReferenceType()
            //{
            //    type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Class),
            //    name = "Class ??",
            //    Value = "QB:37"
            //};
            //creditMemo.TxnDate = DateTime.Today.Date;
            //creditMemo.TxnDateSpecified = true;

            //Line[] lines = new Line[dtInvoice.Rows.Count];
            //int ivLineCnt = 0;
            //foreach (System.Data.DataRow dtRow in dtInvoice.Rows)
            //{
            //    Line line = new Line();
            //    line.Amount = decimal.Round((decimal)dtRow["InvoiceAmount"], 2);
            //    line.AmountSpecified = true;
            //    line.Description = (string)dtRow["DistType"];

            //    if ((string)dtRow["DistType"] == "Sales")
            //        line.DetailType = LineDetailTypeEnum.SalesItemLineDetail;
            //    else if ((string)dtRow["DistType"] == "Shipping")
            //        line.DetailType = LineDetailTypeEnum.SalesItemLineDetail;
            //    else if ((string)dtRow["DistType"] == "Discount")
            //        line.DetailType = LineDetailTypeEnum.DiscountLineDetail;
            //    else if ((string)dtRow["DistType"] == "Deposit")
            //        line.DetailType = LineDetailTypeEnum.DepositLineDetail;
            //    else
            //        line.DetailType = LineDetailTypeEnum.SalesItemLineDetail;
            //    line.DetailTypeSpecified = true;
            //    line.Id = (ivLineCnt + 1).ToString();

            //    lines[ivLineCnt] = line;
            //    ivLineCnt++;
            //}

            //creditMemo.Line = lines;

            //CreditMemo resultCreditMemo = services.Add(creditMemo) as CreditMemo;
        }

        /// <summary>
        /// CReate Journal Entry
        /// </summary>
        /// <param name="creditDT"></param>
        /// <param name="debitDT"></param>
        /// <param name="transDate"></param>
        public void CreateJournalEntry(DataTable creditDT, DataTable debitDT, DateTime transDate)
        {
            JournalEntry journalEntry = new JournalEntry();
            journalEntry.Adjustment = true;
            journalEntry.AdjustmentSpecified = true;

            journalEntry.DocNumber = "DocNumber" + Helper.GetGuid().Substring(0, 5);
            if (transDate == null)
                journalEntry.TxnDate = DateTime.UtcNow.Date;
            else
                journalEntry.TxnDate = transDate;
            journalEntry.TxnDateSpecified = true;

            List<Line> lineList = new List<Line>();
            foreach (DataRow creditRow in creditDT.Rows)
            {
                Line creditLine = new Line();
                creditLine.Description = ((string)creditRow["Memo"]);
                creditLine.Amount = decimal.Round((decimal)creditRow["Amount"], 2);
                creditLine.AmountSpecified = true;
                creditLine.DetailType = LineDetailTypeEnum.JournalEntryLineDetail;
                creditLine.DetailTypeSpecified = true;
                JournalEntryLineDetail journalEntryLineDetailCredit = new JournalEntryLineDetail();
                journalEntryLineDetailCredit.PostingType = PostingTypeEnum.Credit;
                journalEntryLineDetailCredit.PostingTypeSpecified = true;
                //Account assetAccount = Helper.FindOrAddAccount(servicecontext, AccountTypeEnum.OtherCurrentAsset, AccountClassificationEnum.Asset);
                journalEntryLineDetailCredit.AccountRef = new ReferenceType()
                {
                    type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Account),
                    name = (string)creditRow["GLAccount"]
                    //Value = assetAccount.Id
                };
                journalEntryLineDetailCredit.ClassRef = new ReferenceType()
                {
                    type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Class),
                    name = ((string)creditRow["ClassName"])
                };
                creditLine.AnyIntuitObject = journalEntryLineDetailCredit;
                lineList.Add(creditLine);
            }
            foreach (DataRow debitRow in debitDT.Rows)
            {
                Line debitLine = new Line();
                debitLine.Description = ((string)debitRow["Memo"]);
                debitLine.Amount = decimal.Round((decimal)debitRow["Amount"], 2);
                debitLine.AmountSpecified = true;
                debitLine.DetailType = LineDetailTypeEnum.JournalEntryLineDetail;
                debitLine.DetailTypeSpecified = true;
                JournalEntryLineDetail journalEntryLineDetail = new JournalEntryLineDetail();
                journalEntryLineDetail.PostingType = PostingTypeEnum.Debit;
                journalEntryLineDetail.PostingTypeSpecified = true;
                Account expenseAccount = Helper.FindOrAddAccount(servicecontext, AccountTypeEnum.Expense, AccountClassificationEnum.Expense);
                journalEntryLineDetail.AccountRef = new ReferenceType()
                {
                    type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Account),
                    name = (string)debitRow["GLAccount"]
                    //Value = expenseAccount.Id 
                };
                journalEntryLineDetail.ClassRef = new ReferenceType()
                {
                    type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Class),
                    name = ((string)debitRow["ClassName"])
                };
                debitLine.AnyIntuitObject = journalEntryLineDetail;
                lineList.Add(debitLine);
            }

            journalEntry.Line = lineList.ToArray();
            services.Add(journalEntry);
        }


        public string AddCustomer(DataTable dtCustomers)
        {
            if (services == null)
                CreateService();
            if (Terms.Count == 0)
                GetTerms();
            if (CustomerTypes.Count == 0)
                GetCustomerTypes();

            Customer customer = new Customer();
            customer.Organization = true;
            customer.OrganizationSpecified = true;
            DataRow dr = dtCustomers.Rows[0];
            customer.CompanyName = (string)dr["Customer_Name"];
            customer.DisplayName = (string)dr["Customer_Name"];

            customer.BillAddr = new PhysicalAddress()
            {
                Line1 = CheckNullDBString(dr["Customer_address_line1"]),
                Line2 = CheckNullDBString(dr["Customer_address_line2"]),
                CountrySubDivisionCode = CheckNullDBString(dr["Customer_address_state"]),
                City = CheckNullDBString(dr["Customer_address_city"]),
                PostalCode = CheckNullDBString(dr["Customer_address_zipcode"])
            };
            customer.ShipAddr = new PhysicalAddress()
            {
                Line1 = CheckNullDBString(dr["Customer_address_line1"]),
                Line2 = CheckNullDBString(dr["Customer_address_line2"]),
                CountrySubDivisionCode = CheckNullDBString(dr["Customer_address_state"]),
                City = CheckNullDBString(dr["Customer_address_city"]),
                PostalCode = CheckNullDBString(dr["Customer_address_zipcode"])
            };
            customer.PrimaryPhone = new TelephoneNumber()
            {
                FreeFormNumber = FormatPhoneNumber(CheckNullDBString(dr["Customer_corp_office_phone"]))
            };
            customer.Fax = new TelephoneNumber()
            {
                FreeFormNumber = FormatPhoneNumber(CheckNullDBString(dr["Customer_corp_fax"]))
            };
            customer.PrimaryEmailAddr = new EmailAddress()
            {
                Address = CheckNullDBString(dr["Customer_corp_email"])
            };



            if (dr["Customer_type_name"] != DBNull.Value)
            {
                customer.CustomerTypeRef = new ReferenceType();
                customer.CustomerTypeRef.name = ((string)dr["Customer_type_name"]).TrimEnd();
                if (!CustomerTypes.ContainsKey(((string)dr["Customer_type_name"]).TrimEnd()))
                {
                    return "Customer Type not in Types collection:" + ((string)dr["Customer_type_name"]).TrimEnd();
                }
                customer.CustomerTypeRef.Value = CustomerTypes[((string)dr["Customer_type_name"]).TrimEnd()];
                customer.CustomerTypeRef.type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.CustomerType);
            }


            if (dr["payment_term_erp_ID"] != DBNull.Value)
            {
                customer.SalesTermRef = new ReferenceType();
                customer.SalesTermRef.name = ((string)dr["payment_term_erp_ID"]).TrimEnd();
                if (!Terms.ContainsKey(((string)dr["payment_term_erp_ID"]).TrimEnd()))
                {
                    return "Payment Terms not in Terms collection:" + ((string)dr["payment_term_erp_ID"]).TrimEnd();
                }
                customer.SalesTermRef.Value = Terms[((string)dr["payment_term_erp_ID"]).TrimEnd()];
                customer.SalesTermRef.type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Term);
            }



            services.Add(customer);
            return "Customer Added";

        }

        public string AddVendor(DataTable dtVendors)
        {
            if (services == null)
                CreateService();
            if (Terms.Count == 0)
                GetTerms();

            foreach (DataRow drow in dtVendors.Rows)
            {
                Vendor vendor = new Vendor();
                vendor.DisplayName = (string)drow["vendor_name"];
                vendor.CompanyName = (string)drow["vendor_name"];

                vendor.BillAddr = new PhysicalAddress()
                {
                    Line1 = CheckNullDBString(drow["vendor_address_line1"]),
                    Line2 = CheckNullDBString(drow["vendor_address_line2"]),
                    Line3 = CheckNullDBString(drow["vendor_address_line3"]),
                    CountrySubDivisionCode = CheckNullDBString(drow["vendor_address_state"]),
                    City = CheckNullDBString(drow["vendor_address_city"]),
                    PostalCode = CheckNullDBString(drow["vendor_address_zipcode"])
                };
                vendor.PrimaryEmailAddr = new EmailAddress()
                {
                    Address = CheckNullDBString(drow["vendor_purchasing_email"])
                };
                vendor.PrimaryPhone = new TelephoneNumber()
                {
                    FreeFormNumber = FormatPhoneNumber(CheckNullDBString(drow["vendor_phone1"]))
                };
                vendor.AlternatePhone = new TelephoneNumber()
                {
                    FreeFormNumber = FormatPhoneNumber(CheckNullDBString(drow["vendor_phone2"]))
                };
                vendor.Fax = new TelephoneNumber()
                {
                    FreeFormNumber = FormatPhoneNumber(CheckNullDBString(drow["vendor_fax"]))
                };

                if (drow["vendor_tax_id"] != DBNull.Value)
                    vendor.TaxIdentifier = (string)drow["vendor_tax_id"];

                if (drow["vendor_account_number"] != DBNull.Value)
                    vendor.AcctNum = (string)drow["vendor_account_number"];

                if (drow["vendor_notes"] != DBNull.Value)
                    vendor.Notes = (string)drow["vendor_notes"];


                vendor.TermRef = new ReferenceType();
                vendor.TermRef.name = "Net Due Upon Receipt";

                if (!Terms.ContainsKey(vendor.TermRef.name))
                {
                    return "Payment Terms not in Terms collection:" + vendor.TermRef.name;
                }

                vendor.TermRef.Value = Terms[vendor.TermRef.name];
                vendor.TermRef.type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Term);

                services.Add(vendor);
            }
            return "Vendor Added";
        }

        public string GetCustomer(string customerID)
        {
            if (services == null)
                CreateService();
            Customer customer = new Customer();
            customer.Id = customerID;
            Customer thisCust = services.FindById(customer);
            return thisCust.DisplayName;
        }

        public string GetInvoice(string invoiceID)
        {
            if (Invoices.Count == 0)
                GetInvoiceList();
            Invoice thisInv = Invoices[invoiceID];
            return thisInv.invoiceStatus;
        }

        public List<string> GetCustomerList()
        {
            CustomerIDs.Clear();
            if (services == null)
                CreateService();
            List<string> custnames = new List<string>();
            Customer customer = new Customer();
            List<Customer> customers = services.FindAll(customer, 1, 1000).ToList<Customer>();

            foreach (Customer cust in customers)
            {
                custnames.Add(cust.FullyQualifiedName + ":" + cust.Id);
                CustomerIDs.Add(cust.FullyQualifiedName, cust.Id);
            }


            return custnames;
        }

        public List<string> GetInvoiceList()
        {
            if (services == null)
                CreateService();
            List<string> invoicesdata = new List<string>();
            Invoice invoice = new Invoice();
            List<Invoice> invoices = services.FindAll(invoice).ToList<Invoice>();
            foreach (Invoice inv in invoices)
            {
                invoicesdata.Add(inv.Id + ": " + inv.DocNumber);
                if (!Invoices.ContainsKey(inv.DocNumber))
                    Invoices.Add(inv.DocNumber, inv);

            }


            return invoicesdata;
        }

        public void GetProductAndServicesPrefs()
        {

            if (services == null)
                CreateService();
            Items.Clear();
            Item productAndServicesPref = new Item();
            List<Item> productAndServicesPrefs = services.FindAll(productAndServicesPref).ToList<Item>();
            foreach (Item item in productAndServicesPrefs)
            {
                Items.Add(item.Name, item.Id);
            }

        }

        public void GetClasses()
        {

            if (services == null)
                CreateService();
            Classes.Clear();
            Class JobTypeClass = new Class();
            List<Class> jobTypeClasses = services.FindAll(JobTypeClass).ToList<Class>();
            foreach (Class jobTypeClass in jobTypeClasses)
            {
                Classes.Add(jobTypeClass.Name, jobTypeClass.Id);
            }

        }

        public void GetTerms()
        {

            if (services == null)
                CreateService();
            Terms.Clear();
            Term paymentTerm = new Term();
            List<Term> paymentTerms = services.FindAll(paymentTerm).ToList<Term>();
            foreach (Term paymentTerm1 in paymentTerms)
            {
                Terms.Add(paymentTerm1.Name, paymentTerm1.Id);
            }

        }

        public void GetDepartments()
        {

            if (services == null)
                CreateService();
            Departments.Clear();
            Department departmentRef = new Department();
            List<Department> DepartmentRefs = services.FindAll(departmentRef).ToList<Department>();
            foreach (Department departmentRef1 in DepartmentRefs)
            {
                Departments.Add(departmentRef1.Name, departmentRef1.Id);
            }

        }

        public void GetCustomerTypes()
        {

            if (services == null)
                CreateService();
            CustomerTypes.Clear();
            CustomerType customerTypeRef = new CustomerType();
            List<CustomerType> CustomerTypeRefs = services.FindAll(customerTypeRef).ToList<CustomerType>();
            foreach (CustomerType customerTypeRef1 in CustomerTypeRefs)
            {
                CustomerTypes.Add(customerTypeRef1.Name, customerTypeRef1.Id);
            }

        }

        public void GetVendors()
        {

            if (services == null)
                CreateService();
            Vendors.Clear();
            Vendor vendorRef = new Vendor();
            List<Vendor> VendorRefs = services.FindAll(vendorRef).ToList<Vendor>();

            foreach (Vendor vendorRef1 in VendorRefs)
            {
                string vendorName = string.Empty;
                if (string.IsNullOrEmpty(vendorRef1.CompanyName))
                    vendorName = vendorRef1.DisplayName;
                else
                    vendorName = vendorRef1.CompanyName;
                Vendors.Add(vendorName, vendorRef1.Id);
            }

        }

        public void GetAccounts()
        {

            if (services == null)
                CreateService();
            Accounts.Clear();
            Account accountRef = new Account();
            List<Account> AccountRefs = services.FindAll(accountRef).ToList<Account>();

            foreach (Account accountRef1 in AccountRefs)
            {
                if (!string.IsNullOrEmpty(accountRef1.AcctNum))
                  Accounts.Add(accountRef1.AcctNum, accountRef1);
            }

        }

        private string CheckNullDBString(object datacol)
        {
            string dbString = string.Empty;
            if (!(datacol is DBNull))
            {
                dbString = (string)datacol;
            }
            return dbString.Trim();
        }
        private string FormatPhoneNumber(string phoneNumber)
        {
            string phoneString = string.Empty;
            if (phoneNumber.Trim().Length > 8)
                phoneString = string.Format("{0}-{1}-{2}", phoneNumber.Substring(0, 3), phoneNumber.Substring(3, 3), phoneNumber.Substring(6, 4));
            else
                phoneString = phoneNumber;
            return phoneString.Trim();
        }
    }
}

