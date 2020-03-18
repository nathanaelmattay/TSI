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
            servicecontext.IppConfiguration.AdvancedLogger.RequestAdvancedLog.ServiceRequestLoggingLocationForFile = @"C:\temp\Serilog_log";//Any drive logging location
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

        /// <summary>
        /// Create Payables Bill
        /// </summary>
        /// <param name="dtBill"></param>
        /// <param name="sentDtBillItems"></param>
        public void CreatePayablesBill(DataTable dtBill, DataTable sentDtBillItems)
        {
            Bill bill = new Bill();
            bill.DocNumber = Guid.NewGuid().ToString("N").Substring(0, 10);
            bill.TxnStatus = "Payable";

            bill.APAccountRef = new ReferenceType()
            {
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Account),
                name = "Account Payable",
                Value = "QB:1"
            };
            bill.VendorRef = new ReferenceType()
            {
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Vendor),
                //name = "sentVendorKey",
                Value = string.Concat("QB:", ((string)dtBill.Rows[0]["VendorName"]))
            };

            //	bill.TxnDate.SetValue((DateTime)dtBill.Rows[0]["BillDate"]);
            //	bill.Memo.SetValue((string)dtBill.Rows[0]["VendorBillMemo"]);
            //	bill.RefNumber.SetValue((string)dtBill.Rows[0]["VendorInvoiceNumber"]);

            bill.TxnTaxDetail = new TxnTaxDetail();
            bill.TxnTaxDetail.DefaultTaxCodeRef = new ReferenceType()
            {
                Value = "QB:123",
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.TaxCode),
                name = "TaxCodeName"
            };
            bill.TxnTaxDetail.TotalTax = new Decimal(0.00);
            bill.TxnTaxDetail.TotalTaxSpecified = true;

            Line[] lines = new Line[sentDtBillItems.Rows.Count];
            int billLineCnt = 0;
            foreach (System.Data.DataRow poItemrow in sentDtBillItems.Rows)
            {
                Line line1 = new Line();
                line1.Amount = (decimal)poItemrow["POLineQty"];
                line1.AmountSpecified = true;
                line1.Description = poItemrow["VendorItemNbr"].ToString();
                line1.LineNum = poItemrow["POLinePrintOrder"].ToString();
                line1.Id = (billLineCnt + 1).ToString();

                //PurchaseOrderItemLineDetail purchaseOrderItemLineDetail = new PurchaseOrderItemLineDetail();
                //purchaseOrderItemLineDetail.Qty = (decimal)poItemrow["POLineQty"];
                //purchaseOrderItemLineDetail.QtySpecified = true;
                //purchaseOrderItemLineDetail.AnyIntuitObject = (decimal)(poItemrow["POLinePrice"]);
                //purchaseOrderItemLineDetail.ItemElementName = ItemChoiceType.UnitPrice;
                //line1.AnyIntuitObject = purchaseOrderItemLineDetail;

                AccountBasedExpenseLineDetail accountBasedExpenseLineDetail = new AccountBasedExpenseLineDetail();
                accountBasedExpenseLineDetail.AccountRef = new ReferenceType()
                {
                    Value = poItemrow["AccountNumber"].ToString()
                };
                accountBasedExpenseLineDetail.ClassRef = new ReferenceType()
                {
                    Value = poItemrow["Class"].ToString()
                };
                accountBasedExpenseLineDetail.BillableStatus = BillableStatusEnum.Billable;
                line1.AnyIntuitObject = accountBasedExpenseLineDetail;

                lines[billLineCnt] = line1;
                billLineCnt++;
            }
            bill.Line = lines;

            Bill resultBill = services.Add(bill) as Bill;
        }

        /// <summary>
        /// Create Receivables Invoice
        /// </summary>
        /// <param name="dtInvoice"></param>
        /// <returns></returns>
        

        public string CreateReceivablesInvoice(DataTable dtInvoice)
        {
            if (services == null)
                CreateService();

            Invoice invoice = new Invoice();
            //invoice.TemplateRef.name = "Trinity Stairs Invoice";
            invoice.TemplateRef = new ReferenceType()
            {
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.CompanyInfo),
                name = "Type",
                Value = "Trinity Stairs Invoice"
            };
            invoice.Id = Convert.ToString(dtInvoice.Rows[0]["InvoiceID"]);
            invoice.CustomerRef = new ReferenceType()
            {
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Customer),
                //name = "Customer Name",
                Value = (string)dtInvoice.Rows[0]["CustomerName"]
            };

            //invoice.ClassRef.name = (string)dtInvoice.Rows[0]["RegionName"];
            invoice.ClassRef = new ReferenceType()
            {
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Class),
                //name = "Class Region",
                Value = (string)dtInvoice.Rows[0]["RegionName"]
            };

            if (dtInvoice.Rows[0]["SalesOrderNumber"] != DBNull.Value)
            {
                string tempPONumber = (string)dtInvoice.Rows[0]["SalesOrderNumber"];
                if (tempPONumber.Length > 25)
                    tempPONumber = tempPONumber.Remove(25);
                invoice.PONumber = tempPONumber;
            }
            invoice.TxnDate = (DateTime)dtInvoice.Rows[0]["InvoiceDate"];
            //invoice.DueDate.SetValue((DateTime)dtInvoice.Rows[0]["InvoiceDate"]);

            //if (dtInvoice.Rows[0]["ScheduledDate"] != DBNull.Value)
            //    invoice.Other.SetValue(((DateTime)dtInvoice.Rows[0]["ScheduledDate"]).ToShortDateString());
            invoice.ShipAddr = new PhysicalAddress()
            {

                Line1 = (string)dtInvoice.Rows[0]["AddressStreet"],
                City = (string)dtInvoice.Rows[0]["AddressCity"],
                CountrySubDivisionCode = (string)dtInvoice.Rows[0]["AddressState"],
                PostalCode = (string)dtInvoice.Rows[0]["AddressZip"]

            };
            

            //invoice.SalesTaxLineAdd.ORSalesTaxLineAdd.Amount.SetValue(Convert.ToDouble(dtInvoice.Rows[0]["InvoiceTax"]));
            // invoice.ShippingLineAdd.Amount.SetValue(Convert.ToDouble(dtInvoice.Rows[0]["InvoiceShipping"]));
            //invoice.ARAccountRef.name = (string)dtInvoice.Rows[0]["RegionARGLCode"];

            bool isTaxable = false;
            if (dtInvoice.Rows[0]["InvoiceTax"] != DBNull.Value && Convert.ToDouble(dtInvoice.Rows[0]["InvoiceTax"]) != 0)
            {
                invoice.TxnTaxDetail = new TxnTaxDetail();
                invoice.TxnTaxDetail.TotalTax = (decimal)dtInvoice.Rows[0]["InvoiceTax"];
                invoice.TxnTaxDetail.TotalTaxSpecified = true;
            }
               
            {
                isTaxable = true;
            }

            //Add Invoice Lines
            foreach (DataRow dtRow in dtInvoice.Rows)
            {
                Line invoiceLine = new Line();
                //invoiceLine.ItemRef.FullName.SetValue((string)dtRow["Jobtype"]);
                invoiceLine.Description = ((string)dtRow["InvoiceLineText"]);
                invoiceLine.Amount = (Convert.ToDecimal(dtRow["InvoiceLineTotalAmount"]));
                //if (isTaxable)
                //{
                //    invoiceLine..IsTaxable.SetValue(true);
                //}
            }

            //if (dtInvoice.Rows[0]["InvoiceDiscount"] != null && Convert.ToDouble(dtInvoice.Rows[0]["InvoiceDiscount"]) != 0)
            //{
            //    IORInvoiceLineAdd invoiceLine = invoice.ORInvoiceLineAddList.Append();
            //    invoiceLine.InvoiceLineAdd.ItemRef.FullName.SetValue((string)dtInvoice.Rows[0]["Jobtype"]);
            //    invoiceLine.InvoiceLineAdd.Desc.SetValue("Discount");
            //    invoiceLine.InvoiceLineAdd.Amount.SetValue(Convert.ToDouble(dtInvoice.Rows[0]["InvoiceDiscount"]));
            //    if (isTaxable)
            //    {
            //        invoiceLine.InvoiceLineAdd.IsTaxable.SetValue(true);
            //    }
            //}
            //if (dtInvoice.Rows[0]["InvoiceShipping"] != null && Convert.ToDouble(dtInvoice.Rows[0]["InvoiceShipping"]) != 0)
            //{
            //    IORInvoiceLineAdd invoiceLine = invoice.ORInvoiceLineAddList.Append();
            //    invoiceLine.InvoiceLineAdd.ItemRef.FullName.SetValue((string)dtInvoice.Rows[0]["Jobtype"]);
            //    invoiceLine.InvoiceLineAdd.Desc.SetValue("Shipping");
            //    invoiceLine.InvoiceLineAdd.Amount.SetValue(Convert.ToDouble(dtInvoice.Rows[0]["InvoiceShipping"]));
            //    if (isTaxable)
            //    {
            //        invoiceLine.InvoiceLineAdd.IsTaxable.SetValue(true);
            //    }
            //}




            // invoice.Id = sentInvKey;
            invoice.DocNumber = Guid.NewGuid().ToString("N").Substring(0, 10);

            //invoice.CustomerMemo = new MemoRef()
            //{
            //    Value = sentInvAddr
            //};

            //invoice.TotalAmt = sentInvAmount;
            invoice.TotalAmtSpecified = true;

           

            //invoice.shipping ?? = sentShippingAmount
            //invoice     = sentDiscountAmount;

            //invoice.DueDate = sentInvoiceDate;
            //invoice.DueDateSpecified = true;

            invoice.ARAccountRef = new ReferenceType()
            {
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Account),
                name = "Account Receivable",
                Value = string.Concat("QB:", (string)dtInvoice.Rows[0]["RegionARGLCode"])
            };


            invoice.TxnDate = DateTime.Today.Date;
            invoice.TxnDateSpecified = true;

            invoice.ApplyTaxAfterDiscount = false;
            invoice.ApplyTaxAfterDiscountSpecified = true;

            List<Line> invLines = new List<Line>();
            int ivLineCnt = 0;
            foreach (System.Data.DataRow dtRow in dtInvoice.Rows)
            {
                Line line = new Line();
                line.Description = (string)dtRow["InvoiceLineText"];
                line.Amount = decimal.Round((decimal)dtRow["InvoiceLineTotalAmount"], 2);
                line.AmountSpecified = true;
                line.DetailType = LineDetailTypeEnum.SalesItemLineDetail;
                line.DetailTypeSpecified = true;
                line.Id = (ivLineCnt + 1).ToString();
                ivLineCnt++;
                
                

                if ((string)dtRow["InvoiceLineText"] == "Sales")
                    line.DetailType = LineDetailTypeEnum.SalesItemLineDetail;
                else if ((string)dtRow["InvoiceLineText"] == "Shipping")
                    line.DetailType = LineDetailTypeEnum.SalesItemLineDetail;
                else if ((string)dtRow["InvoiceLineText"] == "Discount")
                    line.DetailType = LineDetailTypeEnum.DiscountLineDetail;
                else if ((string)dtRow["InvoiceLineText"] == "Deposit")
                    line.DetailType = LineDetailTypeEnum.DepositLineDetail;
                else
                    line.DetailType = LineDetailTypeEnum.SalesItemLineDetail;
                line.DetailTypeSpecified = true;
                line.Id = (ivLineCnt + 1).ToString();

                //Line Sales Item Line Detail
                SalesItemLineDetail lineSalesItemLineDetail = new SalesItemLineDetail();
                //Line Sales Item Line Detail - ItemRef
                lineSalesItemLineDetail.ItemRef = new ReferenceType()
                {
                    name = "Sales",
                    Value = "1"
                };
                //Line Sales Item Line Detail - UnitPrice
                lineSalesItemLineDetail.AnyIntuitObject = decimal.Round((decimal)dtRow["InvoiceLineTotalAmount"], 2);
                lineSalesItemLineDetail.ItemElementName = ItemChoiceType.UnitPrice;
                //Line Sales Item Line Detail - Qty
                lineSalesItemLineDetail.Qty = 1;
                lineSalesItemLineDetail.QtySpecified = true;
                //Line Sales Item Line Detail - TaxCodeRef
                //For US companies, this can be 'TAX' or 'NON'
                lineSalesItemLineDetail.TaxCodeRef = new ReferenceType()
                {
                    Value = "TAX"
                };
                //Line Sales Item Line Detail - ServiceDate 
                lineSalesItemLineDetail.ServiceDate = DateTime.Now.Date;
                lineSalesItemLineDetail.ServiceDateSpecified = true;
                //Assign Sales Item Line Detail to Line Item
                line.AnyIntuitObject = lineSalesItemLineDetail;

                invLines.Add(line);

                line.DetailTypeSpecified = true;
                line.Id = (ivLineCnt + 1).ToString();
                ivLineCnt++;
                invLines.Add(line);

            }

            invoice.Line = invLines.ToArray();

            Invoice resultBill = services.Add(invoice) as Invoice;
            return "Invoice Loaded";
        }


        public void CreateReceivablesCreditMemo(DataTable dtInvoice)
        {
            CreditMemo creditMemo = new CreditMemo();
            creditMemo.DocNumber = Guid.NewGuid().ToString("N").Substring(0, 10);

            creditMemo.CustomerRef = new ReferenceType()
            {
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Customer),
                //name = "TestCustomer",
                Value = string.Concat("QB:", (string)dtInvoice.Rows[0]["CustomerName"])
            };

            if (dtInvoice.Rows[0]["SalesOrderNumber"] != DBNull.Value)
                creditMemo.PONumber = (string)dtInvoice.Rows[0]["SalesOrderNumber"];
            //creditMemo.CustomerMemo = new MemoRef()
            //{
            //    Value = sentInvAddr
            //};

            //creditMemo.TotalAmt = sentInvAmount;
            creditMemo.TotalAmtSpecified = false;

            if (dtInvoice.Rows[0]["InvoiceTax"] != null && Convert.ToDecimal(dtInvoice.Rows[0]["InvoiceTax"]) != 0)
            {
                creditMemo.TxnTaxDetail = new TxnTaxDetail();
                creditMemo.TxnTaxDetail.TotalTax = (decimal)dtInvoice.Rows[0]["InvoiceTax"];
                creditMemo.TxnTaxDetail.TotalTaxSpecified = true;
            }

            creditMemo.DueDate = (DateTime)dtInvoice.Rows[0]["InvoiceDate"];
            creditMemo.DueDateSpecified = true;

            creditMemo.ARAccountRef = new ReferenceType()
            {
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Account),
                name = "Account Receivable",
                Value = "QB:37"
            };

            creditMemo.ClassRef = new ReferenceType()
            {
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Class),
                name = "Class ??",
                Value = "QB:37"
            };
            creditMemo.TxnDate = DateTime.Today.Date;
            creditMemo.TxnDateSpecified = true;

            Line[] lines = new Line[dtInvoice.Rows.Count];
            int ivLineCnt = 0;
            foreach (System.Data.DataRow dtRow in dtInvoice.Rows)
            {
                Line line = new Line();
                line.Amount = decimal.Round((decimal)dtRow["InvoiceAmount"], 2);
                line.AmountSpecified = true;
                line.Description = (string)dtRow["DistType"];

                if ((string)dtRow["DistType"] == "Sales")
                    line.DetailType = LineDetailTypeEnum.SalesItemLineDetail;
                else if ((string)dtRow["DistType"] == "Shipping")
                    line.DetailType = LineDetailTypeEnum.SalesItemLineDetail;
                else if ((string)dtRow["DistType"] == "Discount")
                    line.DetailType = LineDetailTypeEnum.DiscountLineDetail;
                else if ((string)dtRow["DistType"] == "Deposit")
                    line.DetailType = LineDetailTypeEnum.DepositLineDetail;
                else
                    line.DetailType = LineDetailTypeEnum.SalesItemLineDetail;
                line.DetailTypeSpecified = true;
                line.Id = (ivLineCnt + 1).ToString();

                lines[ivLineCnt] = line;
                ivLineCnt++;
            }

            creditMemo.Line = lines;

            CreditMemo resultCreditMemo = services.Add(creditMemo) as CreditMemo;
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

        /// <summary>
        /// Create new Customer
        /// </summary>
        /// <param name="dtCustomers"></param>
        /// <returns></returns>
        public string AddCustomer(DataTable dtCustomers)
        {
            if (services == null)
                CreateService();

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

            string custTypeRef = string.Empty;
            if (CheckNullDBString(dr["Customer_type_name"]) == "Production")
                custTypeRef = "800000000000226295";
            else if (CheckNullDBString(dr["Customer_type_name"]) == "Custom")
                custTypeRef = "800000000000226296";
            else if (CheckNullDBString(dr["Customer_type_name"]) == "Homeowner")
                custTypeRef = "800000000000226297";
            if (!string.IsNullOrEmpty(custTypeRef))
            {
                customer.CustomerTypeRef = new ReferenceType()
                {
                    Value = custTypeRef
                };
            }

            string salesTypeRef = string.Empty;
            if (CheckNullDBString(dr["payment_term_erp_ID"]).Trim() == "Net Due Upon Receipt")
                salesTypeRef = "41";
            else if (CheckNullDBString(dr["payment_term_erp_ID"]).Trim() == "Net 15")
                salesTypeRef = "35";
            else if (CheckNullDBString(dr["payment_term_erp_ID"]).Trim() == "50%Deposit/Net 15")
                salesTypeRef = "59";
            else if (CheckNullDBString(dr["payment_term_erp_ID"]).Trim() == "50%Deposit/Net Due")
                salesTypeRef = "60";
            if (!string.IsNullOrEmpty(salesTypeRef))
            {
                customer.SalesTermRef = new ReferenceType()
                {
                    Value = salesTypeRef
                };
            }

            services.Add(customer);
            return "Customer Added";
        }

        /// <summary>
        /// Create New Vendor
        /// </summary>
        /// <param name="dtVendors"></param>
        /// <returns></returns>
        public string AddVendor(DataTable dtVendors)
        {
            foreach (DataRow drow in dtVendors.Rows)
            {
                Vendor vendor = new Vendor();
                vendor.DisplayName = (string)drow["vendor_name"];
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
                string salesTypeRef = string.Empty;
                if (CheckNullDBString(drow["payment_term_erp_ID"]).Trim() == "Net Due Upon Receipt")
                    salesTypeRef = "41";
                else if (CheckNullDBString(drow["payment_term_erp_ID"]).Trim() == "Net 15")
                    salesTypeRef = "35";
                else if (CheckNullDBString(drow["payment_term_erp_ID"]).Trim() == "50%Deposit/Net 15")
                    salesTypeRef = "59";
                else if (CheckNullDBString(drow["payment_term_erp_ID"]).Trim() == "50%Deposit/Net Due")
                    salesTypeRef = "60";
                if (!string.IsNullOrEmpty(salesTypeRef))
                {
                    vendor.TermRef = new ReferenceType()
                    {
                        Value = salesTypeRef
                    };
                }

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

        public List<string> GetCustomerList()
        {
            if (services == null)
                CreateService();
            List<string> custnames = new List<string>();
            Customer customer = new Customer();
            List<Customer> customers = services.FindAll(customer, 1, 1000).ToList<Customer>();
            foreach (Customer cust in customers)
            {
                custnames.Add(cust.FullyQualifiedName + ":" + cust.Id);
            }


            return custnames;
        }

        public List<string> GetInvoiceList()
        {
            if (services == null)
                CreateService();
            List<string> invoicesdata = new List<string>();
            Invoice invoice = new Invoice();
            List<Invoice> invoices = services.FindAll(invoice, 1, 100).ToList<Invoice>();
            foreach (Invoice inv in invoices)
            {
                invoicesdata.Add(inv.Id + ": " + inv.DocNumber);
            }


            return invoicesdata;
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
