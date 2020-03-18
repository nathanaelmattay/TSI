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

        static string authCode;
        static string idToken;
        public static IList<JsonWebKey> keys;
        public static Dictionary<string, string> dictionary = new Dictionary<string, string>();
        public ServiceContext servicecontext;
        public DataService services;
        public static OAuth2Client oauthClient = new OAuth2Client(clientID, clientSecret, redirectURI, appEnvironment);
        public static string AccessToken;
        public static string RefreshToken;
<<<<<<< HEAD
        private static Dictionary<string,string> CustomerIDs = new Dictionary<string, string>();
        private static Dictionary<string, Invoice> Invoices = new Dictionary<string, Invoice>();
        private static Dictionary<string, string> Items = new Dictionary<string, string>();
        private static Dictionary<string, string> Classes = new Dictionary<string, string>();
        private static Dictionary<string, string> Terms = new Dictionary<string, string>();
        private static Dictionary<string, string> Departments = new Dictionary<string, string>();
=======
>>>>>>> 8bc66b9d9ec2097ad29d6c11f6a9edd92e9b94d7
        //private string authorizeUrl;
        public QBMethods()
        {



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
            string connectionString = ConfigurationManager.ConnectionStrings["DBConnectionString"].ConnectionString;
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
<<<<<<< HEAD
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
            
=======

>>>>>>> 8bc66b9d9ec2097ad29d6c11f6a9edd92e9b94d7
            Invoice invoice = new Invoice();
            invoice.TemplateRef = new ReferenceType()
            {
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.CompanyInfo),
                name = "Type",
                Value = "Trinity Stairs Invoice"
            };
            invoice.DocNumber = Convert.ToString(dtInvoice.Rows[0]["InvoiceID"]);
            invoice.CustomerRef = new ReferenceType()
            {
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Customer),
                //name = "Customer Name",
                name = (string)dtInvoice.Rows[0]["CustomerName"]
            };

            invoice.ClassRef = new ReferenceType()
            {
                type = Enum.GetName(typeof(objectNameEnumType), objectNameEnumType.Class),
                //name = "Class Region",
                name = (string)dtInvoice.Rows[0]["RegionName"]
            };

            if (dtInvoice.Rows[0]["SalesOrderNumber"] != DBNull.Value)
            {
                string tempPONumber = (string)dtInvoice.Rows[0]["SalesOrderNumber"];
                if (tempPONumber.Length > 25)
                    tempPONumber = tempPONumber.Remove(25);
                invoice.PONumber = tempPONumber;
            }
            invoice.TxnDate = (DateTime)dtInvoice.Rows[0]["InvoiceDate"];
<<<<<<< HEAD
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
            
=======

            invoice.DueDate = (DateTime)dtInvoice.Rows[0]["InvoiceDate"];
>>>>>>> 8bc66b9d9ec2097ad29d6c11f6a9edd92e9b94d7

            //if (dtInvoice.Rows[0]["ScheduledDate"] != DBNull.Value)
            //    invoice.Other.SetValue(((DateTime)dtInvoice.Rows[0]["ScheduledDate"]).ToShortDateString());
            invoice.ShipAddr = new PhysicalAddress()
            {
                Line1 = (string)dtInvoice.Rows[0]["AddressStreet"],
                CountrySubDivisionCode = (string)dtInvoice.Rows[0]["AddressState"],
                City = (string)dtInvoice.Rows[0]["AddressCity"],
                PostalCode = (string)dtInvoice.Rows[0]["AddressZip"]
            };

            invoice.TxnTaxDetail = new TxnTaxDetail();
            invoice.TxnTaxDetail.TotalTax = (decimal)dtInvoice.Rows[0]["InvoiceTax"];
            invoice.TxnTaxDetail.TotalTaxSpecified = true;


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

<<<<<<< HEAD
              

                //invoiceLine.

=======
>>>>>>> 8bc66b9d9ec2097ad29d6c11f6a9edd92e9b94d7
                SalesItemLineDetail lineSalesItemLineDetail = new SalesItemLineDetail();
                lineSalesItemLineDetail.ItemRef = new ReferenceType()
                {
                    name = (string)dtRow["JobType"]
                };
                lineSalesItemLineDetail.ClassRef = new ReferenceType()
                {
                    name = (string)dtRow["RegionName"]
                };
                //Line Sales Item Line Detail - TaxCodeRef
                //For US companies, this can be 'TAX' or 'NON'
                lineSalesItemLineDetail.TaxCodeRef = new ReferenceType();
                if ((decimal)dtRow["InvoiceTax"] > 0)
                {
                    lineSalesItemLineDetail.TaxCodeRef.Value = "TAX";
                }
                else
                {
                    lineSalesItemLineDetail.TaxCodeRef.Value = "NON";
                }
                //Line Sales Item Line Detail - ServiceDate 
                lineSalesItemLineDetail.ServiceDate = DateTime.Now.Date;
                lineSalesItemLineDetail.ServiceDateSpecified = true;
                //Assign Sales Item Line Detail to Line Item
                invoiceLine.AnyIntuitObject = lineSalesItemLineDetail;

                invLines.Add(invoiceLine);
            }

            if (dtInvoice.Rows[0]["InvoiceDiscount"] != null && Convert.ToDecimal(dtInvoice.Rows[0]["InvoiceDiscount"]) != 0)
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
            if ((decimal)dtInvoice.Rows[0]["InvoiceShipping"] > 0)
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


            invoice.Line = invLines.ToArray();

            Invoice resultBill = services.Add(invoice) as Invoice;

            return "Invoice Loaded";
        }

        /// <summary>
        /// Create Receivables Credit Memo
        /// </summary>
        /// <param name="dtInvoice"></param>
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
            Customer customer = new Customer();
            customer.Id = customerID;
            Customer thisCust = services.FindById(customer);
            return thisCust.DisplayName;
        }

        public List<string> GetCustomerList()
        {
            List<string> custnames = new List<string>();
            Customer customer = new Customer();
            List<Customer> customers = services.FindAll(customer, 1, 1000).ToList<Customer>();
            foreach (Customer cust in customers)
            {
                custnames.Add(cust.FullyQualifiedName);
            }


            return custnames;
        }

        public List<string> GetInvoiceList()
        {
            List<string> invoicesdata = new List<string>();
            Invoice invoice = new Invoice();
            List<Invoice> invoices = services.FindAll(invoice, 1, 100).ToList<Invoice>();
            foreach (Invoice inv in invoices)
            {
                invoicesdata.Add(inv.Id + ": " + inv.DocNumber);
            }


            return invoicesdata;
        }

<<<<<<< HEAD
        public void GetProductAndServicesPrefs()
        {

            if (services == null)
                CreateService();
            Items.Clear();
            Item productAndServicesPref = new Item();
            List<Item> productAndServicesPrefs = services.FindAll(productAndServicesPref).ToList<Item>();
            foreach (Item item in productAndServicesPrefs)
            {
                Items.Add(item.Name,item.Id);
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
=======
>>>>>>> 8bc66b9d9ec2097ad29d6c11f6a9edd92e9b94d7
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
