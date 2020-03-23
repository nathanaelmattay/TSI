using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using TSI.QBInterface;
using Intuit.Ipp.OAuth2PlatformClient;
using System.Threading.Tasks;
using Intuit.Ipp.Core;
using Intuit.Ipp.Data;
using Intuit.Ipp.DataService;
using Intuit.Ipp.QueryFilter;
using Intuit.Ipp.Security;
using Intuit.Ipp.Exception;
using Intuit.Ipp.ReportService;
using System.Diagnostics;
using System.Web.UI;
using Intuit.Ipp.Core.Configuration;
using System.Net;
using System.Web;



namespace JSQBTest
{
    public partial class Form1 : Form
    {
        QBMethods qbMethods;

        public Form1()
        {
            InitializeComponent();
            qbMethods = new QBMethods();
            QBMethods.GetDBTokens(19);
        }

        static string redirectURI = ConfigurationManager.AppSettings["redirectURI"];
        static string clientID = ConfigurationManager.AppSettings["clientID"];
        static string clientSecret = ConfigurationManager.AppSettings["clientSecret"];
        static string logPath = ConfigurationManager.AppSettings["logPath"];
        static string appEnvironment = ConfigurationManager.AppSettings["appEnvironment"];
        static string accessToken;
        static string refreshToken;
        static string authCode;
        public static IList<JsonWebKey> keys;
        /// <summary>
        /// Load Customers
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            CheckTokens();
            SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["JobScheduler2007"].ConnectionString);
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            try
            {
                command.CommandText = "Select C.*, R.region_name from Customers C inner join Regions R on R.region_id = C.Region_id";
                command.CommandType = CommandType.Text;
                command.Connection.Open();

                DataTable dtCustomers = new DataTable();

                dtCustomers.Load(command.ExecuteReader());

                

                string returnString = string.Empty;
                //returnString = qbMethods.AddCustomer((dtCustomers));

                textBox1.Text = returnString;

                command.Connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message + Environment.NewLine + ex.StackTrace);
            }
            finally
            {
                command.Connection.Close();
            }
        }


        /// <summary>
        /// Delete Customers
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            CheckTokens();
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                //textBox1.Text = qbMethods.DeleteCustomers();
                Cursor.Current = Cursors.Arrow;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message + Environment.NewLine + ex.StackTrace);
            }

        }

        /// <summary>
        /// Load vendors
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            CheckTokens();
            SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["JobScheduler2007"].ConnectionString);
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            try
            {
                command.CommandText = "Select * from Vendors2";
                command.CommandType = CommandType.Text;
                command.Connection.Open();

                DataTable dtVendors = new DataTable();

                dtVendors.Load(command.ExecuteReader());

                

                string returnString = string.Empty;
                //  returnString = qbMethods.CreateNewVendor(dtVendors);

                textBox1.Text = returnString;

                command.Connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message + Environment.NewLine + ex.StackTrace);
            }
            finally
            {
                command.Connection.Close();
            }
        }

        /// <summary>
        ///  Delete All Vendors
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            CheckTokens();
            try
            {
                //    textBox1.Text = qbMethods.DeleteVendors();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message + Environment.NewLine + ex.StackTrace);
            }
        }

        /// <summary>
        /// Load one Vendor based on ID
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            CheckTokens();
            SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["JobScheduler2007"].ConnectionString);
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            try
            {
                int vInt;
                if (!int.TryParse(txtNunber.Text, out vInt))
                    return;
                command.CommandText = "Select * from Vendors where vendor_id = " + vInt.ToString();
                command.CommandType = CommandType.Text;
                command.Connection.Open();

                DataTable dtVendors = new DataTable();

                dtVendors.Load(command.ExecuteReader());

               

                string returnString = string.Empty;
                returnString = qbMethods.AddVendor(dtVendors);

                textBox1.Text = returnString;

                command.Connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message + Environment.NewLine + ex.StackTrace);
            }
            finally
            {
                command.Connection.Close();
            }
        }

        /// <summary>
        /// Add one customer
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e)
        {
            CheckTokens();
            SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["JobScheduler2007"].ConnectionString);
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            try
            {
                int vInt;
                if (!int.TryParse(txtNunber.Text, out vInt))
                    return;
                command.CommandText = "Select C.*, R.region_name, CT.customer_type_name, PT.payment_term_erp_ID from Customers C inner join Regions R on R.region_id = C.Region_id " +
                    " inner join Customer_Types CT on CT.customer_type_id = C.customer_type_id " +
                    " left outer join payment_terms PT on PT.payment_term_id = C.payment_term_id where customer_id = " + vInt.ToString();
                command.CommandType = CommandType.Text;
                command.Connection.Open();

                DataTable dtCustomers = new DataTable();

                dtCustomers.Load(command.ExecuteReader());

                

                string returnString = string.Empty;

                returnString = qbMethods.AddCustomer(dtCustomers);

                textBox1.Text = returnString;

                command.Connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message + Environment.NewLine + ex.StackTrace);
            }
            finally
            {
                command.Connection.Close();
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            CheckTokens();
            SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["JobScheduler2007"].ConnectionString);
            string _sql = "dbo.tspInvoices2_GetInvoiceByID";
            SqlCommand command = new SqlCommand(_sql, connection);
            try
            {
                int vInt;
                if (!int.TryParse(txtNunber.Text, out vInt))
                    return;

                DataTable dt = new DataTable("dtInvoice");
                command.CommandType = CommandType.StoredProcedure;
                connection.Open();
                command.Parameters.Add("@InvoiceID", SqlDbType.Int).Value = vInt;
                dt.Load(command.ExecuteReader());

                

                string returnString = string.Empty;

                returnString = qbMethods.CreateReceivablesInvoice(dt);

                textBox1.Text = returnString;

                command.Connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message + Environment.NewLine + ex.StackTrace);
            }
            finally
            {
                command.Connection.Close();
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            CheckTokens();
            try
            {


                

                string returnString = string.Empty;
                //returnString = qbMethods.CreateJournalEntry("Test Job!", "Frisco", "1210", 
                //    "5010", Convert.ToDecimal(this.txtNunber.Text));

                textBox1.Text = returnString;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message + Environment.NewLine + ex.StackTrace);
            }
            finally
            {

            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            CheckTokens();
            List<string> custList = qbMethods.GetCustomerList();
            StringBuilder custnames = new StringBuilder();

            foreach (string cname in custList)
            {
                custnames.AppendLine(cname);
            }

            textBox1.Text = custnames.ToString();
        }

        private void button10_Click(object sender, EventArgs e)
        {

            CheckTokens();
            List<string> invList = qbMethods.GetInvoiceList();
            StringBuilder invNumbers = new StringBuilder();

            foreach (string iname in invList)
            {
                invNumbers.AppendLine(iname);
            }

            textBox1.Text = invNumbers.ToString();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            CheckTokens();
            string cust = qbMethods.GetCustomer(txtNunber.Text);

            textBox1.Text = cust;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

            AuthenticateQB();



            

        }



        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)
        {
            authCode = textBox2.Text;
            Task<int> task = GetTokens();
            textBox1.Text = "Tokens Exchanged! You are Authenticated.";

            
        }
        private static async Task<int> GetTokens()

        {
            QBMethods.UpdateSecurityToAuthenticate(19);
            OAuth2Client oauth2Client = new OAuth2Client(clientID, clientSecret, redirectURI, appEnvironment); // environment is “sandbox” or “production”

            // Get OAuth2 Bearer token
            var tokenResponse = await oauth2Client.GetBearerTokenAsync(authCode);
            //retrieve access_token and refresh_token
            accessToken = tokenResponse.AccessToken;
            refreshToken = tokenResponse.RefreshToken;
            QBMethods.AccessToken = tokenResponse.AccessToken;
            QBMethods.RefreshToken = tokenResponse.RefreshToken;
            if (accessToken != string.Empty)
            {
             QBMethods.UpdateSecurityTokensInDB();
            }

            QBMethods.AccessToken = accessToken;
            QBMethods.RefreshToken = refreshToken;
            QBMethods.CreateService();
            return 1;
        }

        private void txtNunber_TextChanged(object sender, EventArgs e)
        {

        }

        private void CheckTokens()
        {
            if (string.IsNullOrEmpty(QBMethods.AccessToken))
            {
                AuthenticateQB();
            }
        }

        private void AuthenticateQB()
        {
            OAuth2Client oauthClient = new OAuth2Client(clientID, clientSecret, redirectURI, appEnvironment);
            List<OidcScopes> scopes = new List<OidcScopes>();
            scopes.Add(OidcScopes.Accounting);
            var authorizeUrl = oauthClient.GetAuthorizationURL(scopes);
            ProcessStartInfo sInfo = new ProcessStartInfo(authorizeUrl);
            
            Process.Start(sInfo);
         
        }



    

        private void button13_Click(object sender, EventArgs e)
        {
            CheckTokens();
            string Inv = qbMethods.GetInvoice(txtNunber.Text);

            textBox1.Text = Inv;
        }

        private void button14_Click(object sender, EventArgs e)
        {
            CheckTokens();
            SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["JobScheduler2007"].ConnectionString);
            string _sql = "dbo.tspVendorBillsAndLines_GetByVendorBillID";
            SqlCommand command = new SqlCommand(_sql, connection);
            try
            {
                int vInt;
                if (!int.TryParse(txtNunber.Text, out vInt))
                    return;

                DataTable dt = new DataTable("dtBill");
                command.CommandType = CommandType.StoredProcedure;
                connection.Open();
                command.Parameters.Add("@VendorBillID", SqlDbType.Int).Value = vInt;
                dt.Load(command.ExecuteReader());



                string returnString = string.Empty;

                returnString = qbMethods.CreatePayablesBill(dt);

                textBox1.Text = returnString;

                command.Connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message + Environment.NewLine + ex.StackTrace);
            }
            finally
            {
                command.Connection.Close();
            }

        }
    }
}

