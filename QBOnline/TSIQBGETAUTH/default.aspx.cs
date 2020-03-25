using Intuit.Ipp.OAuth2PlatformClient;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.UI;
using TSI.QBInterface;
using Intuit.Ipp.Exception;
using Intuit.Ipp.Core;
using Intuit.Ipp.Security;
using Intuit.Ipp.DataService;
using System.Text;
using System.Data;
using System.Windows.Forms;
using System.Diagnostics;


namespace TSIQBGETAUTH
{

    public partial class _default : System.Web.UI.Page
    {
        QBMethods qbMethods;


        
        //private string accessToken = string.Empty;
        //private string refreshToken = string.Empty;
        static string redirectURI = ConfigurationManager.AppSettings["redirectURI"];
        static string clientID = ConfigurationManager.AppSettings["clientID"];
        static string clientSecret = ConfigurationManager.AppSettings["clientSecret"];
        static string appEnvironment = ConfigurationManager.AppSettings["appEnvironment"];
        static string realmId = ConfigurationManager.AppSettings["realmId"];
        static string authCode;
        /*static OAuth2Client oauth2Client = new OAuth2Client(clientID, clientSecret, redirectURI, appEnvironment);*/ // environment is “sandbox” or “production”
        static string accessToken;
        static string refreshToken;


        protected void Page_Load(object sender, EventArgs e)
        {
            AsyncMode = true;
            //if (Request.QueryString["code"] != null)

            authCode = Request.QueryString["code"];
            if (!string.IsNullOrEmpty(authCode))
            {
                if (Page.IsAsync)
                {

                    PageAsyncTask t = new PageAsyncTask(GetTokens);
                    Page.RegisterAsyncTask(t);
                    Page.ExecuteRegisteredAsyncTasks();

                }

            }



        }


        private static async Task<int> GetTokens()
        {

            UpdateSecurityToAuthenticate(19);
            OAuth2Client oauth2Client = new OAuth2Client(clientID, clientSecret, redirectURI, appEnvironment); // environment is “sandbox” or “production”

            // Get OAuth2 Bearer token
            var tokenResponse = await oauth2Client.GetBearerTokenAsync(authCode);
            //retrieve access_token and refresh_token
            accessToken = tokenResponse.AccessToken;
            refreshToken = tokenResponse.RefreshToken;
            if (accessToken != string.Empty)
            {
                UpdateSecurityTokensInDB();
            }
            return 1;
        }

        private static void UpdateSecurityTokensInDB()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["JobScheduler2007"].ConnectionString;
            SqlConnection connection = new SqlConnection(connectionString);

            try
            {

                connection.Open();

                SqlCommand Updatecmd = new SqlCommand("Update QuickBooksSecurityTokens SET AccessToken = @AccessToken, RefreshToken = @RefreshToken, IsAuthenticate = 0 where IsAuthenticate = 1",
                    connection);

                int updCount = 0;

                Updatecmd.Parameters.Clear();
                Updatecmd.Parameters.AddWithValue("@AccessToken", accessToken);
                Updatecmd.Parameters.AddWithValue("@RefreshToken", refreshToken);
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

        private static void UpdateSecurityToAuthenticate(int employeeID)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["JobScheduler2007"].ConnectionString;
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

    }
}