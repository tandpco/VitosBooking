using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;

public partial class Master_Page_Employee : System.Web.UI.MasterPage
{
    public string EmployeeID;

    public string masterpage_empid, masterpage_name, masterpage_driverstatus, masterpage_lastname, master_birthdate, masterpage_storeid; 


    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            if (!string.IsNullOrEmpty(Request.Form["EmployeeID"]))

                EmployeeID = Request.Form["EmployeeID"];
            else
                EmployeeID = Request.QueryString["EmployeeID"];

            
            GetEmployeeInformation();
                      
        }
    }

    public void GetEmployeeInformation()
    {
        #region<<ASP Code>>
        //Dim  RSEmp

        //'SQL="Select * from tblEmployee where EmployeeID = " & Session("EmployeeID") &" and StoreID = "&Session("StoreID")&" and IsActive='True'"
        //SQL="Select * from tblEmployee where EmployeeID = " & Session("EmployeeID") &" and IsActive='True'"

        //Set RSEmp=Conn.Execute(SQL)

        //If RSEmp.EOF And RSEmp.BOF Then  'Not Found REDIRECT back to sign in page
        //    NotificationMsg = "That EmployeeID was not found or not authorized"
        //    Response.Redirect("notification.asp?msg=" & Server.URLEncode(NotificationMsg)&"&NextPage=default.asp")
        //End If

        //Session("EmpID") = RSEmp("EmpID")
        #endregion

        using (System.Data.SqlClient.SqlConnection con = new SqlConnection(Convert.ToString(ConfigurationManager.AppSettings["DBConnectionstring"])))
        {
            string sqlst = "Select * from tblEmployee where EmpID = " + EmployeeID + " and IsActive='True'"; //'If used swipe card
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    con.Open();
                    cmd.Connection = con;
                    cmd.CommandText = sqlst;
                    cmd.CommandType = CommandType.Text;
                    SqlDataReader dr = cmd.ExecuteReader();
                    if (dr.Read())
                    {
                        masterpage_empid = dr["EmpID"].ToString();
                        //Session["EmployeeID"] = dr["EmployeeID"];

                        masterpage_name = dr["FirstName"].ToString() + " " + dr["LastName"].ToString();
                        masterpage_driverstatus = dr["DriverStatus"].ToString();
                        masterpage_lastname = dr["LastName"].ToString();
                        masterpage_storeid = dr["StoreID"].ToString();
                        string getbirthdate = dr["BirthDate"].ToString();

                        if (dr["Birthdate"].ToString() == string.Empty)
                        {
                            string NotificationMsg = "You cannot sign in without a valid birthdate entered in the system, please contact your supervisor";
                            Response.Redirect("notification.asp?msg=" + Server.UrlEncode(NotificationMsg) + "&NextPage=default.asp");
                        }
                        if (!string.IsNullOrEmpty(dr["Birthdate"].ToString()))
                        {
                            Int32 intAge;
                            int d1 = DateTime.Now.Year;
                            DateTime dt = Convert.ToDateTime(dr["Birthdate"].ToString());
                            int d2 = dt.Year;
                            intAge = d1 - d2;

                            int dt1 = dt.CompareTo(DateTime.Now);

                            if (dt1 < 0)
                            {
                                intAge = intAge - 1;
                            }
                            //Session["intAge"] = intAge;
                            
                        }

                        #region << ASP code>>
                        //    intAge = DateDiff("yyyy", RSEmp("Birthdate"), now())
                        //    If Now() < DateSerial(Year(now()), Month(RSEmp("Birthdate")), Day(RSEmp("Birthdate"))) Then
                        //        intAge = intAge - 1
                        //    End If
                        //session("intAge") = intAge
                        #endregion
                    }
                    else
                    {
                        string NotificationMsg = "Not Found REDIRECT back to sign in page";
                        Response.Redirect("notification.asp?msg=" + Server.UrlEncode(NotificationMsg) + "&NextPage=default.asp");
                    }
                }
                catch (Exception Ex)
                {
//                    ErrorLog.WriteToFile(Server.MapPath("ErrorLog/ErrorLog.txt"), Ex.Message);
                }
                finally
                {
                    con.Close();

                }
            }
        }

    }
}
