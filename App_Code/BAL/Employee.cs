using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Web.UI.WebControls;
using Microsoft.ApplicationBlocks.Data;

/// <summary>
/// Summary description for Employee
/// </summary>
public class Employee
{
	public Employee()
	{
		//
		// TODO: Add constructor logic here
		//
    }

    private string storeId;
    private string firstname;
    private string lastname;
    private string middlename;
    private string address;
    private string city;
    private string state;
    private string zip;
    private string homephone;
    private string cellphone;
    private string cellphonecarrierid;
    private string emailaddress;
    private int systemroleid;
    private DateTime birthdate;
    private DateTime hiredate;
    private DateTime enddate;
    private string maritalstatus;
    private int noofdependents;
    private string insurance;
    private DateTime insurancedate;
    private string additionalwithholdings;
    private string drivereligible;
    private string driverstatus;
    private bool isactive;
    private string owesforuniforms;
    private string owesforcashadvance;
    private string employeenotes;
    private int employeeid;
    private string employeeusername;
    private string gethiredate;
    private string getinsurancedate;
    private string teamname;
    private int rank;
    private int parentid;
    private int employeeloginid;
    private int socialsecuritynumber;
    

    #region Properties

    public string StoreId
    {
        get
        {
            return storeId;
        }
        set
        {
            storeId = value;
        }
    }

    public string FirstName
    {
        get
        {
            return firstname;
        }
        set
        {
            firstname = value;
        }
    }

    public string MiddleName
    {
        get
        {
            return middlename;
        }
        set
        {
            middlename = value;
        }
    }

    public string LastName
    {
        get
        {
            return lastname;
        }
        set
        {
            lastname = value;
        }
    }

    public string Address
    {
        get
        {
            return address;
        }
        set
        {
            address = value;
        }
    }

    public string City
    {
        get
        {
            return city;
        }
        set
        {
            city = value;
        }
    }

    public string State
    {
        get
        {
            return state;
        }
        set
        {
            state = value;
        }
    }

    public string Zip
    {
        get
        {
            return zip;
        }
        set
        {
            zip = value;
        }
    }

    public string HomePhone
    {
        get
        {
            return homephone;
        }
        set
        {
            homephone = value;
        }
    }

    public string CellPhone
    {
        get
        {
            return cellphone;
        }
        set
        {
            cellphone = value;
        }
    }

    public string CellPhoneCarrierId
    {
        get
        {
            return cellphonecarrierid;
        }
        set
        {
            cellphonecarrierid = value;
        }
    }

    public string EmailAddress
    {
        get
        {
            return emailaddress;
        }
        set
        {
            emailaddress = value;
        }
    }

    public int SystemRoleId
    {
        get
        {
            return systemroleid;
        }
        set
        {
            systemroleid = value;
        }
    }

    public DateTime BirthDate
    {
        get
        {
            return birthdate;
        }
        set
        {
            birthdate = value;
        }
    }

    public DateTime HireDate
    {
        get
        {
            return hiredate;
        }
        set
        {
            hiredate = value;
        }
    }

    public DateTime EndDate
    {
        get
        {
            return enddate;
        }
        set
        {
            enddate = value;
        }
    }

    public string MaritalStatus
    {
        get
        {
            return maritalstatus;
        }
        set
        {
            maritalstatus = value;
        }
    }

    public int NumberOfDependents
    {
        get
        {
            return noofdependents;
        }
        set
        {
            noofdependents = value;
        }
    }

    public string Insurance
    {
        get
        {
            return insurance;
        }
        set
        {
            insurance = value;
        }
    }

    public DateTime InsuranceDate
    {
        get
        {
            return insurancedate;
        }
        set
        {
            insurancedate = value;
        }
    }

    public string AdditionalWithHoldings
    {
        get
        {
            return additionalwithholdings;
        }
        set
        {
            additionalwithholdings = value;
        }
    }

    public string DriverEligible
    {
        get
        {
            return drivereligible;
        }
        set
        {
            drivereligible = value;
        }
    }

    public string DriverStatus
    {
        get
        {
            return driverstatus;
        }
        set
        {
            driverstatus = value;
        }
    }

    public bool IsActive
    {
        get
        {
            return isactive;
        }
        set
        {
            isactive = value;
        }
    }

    public string OwesForUniforms
    {
        get
        {
            return owesforuniforms;
        }
        set
        {
            owesforuniforms = value;
        }
    }

    public string OwesForCashAdvance
    {
        get
        {
            return owesforcashadvance;
        }
        set
        {
            owesforcashadvance = value;
        }
    }

    public string EmployeeNotes
    {
        get
        {
            return employeenotes;
        }
        set
        {
            employeenotes = value;
        }
    }

    public int EmployeeId
    {
        get
        {
            return employeeid;
        }
        set
        {
            employeeid = value;
        }
    }

    public string EmployeeUserName
    {
        get
        {
            return employeeusername;
        }
        set
        {
            employeeusername = value;
        }
    }

    public string GetHireDate
    {
        get
        {
            return gethiredate;
        }
        set
        {
            gethiredate = value;
        }
    }

    public string GetInsuranceDate
    {
        get
        {
            return getinsurancedate;
        }
        set
        {
            getinsurancedate = value;
        }
    }

    public string TeamName
    {
        get
        {
            return teamname;
        }
        set
        {
            teamname = value;
        }
    }

    public int Rank
    {
        get
        {
            return rank;
        }
        set
        {
            rank = value;
        }
    }

    public int ParentID
    {
        get
        {
            return parentid;
        }
        set
        {
            parentid = value;
        }
    }

    public int EmployeeLoginId
    {
        get
        {
            return employeeloginid;
        }
        set
        {
            employeeloginid = value;
        }
    }

    public int SocialSecurityNumber
    {
        get
        {
            return socialsecuritynumber;
        }
        set
        {
            socialsecuritynumber = value;
        }
    }

    #endregion

    #region Methods

    public static DataTable BindDropDownList(string tableName, string whereCondition, DropDownList ddlName)
    {
        DataTable dt = null;
        using (SqlConnection con = new SqlConnection(Convert.ToString(ConfigurationManager.AppSettings["DBConnectionstring"])))
        {

            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    dt = SqlHelper.ExecuteDataset(con, CommandType.StoredProcedure, "SP_LoadDropDown", new SqlParameter("@TableName", tableName),
                                                                                                 new SqlParameter("@Where", whereCondition)).Tables[0];

                    ddlName.DataSource = dt;
                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
        }
        return dt;
    }

    public int InsertEmployee(int EmployeeID, string StoreId, string FirstName, string MiddleName, string LastName, string Address, string City, string State, string Zip,
            string HomePhone, string CellPhone, string CellPhoneCarrierId, string EmailAddress, int SystemRoleId, DateTime BirthDate, DateTime HireDate, 
            string MaritalStatus, int NumberOfDependents, string Insurance, DateTime InsuranceDate, string AdditionalWithHoldings, string DriverEligible, string DriverStatus,
            bool IsActive, int EmployeeloginID, int SocialSecurityNumber)
    {
        int i = 0;
        using (SqlConnection con = new SqlConnection(Convert.ToString(ConfigurationManager.AppSettings["DBConnectionstring"])))
        {
            SqlCommand cmd = new SqlCommand();

            i =Convert.ToInt32(SqlHelper.ExecuteScalar(con, CommandType.StoredProcedure, "SP_InsertEmployee", new SqlParameter("@EmployeeID", EmployeeID),
                                                                                                    new SqlParameter("@StoreID", StoreId),
                                                                                                    new SqlParameter("@FirstName", FirstName),
                                                                                                    new SqlParameter("@MiddleName", MiddleName),
                                                                                                    new SqlParameter("@LastName", LastName),
                                                                                                    new SqlParameter("@Address", Address),
                                                                                                    new SqlParameter("@City", City),
                                                                                                    new SqlParameter("@State", State),
                                                                                                    new SqlParameter("@Zip", Zip),
                                                                                                    new SqlParameter("@HomePhone", HomePhone),
                                                                                                    new SqlParameter("@CellPhone", CellPhone),
                                                                                                    new SqlParameter("@CellPhoneCarrierID", CellPhoneCarrierId),
                                                                                                    new SqlParameter("@EmailAddress", EmailAddress),
                                                                                                    new SqlParameter("@SystemRoleID", SystemRoleId),
                                                                                                    new SqlParameter("@BirthDate", BirthDate),
                                                                                                    new SqlParameter("@HireDate", HireDate),
                                                                                                    new SqlParameter("@MaritalStatus", MaritalStatus),
                                                                                                    new SqlParameter("@NumberOfDependents", NumberOfDependents),
                                                                                                    new SqlParameter("@Insurance", Insurance),
                                                                                                    new SqlParameter("@InsuranceDate", InsuranceDate),
                                                                                                    new SqlParameter("@AdditionalWithHoldings", AdditionalWithHoldings),
                                                                                                    new SqlParameter("@rbndriverEligible", DriverEligible),
                                                                                                    new SqlParameter("@rbnDriverStatus", DriverStatus),
                                                                                                    new SqlParameter("@IsActive", IsActive),
                                                                                                    new SqlParameter("@EmployeeLoginID", EmployeeloginID),
                                                                                                    new SqlParameter("@SocialSecurityNumber", SocialSecurityNumber)
                                                                                                    ));

        }
        return i;
    }

    public List<Employee> LoadEmployeeData(string EmployeeId)
    {
        DataTable dt = null;

        List<Employee> lstemployee = null;
        using (SqlConnection con = new SqlConnection(Convert.ToString(ConfigurationManager.AppSettings["DBConnectionstring"])))
        {

            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    dt = SqlHelper.ExecuteDataset(con, CommandType.StoredProcedure, "SP_LoadEmployee", new SqlParameter("@EmployeeId", EmployeeId)).Tables[0];
                    lstemployee = new List<Employee>();
                    Employee clsemployee = new Employee();
                    clsemployee.employeeusername = dt.Rows[0]["empid"].ToString();
                    clsemployee.FirstName = dt.Rows[0]["firstname"].ToString();
                    clsemployee.MiddleName = dt.Rows[0]["lastname"].ToString();
                    clsemployee.LastName = dt.Rows[0]["lastname"].ToString();
                    clsemployee.Address = dt.Rows[0]["address"].ToString();
                    clsemployee.City = dt.Rows[0]["city"].ToString();
                    clsemployee.State = dt.Rows[0]["StateName"].ToString();
                    clsemployee.Zip = dt.Rows[0]["Zip"].ToString();
                    clsemployee.HomePhone = dt.Rows[0]["homephone"].ToString();
                    clsemployee.CellPhone = dt.Rows[0]["cellphone"].ToString();
                    clsemployee.CellPhoneCarrierId = dt.Rows[0]["cellphonecarrierid"].ToString();
                    clsemployee.EmailAddress = dt.Rows[0]["emailaddress"].ToString();
                    clsemployee.SystemRoleId =Convert.ToInt32( dt.Rows[0]["systemroleid"]);
                    clsemployee.BirthDate = Convert.ToDateTime(dt.Rows[0]["birthdate"]);
                    if (string.IsNullOrEmpty(dt.Rows[0]["hiredate"].ToString()))
                        clsemployee.GetHireDate = null;
                    else
                    {
                        clsemployee.HireDate = Convert.ToDateTime(dt.Rows[0]["hiredate"]);
                        clsemployee.GetHireDate = clsemployee.HireDate.ToString();
                    }
                    clsemployee.MaritalStatus = dt.Rows[0]["maritalstatus"].ToString();
                    if (string.IsNullOrEmpty(dt.Rows[0]["numberofdependents"].ToString()))
                        clsemployee.NumberOfDependents = 0;
                    else
                        clsemployee.NumberOfDependents = Convert.ToInt32(dt.Rows[0]["numberofdependents"]);
                    clsemployee.Insurance = dt.Rows[0]["insurance"].ToString();
                    if (string.IsNullOrEmpty(dt.Rows[0]["insurancedate"].ToString()))
                        clsemployee.GetInsuranceDate = null;
                    else
                    {
                        clsemployee.InsuranceDate = Convert.ToDateTime(dt.Rows[0]["insurancedate"]);
                        clsemployee.GetHireDate = clsemployee.GetInsuranceDate.ToString();
                    }
                    clsemployee.AdditionalWithHoldings = dt.Rows[0]["additionalwitholdings"].ToString();
                    clsemployee.DriverEligible = dt.Rows[0]["drivereligible"].ToString();
                    clsemployee.DriverStatus = dt.Rows[0]["driverstatus"].ToString();
                    clsemployee.IsActive =Convert.ToBoolean(dt.Rows[0]["isactive"]);
                    clsemployee.StoreId = dt.Rows[0]["storeid"].ToString();
                    clsemployee.SocialSecurityNumber = (!string.IsNullOrEmpty(dt.Rows[0]["SocialSecurityNumber"].ToString())) ?  Convert.ToInt32(dt.Rows[0]["SocialSecurityNumber"]): 0;
                    clsemployee.EmployeeLoginId = (!string.IsNullOrEmpty(dt.Rows[0]["EmployeeID"].ToString())) ? Convert.ToInt32(dt.Rows[0]["EmployeeID"]) : 0;
                    lstemployee.Add(clsemployee);
                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
        }
        return lstemployee;
    }


    public int UpdateEmployee(int EmployeeID, string StoreId, string FirstName, string MiddleName, string LastName, string Address, string City, string State, string Zip,
            string HomePhone, string CellPhone, string CellPhoneCarrierId, string EmailAddress, int SystemRoleId, DateTime BirthDate, DateTime HireDate,
            string MaritalStatus, int NumberOfDependents, string Insurance, DateTime InsuranceDate, string AdditionalWithHoldings, string DriverEligible, string DriverStatus,
            bool IsActive)
    {
        int i = 0;
        using (SqlConnection con = new SqlConnection(Convert.ToString(ConfigurationManager.AppSettings["DBConnectionstring"])))
        {
            SqlCommand cmd = new SqlCommand();

            i = Convert.ToInt32(SqlHelper.ExecuteScalar(con, CommandType.StoredProcedure, "SP_InsertEmployee", new SqlParameter("@EmployeeID", EmployeeID),
                                                                                                    new SqlParameter("@StoreID", StoreId),
                                                                                                    new SqlParameter("@FirstName", FirstName),
                                                                                                    new SqlParameter("@MiddleName", MiddleName),
                                                                                                    new SqlParameter("@LastName", LastName),
                                                                                                    new SqlParameter("@Address", Address),
                                                                                                    new SqlParameter("@City", City),
                                                                                                    new SqlParameter("@State", State),
                                                                                                    new SqlParameter("@Zip", Zip),
                                                                                                    new SqlParameter("@HomePhone", HomePhone),
                                                                                                    new SqlParameter("@CellPhone", CellPhone),
                                                                                                    new SqlParameter("@CellPhoneCarrierID", CellPhoneCarrierId),
                                                                                                    new SqlParameter("@EmailAddress", EmailAddress),
                                                                                                    new SqlParameter("@SystemRoleID", SystemRoleId),
                                                                                                    new SqlParameter("@BirthDate", BirthDate),
                                                                                                    new SqlParameter("@HireDate", HireDate),
                                                                                                    new SqlParameter("@MaritalStatus", MaritalStatus),
                                                                                                    new SqlParameter("@NumberOfDependents", NumberOfDependents),
                                                                                                    new SqlParameter("@Insurance", Insurance),
                                                                                                    new SqlParameter("@InsuranceDate", InsuranceDate),
                                                                                                    new SqlParameter("@AdditionalWithHoldings", AdditionalWithHoldings),
                                                                                                    new SqlParameter("@rbndriverEligible", DriverEligible),
                                                                                                    new SqlParameter("@rbnDriverStatus", DriverStatus),
                                                                                                    new SqlParameter("@IsActive", IsActive)
                                                                                                    ));

        }
        return i;
    }

    public DataTable Search_VitoExistUser(string storeid, string systemroleid)
    {
        DataTable dt = null;
        using (SqlConnection con = new SqlConnection(Convert.ToString(ConfigurationManager.AppSettings["DBConnectionstring"])))
        {

            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    dt = SqlHelper.ExecuteDataset(con, CommandType.StoredProcedure, "SP_Search_VitoExistUser", new SqlParameter("@StoreID", storeid),
                                                                                                                new SqlParameter("@SystemRoleID", systemroleid)).Tables[0];
                    
                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
        }
        return dt;
    }

    public int UpdateLitmosAccountFlag(string EmployeeID)
    {
        int i = 0;
        using (SqlConnection con = new SqlConnection(Convert.ToString(ConfigurationManager.AppSettings["DBConnectionstring"])))
        {
            SqlCommand cmd = new SqlCommand();

            i = Convert.ToInt32(SqlHelper.ExecuteScalar(con, CommandType.StoredProcedure, "SP_UpdateLitmosAccountFlag", new SqlParameter("@EmployeeID", EmployeeID)));
        }
        return i;
    }

    public DataTable GetLitmosTeamName(int storeID)
    {
        DataTable dt = null;
        using (SqlConnection con = new SqlConnection(Convert.ToString(ConfigurationManager.AppSettings["DBConnectionstring"])))
        {

            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    dt = SqlHelper.ExecuteDataset(con, CommandType.StoredProcedure, "SP_GetTeam", new SqlParameter("@StoreID", storeID)).Tables[0];

                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
        }
        return dt;
    }

    public DataTable GetEmployeeRegsitrationDate(int EmployeeID)
    {
        DataTable dt = null;
        using (SqlConnection con = new SqlConnection(Convert.ToString(ConfigurationManager.AppSettings["DBConnectionstring"])))
        {

            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    dt = SqlHelper.ExecuteDataset(con, CommandType.StoredProcedure, "SP_GetEmployeeRegistrationDate", new SqlParameter("@EmployeeID", EmployeeID)).Tables[0];

                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
        }
        return dt;
    }

    public int InsertTeam()
    {
        int i = 0;
        using (SqlConnection con = new SqlConnection(Convert.ToString(ConfigurationManager.AppSettings["DBConnectionstring"])))
        {
            SqlCommand cmd = new SqlCommand();

            i = Convert.ToInt32(SqlHelper.ExecuteScalar(con, CommandType.StoredProcedure, "SP_InsertTeam", new SqlParameter("@TeamName", teamname),
                                                                                                    new SqlParameter("@Rank", rank),
                                                                                                    new SqlParameter("@ParentID", parentid)));

        }
        return i;
    }

    public DataTable GetEmployeeDueCourseDetails(int _employeeid)
    {
        DataTable dt = new DataTable();
        using (SqlConnection con = new SqlConnection(Convert.ToString(ConfigurationManager.AppSettings["DBConnectionstring"])))
        {

            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    dt = SqlHelper.ExecuteDataset(con, CommandType.StoredProcedure, "sp_posSelectDueCourseDetails", new SqlParameter("@EmpID", _employeeid)).Tables[0];

                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
        }
        return dt;
    }

    public DataTable GetManagerOverrideDueAlertDate(int empID)
    {
        DataTable dt = null;
        using (SqlConnection con = new SqlConnection(Convert.ToString(ConfigurationManager.AppSettings["DBConnectionstring"])))
        {

            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    dt = SqlHelper.ExecuteDataset(con, CommandType.StoredProcedure, "sp_posGetManagerOverrideDueAlertDate", new SqlParameter("@EmpID", empID)).Tables[0];

                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
        }
        return dt;
    }

    public int FindUserRole(int empID, int passcode)
    {
        int i = 0;
        using (SqlConnection con = new SqlConnection(Convert.ToString(ConfigurationManager.AppSettings["DBConnectionstring"])))
        {

            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    i = Convert.ToInt32(SqlHelper.ExecuteScalar(con, CommandType.StoredProcedure, "sp_findUserRole",
                                                                                                    new SqlParameter("@userid", empID),
                                                                                                    new SqlParameter("@passcode", passcode)));

                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
        }
        return i;
    }

    public int ManagerOverrideUserDueCourse(int empID, DateTime OverrideDate, int option)
    {
        int i = 0;
        using (SqlConnection con = new SqlConnection(Convert.ToString(ConfigurationManager.AppSettings["DBConnectionstring"])))
        {

            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    i = Convert.ToInt32(SqlHelper.ExecuteScalar(con, CommandType.StoredProcedure, "sp_posManagerOverrideUserDueCourse",
                                                                                                    new SqlParameter("@EmpID", empID),
                                                                                                    new SqlParameter("@OverrideDate", OverrideDate),
                                                                                                    new SqlParameter("@Option", option)));

                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
        }
        return i;
    }

    public S_MethodResult SendLitmosNotificationToManagerByStoreID(int storeID)
    {
        S_MethodResult sMethodResult = new S_MethodResult(false, -1, "", null);

        try
        {
            sMethodResult.ReturnValue = SqlHelper.ExecuteDataset(
                //DatabaseConnection.Connection,
                            DatabaseConnection.ConnectionString,
                            CommandType.StoredProcedure,
                            "sp_mgmtSendLitmosNotificationToManagerByStoreID",
                            new SqlParameter("storeID", storeID)
                        ).Tables[0];

            sMethodResult.ReturnCode = 1;
            sMethodResult.Message = "";
            sMethodResult.Success = true;
        }
        catch (Exception e)
        {
            sMethodResult.ReturnCode = -1;
            sMethodResult.Message = e.Message;
            sMethodResult.Success = false;
        }
        return sMethodResult;
    }

    public static DataTable GetLitmosAPIStatus()
    {
        DataTable dt = new DataTable();
        using (SqlConnection con = new SqlConnection(Convert.ToString(ConfigurationManager.AppSettings["DBConnectionstring"])))
        {

            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    dt = SqlHelper.ExecuteDataset(con, CommandType.StoredProcedure, "sp_posGetLitmosAPIStatus").Tables[0];

                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
        }
        return dt;
    }
    
    #endregion

}