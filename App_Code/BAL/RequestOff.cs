using System;
using Microsoft.ApplicationBlocks.Data;
using System.Data;
using System.Data.SqlClient;
/// <summary>
/// Summary description for RequestOff
/// </summary>
public class RequestOff
{

    #region Public Properties

    public int EmployeeID
    { get; set; }

    public int StoreID
    { get; set; }

    public DateTime OffDate
    { get; set; }

    public string Reason
    { get; set; }

    public string DaySession
    { get; set; }

    #endregion

    #region Public Methods

    public int AddRequestOff()
    {
        int result = 0;

        try
        {
            SqlParameter ReturnValue = new SqlParameter("ReturnValue", SqlDbType.Int);
            ReturnValue.Direction = ParameterDirection.ReturnValue;

            SqlHelper.ExecuteNonQuery(
                            DatabaseConnection.Connection,
                            CommandType.StoredProcedure,
                            "sp_scdlAddRequestOff",
                            new SqlParameter[] { 
                                    new SqlParameter("EmployeeID", EmployeeID),
                                    new SqlParameter("StoreID", StoreID),
                                    new SqlParameter("OffDate", OffDate),
                                    new SqlParameter("Reason", Reason),
                                    new SqlParameter("Session", DaySession),
                                    ReturnValue
                                }
                        );

            result = Convert.ToInt32(ReturnValue.Value);
            
        }
        catch (Exception e)
        {
            
            throw e;
            
        }
        return result;
    }

    public DataTable GetEmployeeDetailByEmployeeID(int employeeID)
    {
        DataTable dt = new DataTable();

        try
        {
            dt = SqlHelper.ExecuteDataset(
                            DatabaseConnection.Connection,
                            CommandType.StoredProcedure,
                            "sp_scdlGetEmployeeDetailByEmployeeID",
                            new SqlParameter("EmployeeID", employeeID)
                        ).Tables[0];

            
        }
        catch (Exception e)
        {
            throw e;
        }
        return dt;
    }

    #endregion
}