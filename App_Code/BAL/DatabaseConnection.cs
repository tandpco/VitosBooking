using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using System.Configuration;

/// <summary>
/// Summary description for DatabaseConnection
/// </summary>
public class DatabaseConnection
{
	public DatabaseConnection()
	{
		//
		// TODO: Add constructor logic here
		//
	}
    
    public static SqlConnection Connection
    {
        get { return new SqlConnection(ConfigurationManager.AppSettings["DBConnectionstring"]); }
    }

    public static string ConnectionString
    {
        get { return ConfigurationManager.AppSettings["DBConnectionstring"]; }
    }
}