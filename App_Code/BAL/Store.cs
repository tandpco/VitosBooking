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
/// Summary description for Store
/// </summary>
public class Store
{
    private string storeId;
    private string storename;
    private string address;
    private string city;
    private string state;
    private string zip;
    private string workphone;
    private string cellphone;
    private string emailaddress;
    private string website;
    private string offset;
    private string networkip;
    private bool isactive;
    private string teamname;
    private int teamid;
    private int storeteamid;


	public Store()
	{
		//
		// TODO: Add constructor logic here
		//
	}

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

    public string StoreName
    {
        get
        {
            return storename;
        }
        set
        {
            storename = value;
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

    public string WorkPhone
    {
        get
        {
            return workphone;
        }
        set
        {
            workphone = value;
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

    public string Website
    {
        get
        {
            return website;
        }
        set
        {
            website = value;
        }
    }

    public string Offset
    {
        get
        {
            return offset;
        }
        set
        {
            offset = value;
        }
    }

    public string NetworkIP
    {
        get
        {
            return networkip;
        }
        set
        {
            networkip = value;
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

    public int TeamID
    {
        get
        {
            return teamid;
        }
        set
        {
            teamid = value;
        }
    }

    public int StoreTeamID
    {
        get
        {
            return storeteamid;
        }
        set
        {
            storeteamid = value;
        }
    }

    public int InsertStore()
    {
        int i = 0;
        using (SqlConnection con = new SqlConnection(Convert.ToString(ConfigurationManager.AppSettings["DBConnectionstring"])))
        {
            SqlCommand cmd = new SqlCommand();

            i = Convert.ToInt32(SqlHelper.ExecuteScalar(con, CommandType.StoredProcedure, "SP_InsertStore", new SqlParameter("@StoreName", storename),
                                                                                                    new SqlParameter("@Address", Address),
                                                                                                    new SqlParameter("@City", City),
                                                                                                    new SqlParameter("@State", State),
                                                                                                    new SqlParameter("@Zip", Zip),
                                                                                                    new SqlParameter("@WorkPhone", workphone),
                                                                                                    new SqlParameter("@CellPhone", cellphone),
                                                                                                    new SqlParameter("@EmailAddress", emailaddress),
                                                                                                    new SqlParameter("@Website", website),
                                                                                                    new SqlParameter("@Offset", offset),
                                                                                                    new SqlParameter("@NetworkIP", networkip)
                                                                                                    ));

        }
        return i;
    }

    public DataTable Search_VitoStores(string storeid)
    {
        DataTable dt = null;
        using (SqlConnection con = new SqlConnection(Convert.ToString(ConfigurationManager.AppSettings["DBConnectionstring"])))
        {

            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    dt = SqlHelper.ExecuteDataset(con, CommandType.StoredProcedure, "SP_Search_VitoStores", new SqlParameter("@StoreID", storeid)).Tables[0];

                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
        }
        return dt;
    }

    public DataTable BindUnAssignedTeamListToStoreTeam()
    {
        DataTable dt = null;
        using (SqlConnection con = new SqlConnection(Convert.ToString(ConfigurationManager.AppSettings["DBConnectionstring"])))
        {
            SqlCommand cmd = new SqlCommand();

            dt = SqlHelper.ExecuteDataset(con, CommandType.StoredProcedure, "SP_GetUnAssignedTeamListToStoreTeam", new SqlParameter("@StoreID", storeId)).Tables[2];

        }
        return dt;
    }

    public int InsertStoreTeam()
    {
        int i = 0;
        using (SqlConnection con = new SqlConnection(Convert.ToString(ConfigurationManager.AppSettings["DBConnectionstring"])))
        {
            SqlCommand cmd = new SqlCommand();

            i = Convert.ToInt32(SqlHelper.ExecuteScalar(con, CommandType.StoredProcedure, "SP_InsertStoreTeam", new SqlParameter("@TeamID", teamid),
                                                                                                    new SqlParameter("@StoreID", storeId)));

        }
        return i;
    }

    public DataTable GetStoreTeams(string storeid)
    {
        DataTable dt = null;
        using (SqlConnection con = new SqlConnection(Convert.ToString(ConfigurationManager.AppSettings["DBConnectionstring"])))
        {

            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    dt = SqlHelper.ExecuteDataset(con, CommandType.StoredProcedure, "SP_GetStoreTeams", new SqlParameter("@StoreID", storeid)).Tables[0];

                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
        }
        return dt;
    }

    public int DeleteStoreTeam()
    {
        int i = 0;
        using (SqlConnection con = new SqlConnection(Convert.ToString(ConfigurationManager.AppSettings["DBConnectionstring"])))
        {
            SqlCommand cmd = new SqlCommand();

            i = Convert.ToInt32(SqlHelper.ExecuteScalar(con, CommandType.StoredProcedure, "SP_DeleteStoreTeam", new SqlParameter("@ID", storeteamid)));

        }
        return i;
    }
}