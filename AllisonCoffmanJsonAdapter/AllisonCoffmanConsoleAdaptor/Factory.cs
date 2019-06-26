using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Data;


namespace AllisonCoffmanConsoleAdaptor
{
    class Factory
    {
        public DataSet GetRetailerName(int id)
        {
            string RetailerName = string.Empty;
            try
            {
                string ConnectionString = Convert.ToString(ConfigurationManager.ConnectionStrings["DefaultConnection"]);
                SqlConnection conn = new SqlConnection(ConnectionString);
                //   string query = "spGetRetailerThirdPartyIntegrationDetailsDummy";
                string query = "spGetRetailerThirdPartyIntegrationDetails";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@thirdPartyIntegrationId", id);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

    }
}
