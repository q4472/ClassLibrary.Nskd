using System;
using System.Data;
using System.Data.SqlClient;

namespace Nskd
{
    public class SqlServer
    {
        public static void LogWrite(String msg, String source, String type, String sqlServerDataSource)
        {
            if (!String.IsNullOrWhiteSpace(sqlServerDataSource))
            {
                String sCnString = String.Format("Data Source={0};Integrated Security=True", sqlServerDataSource);
                SqlCommand cmd = new SqlCommand
                {
                    Connection = new SqlConnection(sCnString),
                    CommandType = CommandType.StoredProcedure,
                    CommandText = "[phs_s].[dbo].[data_server_log_insert]"
                };
                cmd.Parameters.AddWithValue("@message", msg);
                cmd.Parameters.AddWithValue("@source", source);
                cmd.Parameters.AddWithValue("@type", type);
                using (cmd.Connection)
                {
                    try
                    {
                        cmd.Connection.Open();
                        cmd.ExecuteNonQuery();
                    }
                    catch (SqlException e) { Console.WriteLine(e.Message); }
                }
            }
        }
    }
}

