using System;
using System.Data.SqlClient;
using System.IO;

namespace InventoryAuditService
{
    public class Database
    {
        SqlConnection sqlConnection;

        public Database()
        {
            try
            {
                
                  sqlConnection = new SqlConnection(@"Data Source = IOPAWSEYWEMBLEY; Initial Catalog = itelAssets; Integrated Security = True");
                //sqlConnection = new SqlConnection(@"Server = localhost; Initial Catalog = itelAssets;Integrated Security=True");
                //sqlConnection = new SqlConnection(@"Server = 192.168.112.142; Database = itelAssetManagementMBJ; User Id = wembley.williams; Password =#Leonardo21;");                
            }
            catch (SqlException e)
            {
                WriteToFile("Sql error" + e.Message);
            }
        }

        public void InsertPC(PCData data)
        {
            try
            {
                sqlConnection.Open();
                SqlCommand sqlCommand = new SqlCommand("INSERT INTO PC(pcSerial,pcVendor,pcModel,pcWindowsVersion,pcName,domain,assetName) " +
                    "VALUES('" + data.serialNumberPC + "','" + data.vendorPC + "','" + data.modelPC + "','" + data.version + "','" + data.systemName + "','" + data.domain +
                    "','" + data.assetName + "')"
                    , sqlConnection);
                sqlCommand.ExecuteNonQuery();
            }
            catch (SqlException e)
            {
                WriteToFile("Sql error from insertPC method" + e.Message);
            }
            finally
            {
                sqlConnection.Close();
            }

        }

        public void InsertMonitor(MonitorData data)
        {
            try
            {
                sqlConnection.Open();
                SqlCommand sqlCommand = new SqlCommand("INSERT INTO Monitor(monitorModel, monitorSerial, monitorVendor, attachedPC) " +
                    "VALUES('" + data.modelM + "','" + data.serialNumberM + "','" + data.vendorM + "','" + data.attachedPC + "')"
                     , sqlConnection);
                sqlCommand.ExecuteNonQuery();
            }
            catch (SqlException e)
            {
                WriteToFile("Sql error from insertMonitor method" + e.Message);
            }
            finally
            {
                sqlConnection.Close();
            }
        }

        public bool UpdatePC(PCData data)
        {
            try
            {
                sqlConnection.Open();
                using (SqlCommand sqlCommand = sqlConnection.CreateCommand())
                {
                    sqlCommand.Parameters.AddWithValue("@pcs", data.serialNumberPC);
                    sqlCommand.Parameters.AddWithValue("@pcv", data.vendorPC);
                    sqlCommand.Parameters.AddWithValue("@pcm", data.modelPC);
                    sqlCommand.Parameters.AddWithValue("@pcwv", data.version);
                    sqlCommand.Parameters.AddWithValue("@pcn", data.systemName);
                    sqlCommand.Parameters.AddWithValue("@d", data.domain);
                    sqlCommand.Parameters.AddWithValue("@an", data.assetName);

                    sqlCommand.CommandText = "UPDATE PC " + " SET pcSerial=@pcs, pcVendor=@pcv, pcModel=@pcm, pcWindowsVersion=@pcwv," +
                    " pcName=@pcn, domain=@d, assetName=@an WHERE pcSerial='" + data.serialNumberPC + "'";

                    sqlCommand.ExecuteNonQuery();
                }

            }
            catch (SqlException e)
            {
                WriteToFile("Sql error from updateMonitor method" + e.Message);
                return false;
            }
            finally
            {
                sqlConnection.Close();
            }
            return true;
        }

        public bool verifyPC(string serial)
        {
            try
            {
                sqlConnection.Open();
                SqlCommand sqlCommand = new SqlCommand("SELECT * FROM PC WHERE pcSerial ='" + serial + "'", sqlConnection);
                if (sqlCommand.ExecuteReader().HasRows)
                    return true;
            }
            catch (SqlException e)
            {
                WriteToFile("Sql error" + e.Message);
            }
            finally
            {
                sqlConnection.Close();
            }
            return false;
        }

        public bool verifyMonitor(string serial)
        {
            try
            {
                sqlConnection.Open();
                SqlCommand sqlCommand = new SqlCommand("SELECT * FROM Monitor WHERE monitorSerial ='" + serial + "'", sqlConnection);
                if (sqlCommand.ExecuteReader().HasRows)
                    return true;

            }
            catch (SqlException e)
            {
                WriteToFile("Sql error" + e.Message);
            }
            finally
            {
                sqlConnection.Close();
            }

            return false;
        }

        private static void WriteToFile(string Message)
        {
            string path = @"C:\Audit.txt";
            try
            {
                if (!File.Exists(path))
                { //Create a file to write to
                    using (StreamWriter writer = File.CreateText(path))
                    {
                        File.SetAttributes(path, File.GetAttributes(path) | FileAttributes.Hidden);
                        writer.WriteLine(string.Format(Message + "  --   {0}", DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")));
                        writer.Close();
                    }
                }
                else
                {
                    using (StreamWriter writer = File.AppendText(path))
                    {
                        writer.WriteLine(string.Format(Message + "  --   {0}", DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")));
                        writer.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

    }
}
