using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO.Compression;
using System.Net;

namespace Thesis
{
    public class File8
    {
        public static void run()
        {
            Console.WriteLine("Sto avviando File8...");
            int n_inseriti = manageExcel("http://zonesismiche.mi.ingv.it/elaborazioni/dati/italia_ag_002_xls.zip");
            Console.WriteLine("Esecuzione File8 terminata. Sono stati inseriti " + n_inseriti + " nuovi record");
        }

        static void insertDB(object tmp1, object tmp2, object tmp3, object tmp4, object tmp5, object tmp6)
        {
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            string query = "INSERT INTO tbl_13029_Stage_INGV_Accelerazione_Sismica (DataImportazione, DataInserimento, DataUltimaModifica, Deleted, Longitudine, Latitudine, Ag, Perc16, Perc84, IDDato, ParentWeb_Id) " +
                "VALUES(@DataImportazione, @DataInserimento, @DataUltimaModifica, @Deleted, @Longitudine, @Latitudine, @Ag, @Perc16, @Perc84, @IDDato, @ParentWeb_Id)";
            SqlCommand command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@DataImportazione", DateTime.Now);
            command.Parameters.AddWithValue("@DataInserimento", DateTime.Now);
            command.Parameters.AddWithValue("@DataUltimaModifica", DateTime.Now);
            command.Parameters.AddWithValue("@Deleted", false);
            command.Parameters.AddWithValue("@Longitudine", tmp2 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Latitudine", tmp3 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Ag", tmp4 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Perc16", tmp5 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Perc84", tmp6 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDDato", tmp1 ?? DBNull.Value);
            command.Parameters.AddWithValue("@ParentWeb_Id", 13029);

            try
            {
                connection.Open();
                command.ExecuteNonQuery();
            }
            catch (SqlException e)
            {
                Console.WriteLine("Error Generated. Details: " + e.ToString());
            }
            finally
            {
                connection.Close();
            }
        }

        static bool alreadyExists(object pid)
        {
            int count = 0;
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            SqlCommand command = new SqlCommand("SELECT COUNT(*) FROM tbl_13029_Stage_INGV_Accelerazione_Sismica WHERE IDDato=@pid", connection);
            command.Parameters.AddWithValue("@pid", pid);

            try
            {
                connection.Open();
                count = (int)command.ExecuteScalar();
            }
            catch (SqlException e)
            {
                Console.WriteLine(e.ToString());
            }
            finally
            {
                connection.Close();
            }
            return count != 0;
        }

        static int manageExcel(String url)
        {
            int n_inseriti = 0;
            String filename = url.Substring(url.LastIndexOf("/") + 1);
            using (WebClient wc = new WebClient())
            {
                wc.DownloadFile(
                    new System.Uri(url),
                    "../../../Excels/Automatici/" + filename
                );
            }

            using (ZipArchive zip = ZipFile.Open("../../../Excels/Automatici/" + filename, ZipArchiveMode.Read))
                foreach (ZipArchiveEntry entry in zip.Entries)
                    if (entry.Name == "italia_ag_002_xls.xls")
                        entry.ExtractToFile("../../../Excels/Automatici/italia_ag_002_xls.xls", true);

            string fullPathToExcel = "../../../Excels/Automatici/italia_ag_002_xls.xls";
            string con = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=yes'", fullPathToExcel);
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from [Italia_1_di_2$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        object s1 = null;
                        object s2 = null;
                        object s3 = null;
                        object s4 = null;
                        object s5 = null;
                        object s6 = null;

                        s1 = dr[0];
                        if (s1.ToString().Contains("ID")) continue;
                        if (!dr[1].ToString().Equals("..")) s2 = dr[1];
                        if (!dr[2].ToString().Equals("..")) s3 = dr[2];
                        if (!dr[3].ToString().Equals("..")) s4 = dr[3];
                        if (!dr[4].ToString().Equals("..")) s5 = dr[4];
                        if (!dr[5].ToString().Equals("..")) s6 = dr[5];

                        if (!alreadyExists(s1))
                        {
                            insertDB(s1, s2, s3, s4, s5, s6);
                            Console.WriteLine("Record inserito");
                            n_inseriti++;
                        }
                    }
                }

                OleDbCommand command2 = new OleDbCommand("select * from [Italia_2_di_2$]", connection);
                using (OleDbDataReader dr = command2.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        object s1 = null;
                        object s2 = null;
                        object s3 = null;
                        object s4 = null;
                        object s5 = null;
                        object s6 = null;

                        s1 = dr[0];
                        if (s1.ToString().Contains("ID")) continue;
                        if (!dr[1].ToString().Equals("..")) s2 = dr[1];
                        if (!dr[2].ToString().Equals("..")) s3 = dr[2];
                        if (!dr[3].ToString().Equals("..")) s4 = dr[3];
                        if (!dr[4].ToString().Equals("..")) s5 = dr[4];
                        if (!dr[5].ToString().Equals("..")) s6 = dr[5];

                        if (!alreadyExists(s1))
                        {
                            insertDB(s1, s2, s3, s4, s5, s6);
                            Console.WriteLine("Record inserito");
                            n_inseriti++;
                        }
                    }
                }
            }
            return n_inseriti;
        }
    }
}
