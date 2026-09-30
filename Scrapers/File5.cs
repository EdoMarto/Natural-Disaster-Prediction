using System.Data.OleDb;
using System.Data.SqlClient;

namespace Thesis
{
    public class File5
    {
        public static void run()
        {
            Console.WriteLine("Sto avviando File5...");
            int n_inseriti = manageExcel("Indicatori rischio sismico 2012.xls");
            Console.WriteLine("Esecuzione File5 terminata. Sono stati inseriti " + n_inseriti + " nuovi record");
        }

        static void insertDB(object tmp1, object tmp2)
        {
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            string query = "INSERT INTO tbl_13029_Stage_Dipartimento_Protezione_Civile_Rischio_Sismico (DataImportazione, DataInserimento, DataUltimaModifica, Deleted, Anno2012, ParentWeb_Id, Comune) " +
                "VALUES(@DataImportazione, @DataInserimento, @DataUltimaModifica, @Deleted, @Anno2012, @ParentWeb_Id, @Comune)";
            SqlCommand command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@DataImportazione", DateTime.Now);
            command.Parameters.AddWithValue("@DataInserimento", DateTime.Now);
            command.Parameters.AddWithValue("@DataUltimaModifica", DateTime.Now);
            command.Parameters.AddWithValue("@Deleted", false);
            command.Parameters.AddWithValue("@Comune", tmp1 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Anno2012", tmp2 ?? DBNull.Value);
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

        static bool alreadyExists(object t1)
        {
            int count = 0;
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            SqlCommand command = new SqlCommand("SELECT COUNT(*) FROM tbl_13029_Stage_Dipartimento_Protezione_Civile_Rischio_Sismico WHERE Comune=@t1", connection);
            command.Parameters.AddWithValue("@t1", t1);

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

        static int manageExcel(String filename)
        {
            int n_inseriti = 0;
            string fullPathToExcel = "../../../Excels/Manuali/" + filename;
            string con = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=yes'", fullPathToExcel);
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from [Test$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        object s1 = null;
                        object s2 = null;

                        s1 = dr[1];
                        if (s1.ToString().Equals("") || s1.ToString().Contains("Fonte:") || s1.ToString().Contains("COMUNI") || s1.ToString().Contains("Note") || s1.ToString().Contains("_")) continue;
                        if (!dr[10].ToString().Equals("-")) s2 = dr[19];

                        if (!alreadyExists(s1))
                        {
                            insertDB(s1, s2);
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
