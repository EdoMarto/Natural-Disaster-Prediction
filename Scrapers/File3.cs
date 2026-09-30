using System.Data.SqlClient;

namespace Thesis
{
    public class File3
    {
        public static void run()
        {
            Console.WriteLine("Sto avviando File3...");
            int n_inseriti = manageExcel("Indicatori incendi foreste.csv");
            Console.WriteLine("Esecuzione File3 terminata. Sono stati inseriti " + n_inseriti + " nuovi record");
        }

        static void insertDB(object tmp1, object tmp2, object tmp3)
        {
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            string query = "INSERT INTO tbl_13029_Stage_CorpoForestale_Incendi_Foreste (DataInserimento, DataUltimaModifica, Deleted, DataImportazione, Anno, Regione, PERC_PercentualeForestalePercorsaDalFuoco, ParentWeb_Id) " +
                "VALUES(@DataInserimento, @DataUltimaModifica, @Deleted, @DataImportazione, @Anno, @Regione, @PERC_PercentualeForestalePercorsaDalFuoco, @ParentWeb_Id)";
            SqlCommand command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@DataInserimento", DateTime.Now);
            command.Parameters.AddWithValue("@DataUltimaModifica", DateTime.Now);
            command.Parameters.AddWithValue("@Deleted", false);
            command.Parameters.AddWithValue("@DataImportazione", DateTime.Now);
            command.Parameters.AddWithValue("@Anno", tmp1 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Regione", tmp2 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_PercentualeForestalePercorsaDalFuoco", tmp3 ?? DBNull.Value);
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

        static bool alreadyExists(object t1, object t2, object t3)
        {
            int count = 0;
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            SqlCommand command = new SqlCommand("SELECT COUNT(*) FROM tbl_13029_Stage_CorpoForestale_Incendi_Foreste WHERE Anno=@t1 AND Regione=@t2 AND PERC_PercentualeForestalePercorsaDalFuoco=@t3", connection);
            command.Parameters.AddWithValue("@t1", t1);
            command.Parameters.AddWithValue("@t2", t2);
            command.Parameters.AddWithValue("@t3", t3);

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
            using (var reader = new StreamReader("../../../Excels/Manuali/" + filename))
            {
                var line = reader.ReadLine();
                while (!reader.EndOfStream)
                {
                    line = reader.ReadLine();
                    var values = line.Split(';');

                    object s1 = values[0];
                    object s2 = values[1];
                    object s3 = Double.Parse(values[2].Replace(".", ","));

                    if (!alreadyExists(s1, s2, s3))
                    {
                        insertDB(s1, s2, s3);
                        Console.WriteLine("Record inserito");
                        n_inseriti++;
                    }
                }
            }
            return n_inseriti;
        }
    }
}
