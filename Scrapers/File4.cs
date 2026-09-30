using System.Data.SqlClient;

namespace Thesis
{
    public class File4
    {
        public static void run()
        {
            Console.WriteLine("Sto avviando File4...");
            int n_inseriti = manageExcel("data-mS0XX.csv");
            Console.WriteLine("Esecuzione File4 terminata. Sono stati inseriti " + n_inseriti + " nuovi record");
        }

        static void insertDB(object tmp1, object tmp2, object tmp3, object tmp4, object tmp5, object tmp6, object tmp7, object tmp8, object tmp9)
        {
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            string query = "INSERT INTO tbl_13029_Stage_CorpoForestale_Numero_Incendi (Regione, DataInserimento, DataUltimaModifica, Deleted, DataImportazione, NumeroIncendi2009, NumeroIncendi2010, ParentWeb_Id, NumeroIncendi2011, NumeroIncendi2012, NumeroIncendi2013, NumeroIncendi2014, NumeroIncendi2015, NumeroIncendi2016) " +
                "VALUES(@Regione, @DataInserimento, @DataUltimaModifica, @Deleted, @DataImportazione, @NumeroIncendi2009, @NumeroIncendi2010, @ParentWeb_Id, @NumeroIncendi2011, @NumeroIncendi2012, @NumeroIncendi2013, @NumeroIncendi2014, @NumeroIncendi2015, @NumeroIncendi2016)";
            SqlCommand command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@Regione", tmp1);
            command.Parameters.AddWithValue("@DataInserimento", DateTime.Now);
            command.Parameters.AddWithValue("@DataUltimaModifica", DateTime.Now);
            command.Parameters.AddWithValue("@Deleted", false);
            command.Parameters.AddWithValue("@DataImportazione", DateTime.Now);
            command.Parameters.AddWithValue("@NumeroIncendi2009", tmp2);
            command.Parameters.AddWithValue("@NumeroIncendi2010", tmp3);
            command.Parameters.AddWithValue("@ParentWeb_Id", 13029);
            command.Parameters.AddWithValue("@NumeroIncendi2011", tmp4);
            command.Parameters.AddWithValue("@NumeroIncendi2012", tmp5);
            command.Parameters.AddWithValue("@NumeroIncendi2013", tmp6);
            command.Parameters.AddWithValue("@NumeroIncendi2014", tmp7);
            command.Parameters.AddWithValue("@NumeroIncendi2015", tmp8);
            command.Parameters.AddWithValue("@NumeroIncendi2016", tmp9);

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
            SqlCommand command = new SqlCommand("SELECT COUNT(*) FROM tbl_13029_Stage_CorpoForestale_Numero_Incendi WHERE Regione=@t1", connection);
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
            using (var reader = new StreamReader("../../../Excels/Manuali/" + filename))
            {
                var line = reader.ReadLine();
                while (!reader.EndOfStream)
                {
                    line = reader.ReadLine();
                    var values = line.Split(',');

                    object s1 = values[0];
                    object s2 = Double.Parse(values[1]);
                    object s3 = Double.Parse(values[2]);
                    object s4 = Double.Parse(values[3]);
                    object s5 = Double.Parse(values[4]);
                    object s6 = Double.Parse(values[5]);
                    object s7 = Double.Parse(values[6]);
                    object s8 = Double.Parse(values[7]);
                    object s9 = Double.Parse(values[8]);

                    if (!alreadyExists(s1))
                    {
                        insertDB(s1, s2, s3, s4, s5, s6, s7, s8, s9);
                        Console.WriteLine("Record inserito");
                        n_inseriti++;
                    }
                }
            }
            return n_inseriti;
        }
    }
}
