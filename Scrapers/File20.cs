using System.Data.OleDb;
using System.Data.SqlClient;

namespace Thesis
{
    public class File20
    {
        public static void run()
        {
            Console.WriteLine("Sto avviando File20...");
            int n_inseriti = manageExcel("Indicatori idrogeologici.xlsx");
            Console.WriteLine("Esecuzione File20 terminata. Sono stati inseriti " + n_inseriti + " nuovi record");
        }

        static void insertDB(object tmp1, object tmp2, object tmp3, object tmp4, object tmp5, object tmp6, object tmp7, object tmp8, object tmp9)
        {
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            string query = "INSERT INTO tbl_13029_Stage_Istat_Idrogeologici (DataImportazione, DataInserimento, DataUltimaModifica, Deleted, AreePericolositaIdraulicaBassa, AreePericolositaIdraulicaMedia, AreePericolositaIdraulicaElevata, AreediAttenzionePAI_AA, AreePericolositaFranaPAIModerata_p1, AreePericolositaFranaPAIMedia_p2, AreePericolositaFranaPAIElevata_p3, AreePericolositaFranaPAIMoltoElevata_p4, ParentWeb_Id, Comune) " +
                "VALUES(@DataImportazione, @DataInserimento, @DataUltimaModifica, @Deleted, @AreePericolositaIdraulicaBassa, @AreePericolositaIdraulicaMedia, @AreePericolositaIdraulicaElevata, @AreediAttenzionePAI_AA, @AreePericolositaFranaPAIModerata_p1, @AreePericolositaFranaPAIMedia_p2, @AreePericolositaFranaPAIElevata_p3, @AreePericolositaFranaPAIMoltoElevata_p4, @ParentWeb_Id, @Comune)";
            SqlCommand command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@DataImportazione", DateTime.Now);
            command.Parameters.AddWithValue("@DataInserimento", DateTime.Now);
            command.Parameters.AddWithValue("@DataUltimaModifica", DateTime.Now);
            command.Parameters.AddWithValue("@Deleted", false);
            command.Parameters.AddWithValue("@AreePericolositaIdraulicaBassa", tmp2);
            command.Parameters.AddWithValue("@AreePericolositaIdraulicaMedia", tmp3);
            command.Parameters.AddWithValue("@AreePericolositaIdraulicaElevata", tmp4);
            command.Parameters.AddWithValue("@AreediAttenzionePAI_AA", tmp5);
            command.Parameters.AddWithValue("@AreePericolositaFranaPAIModerata_p1", tmp6);
            command.Parameters.AddWithValue("@AreePericolositaFranaPAIMedia_p2", tmp7);
            command.Parameters.AddWithValue("@AreePericolositaFranaPAIElevata_p3", tmp8);
            command.Parameters.AddWithValue("@AreePericolositaFranaPAIMoltoElevata_p4", tmp9);
            command.Parameters.AddWithValue("@ParentWeb_Id", 13029);
            command.Parameters.AddWithValue("@Comune", tmp1);

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

        static bool alreadyExists(object comune)
        {
            int count = 0;
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            SqlCommand command = new SqlCommand("SELECT COUNT(*) FROM tbl_13029_Stage_Istat_Idrogeologici WHERE Comune=@comune", connection);
            command.Parameters.AddWithValue("@comune", (string)comune);

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
                        object s3 = null;
                        object s4 = null;
                        object s5 = null;
                        object s6 = null;
                        object s7 = null;
                        object s8 = null;
                        object s9 = null;

                        s1 = dr[0];
                        if (s1.ToString().Equals("") || s1.ToString().Contains("Dataset:") || s1.ToString().Contains("Seleziona") || s1.ToString().Contains("Tipo") || s1.ToString().Contains("Territorio") || s1.ToString().Contains("Dati") || s1.ToString().Contains("Legend") || s1.ToString().Contains("u:")) continue;
                        if (!dr[2].ToString().Equals("....")) s2 = dr[2];
                        if (!dr[3].ToString().Equals("....")) s3 = dr[3];
                        if (!dr[4].ToString().Equals("....")) s4 = dr[4];
                        if (!dr[5].ToString().Equals("....")) s5 = dr[5];
                        if (!dr[6].ToString().Equals("....")) s6 = dr[6];
                        if (!dr[7].ToString().Equals("....")) s7 = dr[7];
                        if (!dr[8].ToString().Equals("....")) s8 = dr[8];
                        if (!dr[9].ToString().Equals("....")) s9 = dr[9];

                        if (!alreadyExists(s1))
                        {
                            insertDB(s1, s2, s3, s4, s5, s6, s7, s8, s9);
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
