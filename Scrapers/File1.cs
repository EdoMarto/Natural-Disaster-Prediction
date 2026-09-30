using OfficeOpenXml;
using System.Data.SqlClient;

namespace Thesis
{
    public class File1
    {

        public static void run()
        {
            Console.WriteLine("Sto avviando File1...");
            int n_inseriti = manageExcel("Indicatori_Intero_territorio_nazionale.xlsx");
            Console.WriteLine("Esecuzione File1 terminata. Sono stati inseriti " + n_inseriti + " nuovi record");
        }

        static void insertDB(object tmp1, object tmp2, object tmp3, object tmp4)
        {
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            string query = "INSERT INTO tbl_13029_Stage_Agenzia_per_la_Coesione_Territoriale (Classe_comune, DataInserimento, DataUltimaModifica, Deleted, Macro_classe, Macro_classe_txt, Classe_comune_txt, DataImportazione, ParentWeb_Id) " +
                "VALUES(@Classe_comune, @DataInserimento, @DataUltimaModifica, @Deleted, @Macro_classe, @Macro_classe_txt, @Classe_comune_txt, @DataImportazione, @ParentWeb_Id)";
            SqlCommand command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@Classe_comune", tmp1);
            command.Parameters.AddWithValue("@DataInserimento", DateTime.Now);
            command.Parameters.AddWithValue("@DataUltimaModifica", DateTime.Now);
            command.Parameters.AddWithValue("@Deleted", false);
            command.Parameters.AddWithValue("@Macro_classe", tmp2);
            command.Parameters.AddWithValue("@Macro_classe_txt", tmp3);
            command.Parameters.AddWithValue("@Classe_comune_txt", tmp4);
            command.Parameters.AddWithValue("@DataImportazione", DateTime.Now);
            command.Parameters.AddWithValue("@ParentWeb_Id", 13029);

            try
            {
                connection.Open();
                command.ExecuteNonQuery();
            }
            catch (SqlException e)
            {
                Console.WriteLine("Errore Generato. Dettagli: " + e.ToString());
            }
            finally
            {
                connection.Close();
            }
        }

        static bool alreadyExists(object t1, object t2, object t3, object t4)
        {
            int count = 0;
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            SqlCommand command = new SqlCommand("SELECT COUNT(*) FROM tbl_13029_Stage_Agenzia_per_la_Coesione_Territoriale WHERE Classe_comune=@t1 AND Macro_classe=@t2 AND Macro_classe_txt=@t3 AND Classe_comune_txt=@t4", connection);
            command.Parameters.AddWithValue("@t1", t1);
            command.Parameters.AddWithValue("@t2", t2);
            command.Parameters.AddWithValue("@t3", t3);
            command.Parameters.AddWithValue("@t4", t4);

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
            using (var pck = new ExcelPackage(new FileInfo("../../../Excels/Manuali/" + filename)))
            {
                ExcelWorksheet workSheet = pck.Workbook.Worksheets["Intero territorio nazionale"];
                var start = new ExcelCellAddress(2, 1);
                var end = new ExcelCellAddress(7955, 397);
                for (int row = start.Row; row <= end.Row; row++)
                {
                    object s1 = null;
                    object s2 = null;
                    object s3 = null;
                    object s4 = null;

                    s1 = workSheet.Cells[row, 26].Text;
                    s2 = workSheet.Cells[row, 28].Text;
                    s3 = workSheet.Cells[row, 29].Text;
                    s4 = workSheet.Cells[row, 27].Text;

                    if (!alreadyExists(s1, s2, s3, s4))
                    {
                        insertDB(s1, s2, s3, s4);
                        Console.WriteLine("Record inserito riga " + row);
                        n_inseriti++;
                    }
                }
            }
            return n_inseriti;
        }
    }
}
