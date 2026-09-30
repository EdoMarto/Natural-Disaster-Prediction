using OfficeOpenXml;
using System.Data.SqlClient;

namespace Thesis
{
    public class File6
    {
        public static void run()
        {
            Console.WriteLine("Sto avviando File6...");
            int n_inseriti = manageExcel("Indicatori_Intero_territorio_nazionale.xlsx");
            Console.WriteLine("Esecuzione File6 terminata. Sono stati inseriti " + n_inseriti + " nuovi record");
        }

        static void insertDB(object tmp1, object tmp2, object tmp3, object tmp4, object tmp5, object tmp6)
        {
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            string query = "INSERT INTO tbl_13029_Stage_Elaborazione_Istat_su_DEM_ISPRA (Alt_min, DataInserimento, DataUltimaModifica, Deleted, Alt_min_p, Alt_min_r, Alt_max, Alt_max_p, Alt_max_r, DataImportazione, ParentWeb_Id) " +
                "VALUES(@Alt_min, @DataInserimento, @DataUltimaModifica, @Deleted, @Alt_min_p, @Alt_min_r, @Alt_max, @Alt_max_p, @Alt_max_r, @DataImportazione, @ParentWeb_Id)";
            SqlCommand command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@Alt_min", tmp1);
            command.Parameters.AddWithValue("@DataInserimento", DateTime.Now);
            command.Parameters.AddWithValue("@DataUltimaModifica", DateTime.Now);
            command.Parameters.AddWithValue("@Deleted", false);
            command.Parameters.AddWithValue("@Alt_min_p", tmp2);
            command.Parameters.AddWithValue("@Alt_min_r", tmp3);
            command.Parameters.AddWithValue("@Alt_max", tmp4);
            command.Parameters.AddWithValue("@Alt_max_p", tmp5);
            command.Parameters.AddWithValue("@Alt_max_r", tmp6);
            command.Parameters.AddWithValue("@DataImportazione", DateTime.Now);
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

        static bool alreadyExists(object t1, object t2, object t3, object t4, object t5, object t6)
        {
            int count = 0;
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            SqlCommand command = new SqlCommand("SELECT COUNT(*) FROM tbl_13029_Stage_Elaborazione_Istat_su_DEM_ISPRA WHERE Alt_min=@t1 AND Alt_min_p=@t2 AND Alt_min_r=@t3 AND Alt_max=@t4 AND Alt_max_p=@t5 AND Alt_max_r=@t6", connection);
            command.Parameters.AddWithValue("@t1", t1);
            command.Parameters.AddWithValue("@t2", t2);
            command.Parameters.AddWithValue("@t3", t3);
            command.Parameters.AddWithValue("@t4", t4);
            command.Parameters.AddWithValue("@t5", t5);
            command.Parameters.AddWithValue("@t6", t6);

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
                    object s5 = null;
                    object s6 = null;

                    s1 = Double.Parse(workSheet.Cells[row, 20].Text);
                    s2 = Double.Parse(workSheet.Cells[row, 21].Text);
                    s3 = Double.Parse(workSheet.Cells[row, 22].Text);
                    s4 = Double.Parse(workSheet.Cells[row, 23].Text);
                    s5 = Double.Parse(workSheet.Cells[row, 24].Text);
                    s6 = Double.Parse(workSheet.Cells[row, 25].Text);

                    if (!alreadyExists(s1, s2, s3, s4, s5, s6))
                    {
                        insertDB(s1, s2, s3, s4, s5, s6);
                        Console.WriteLine("Record inserito riga " + row);
                        n_inseriti++;
                    }
                }
            }
            return n_inseriti;
        }
    }
}
