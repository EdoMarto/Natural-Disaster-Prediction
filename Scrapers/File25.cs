using OfficeOpenXml;
using System.Data.SqlClient;

namespace Thesis
{
    public class File25
    {
        public static void run()
        {
            Console.WriteLine("Sto avviando File25...");
            int n_inseriti = manageExcel("Indicatori_Intero_territorio_nazionale.xlsx");
            Console.WriteLine("Esecuzione File25 terminata. Sono stati inseriti " + n_inseriti + " nuovi record");
        }

        static void insertDB(object tmp1, object tmp2, object tmp3)
        {
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            string query = "INSERT INTO tbl_13029_Stage_Min_Amb_Mare (PRES_EUAP, DataInserimento, DataUltimaModifica, Deleted, PRES_EUAP_TXT, PRES_NAT00, DataImportazione, ParentWeb_Id) " +
                "VALUES(@PRES_EUAP, @DataInserimento, @DataUltimaModifica, @Deleted, @PRES_EUAP_TXT, @PRES_NAT00, @DataImportazione, @ParentWeb_Id)";
            SqlCommand command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@PRES_EUAP", tmp1);
            command.Parameters.AddWithValue("@DataInserimento", DateTime.Now);
            command.Parameters.AddWithValue("@DataUltimaModifica", DateTime.Now);
            command.Parameters.AddWithValue("@Deleted", false);
            command.Parameters.AddWithValue("@PRES_EUAP_TXT", tmp2);
            command.Parameters.AddWithValue("@PRES_NAT00", tmp3);
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

        static bool alreadyExists(object t1, object t2, object t3)
        {
            int count = 0;
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            SqlCommand command = new SqlCommand("SELECT COUNT(*) FROM tbl_13029_Stage_Min_Amb_Mare WHERE PRES_EUAP=@t1 AND PRES_EUAP_TXT=@t2 AND PRES_NAT00=@t3", connection);
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

                    s1 = Double.Parse(workSheet.Cells[row, 30].Text);
                    s2 = workSheet.Cells[row, 31].Text;
                    s3 = Double.Parse(workSheet.Cells[row, 32].Text);

                    if (!alreadyExists(s1, s2, s3))
                    {
                        insertDB(s1, s2, s3);
                        Console.WriteLine("Record inserito riga " + row);
                        n_inseriti++;
                    }
                }
            }
            return n_inseriti;
        }
    }
}
