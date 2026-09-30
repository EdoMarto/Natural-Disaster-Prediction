using OfficeOpenXml;
using System.Data.SqlClient;

namespace Thesis
{
    public class File7
    {
        public static void run()
        {
            Console.WriteLine("Sto avviando File7...");
            int n_inseriti = manageExcel("Indicatori_Intero_territorio_nazionale.xlsx");
            Console.WriteLine("Esecuzione File7 terminata. Sono stati inseriti " + n_inseriti + " nuovi record");
        }

        static void insertDB(object tmp1, object tmp2, object tmp3, object tmp4, object tmp5, object tmp6, object tmp7, object tmp8, object tmp9, object tmp10, object tmp11)
        {
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            string query = "INSERT INTO tbl_13029_Stage_INGV (AgMax_50, DataInserimento, DataUltimaModifica, Deleted, AgMin_50, ZNVULC_G, ZNVULC_G_TXT, ZNVULC_R, ZNVULC_R_TXT, ZNVULC_MVPU, ZNVULC_SOM, ZNVULC_SOM_TXT, ZNVULC_ALTRI, ZNVULC_ALTRI_TXT, DataImportazione, ParentWeb_Id) " +
                "VALUES(@AgMax_50, @DataInserimento, @DataUltimaModifica, @Deleted, @AgMin_50, @ZNVULC_G, @ZNVULC_G_TXT, @ZNVULC_R, @ZNVULC_R_TXT, @ZNVULC_MVPU, @ZNVULC_SOM, @ZNVULC_SOM_TXT, @ZNVULC_ALTRI, @ZNVULC_ALTRI_TXT, @DataImportazione, @ParentWeb_Id)";
            SqlCommand command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@AgMax_50", tmp1);
            command.Parameters.AddWithValue("@DataInserimento", DateTime.Now);
            command.Parameters.AddWithValue("@DataUltimaModifica", DateTime.Now);
            command.Parameters.AddWithValue("@Deleted", false);
            command.Parameters.AddWithValue("@AgMin_50", tmp2);
            command.Parameters.AddWithValue("@ZNVULC_G", tmp3);
            command.Parameters.AddWithValue("@ZNVULC_G_TXT", tmp4);
            command.Parameters.AddWithValue("@ZNVULC_R", tmp5);
            command.Parameters.AddWithValue("@ZNVULC_R_TXT", tmp6);
            command.Parameters.AddWithValue("@ZNVULC_MVPU", tmp7);
            command.Parameters.AddWithValue("@ZNVULC_SOM", tmp8);
            command.Parameters.AddWithValue("@ZNVULC_SOM_TXT", tmp9);
            command.Parameters.AddWithValue("@ZNVULC_ALTRI", tmp10);
            command.Parameters.AddWithValue("@ZNVULC_ALTRI_TXT", tmp11);
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

        static bool alreadyExists(object t1, object t2, object t3, object t4, object t5, object t6, object t7, object t8, object t9, object t10, object t11)
        {
            int count = 0;
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            SqlCommand command = new SqlCommand("SELECT COUNT(*) FROM tbl_13029_Stage_INGV WHERE AgMax_50=@t1 AND AgMin_50=@t2 AND ZNVULC_G=@t3 AND ZNVULC_G_TXT=@t4 AND ZNVULC_R=@t5 AND ZNVULC_R_TXT=@t6 AND ZNVULC_MVPU=@t7 AND ZNVULC_SOM=@t8 AND ZNVULC_SOM_TXT=@t9 AND ZNVULC_ALTRI=@t10 AND ZNVULC_ALTRI_TXT=@t11", connection);
            command.Parameters.AddWithValue("@t1", t1);
            command.Parameters.AddWithValue("@t2", t2);
            command.Parameters.AddWithValue("@t3", t3);
            command.Parameters.AddWithValue("@t4", t4);
            command.Parameters.AddWithValue("@t5", t5);
            command.Parameters.AddWithValue("@t6", t6);
            command.Parameters.AddWithValue("@t7", t2);
            command.Parameters.AddWithValue("@t8", t3);
            command.Parameters.AddWithValue("@t9", t4);
            command.Parameters.AddWithValue("@t10", t5);
            command.Parameters.AddWithValue("@t11", t6);

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
                    object s7 = null;
                    object s8 = null;
                    object s9 = null;
                    object s10 = null;
                    object s11 = null;

                    if (!workSheet.Cells[row, 80].Text.Equals("--"))
                    {
                        s1 = Double.Parse(workSheet.Cells[row, 80].Text);
                        s2 = Double.Parse(workSheet.Cells[row, 81].Text);
                        s3 = !string.IsNullOrEmpty(workSheet.Cells[row, 134].Text) && (workSheet.Cells[row, 134].Text[0] == '1' || workSheet.Cells[row, 134].Text[0] == 's');
                        s4 = !string.IsNullOrEmpty(workSheet.Cells[row, 135].Text) && (workSheet.Cells[row, 135].Text[0] == '1' || workSheet.Cells[row, 135].Text[0] == 's');
                        s5 = !string.IsNullOrEmpty(workSheet.Cells[row, 136].Text) && (workSheet.Cells[row, 136].Text[0] == '1' || workSheet.Cells[row, 136].Text[0] == 's');
                        s6 = !string.IsNullOrEmpty(workSheet.Cells[row, 137].Text) && (workSheet.Cells[row, 137].Text[0] == '1' || workSheet.Cells[row, 137].Text[0] == 's');
                        s7 = !string.IsNullOrEmpty(workSheet.Cells[row, 138].Text) && (workSheet.Cells[row, 138].Text[0] == '1' || workSheet.Cells[row, 138].Text[0] == 's');
                        s8 = !string.IsNullOrEmpty(workSheet.Cells[row, 139].Text) && (workSheet.Cells[row, 139].Text[0] == '1' || workSheet.Cells[row, 139].Text[0] == 's');
                        s9 = !string.IsNullOrEmpty(workSheet.Cells[row, 140].Text) && (workSheet.Cells[row, 140].Text[0] == '1' || workSheet.Cells[row, 140].Text[0] == 's');
                        s10 = !string.IsNullOrEmpty(workSheet.Cells[row, 141].Text) && (workSheet.Cells[row, 141].Text[0] == '1' || workSheet.Cells[row, 141].Text[0] == 's');
                        s11 = !string.IsNullOrEmpty(workSheet.Cells[row, 142].Text) && (workSheet.Cells[row, 142].Text[0] == '1' || workSheet.Cells[row, 142].Text[0] == 's');

                        if (!alreadyExists(s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11))
                        {
                            insertDB(s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11);
                            Console.WriteLine("Record inserito riga " + row);
                            n_inseriti++;
                        }
                    }
                }
            }
            return n_inseriti;
        }
    }
}
