using OfficeOpenXml;
using System.Data.SqlClient;

namespace Thesis
{
    public class File24
    {
        public static void run()
        {
            Console.WriteLine("Sto avviando File24...");
            int n_inseriti = manageExcel("Indicatori_Intero_territorio_nazionale.xlsx");
            Console.WriteLine("Esecuzione File24 terminata. Sono stati inseriti " + n_inseriti + " nuovi record");
        }

        static void insertDB(object tmp1, object tmp2, object tmp3, object tmp4, object tmp5, object tmp6, object tmp7, object tmp8, object tmp9, object tmp10, object tmp11, object tmp12)
        {
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            string query = "INSERT INTO tbl_13029_Stage_MIBACT (TOT_BENI, DataInserimento, DataUltimaModifica, Deleted, TOT_BENI_P, TOT_BENI_R, TOT_ARCHIT, TOT_ARCHIT_P, TOT_ARCHIT_R, TOT_ARCHEO, TOT_ARCHEO_P, TOT_ARCHEO_R, TOT_PARC_GIAR, TOT_PARC_GIAR_P, TOT_PARC_GIAR_R, DataImportazione, ParentWeb_Id) " +
                "VALUES(@TOT_BENI, @DataInserimento, @DataUltimaModifica, @Deleted, @TOT_BENI_P, @TOT_BENI_R, @TOT_ARCHIT, @TOT_ARCHIT_P, @TOT_ARCHIT_R, @TOT_ARCHEO, @TOT_ARCHEO_P, @TOT_ARCHEO_R, @TOT_PARC_GIAR, @TOT_PARC_GIAR_P, @TOT_PARC_GIAR_R, @DataImportazione, @ParentWeb_Id)";
            SqlCommand command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@TOT_BENI", tmp1 ?? DBNull.Value);
            command.Parameters.AddWithValue("@DataInserimento", DateTime.Now);
            command.Parameters.AddWithValue("@DataUltimaModifica", DateTime.Now);
            command.Parameters.AddWithValue("@Deleted", false);
            command.Parameters.AddWithValue("@TOT_BENI_P", tmp2 ?? DBNull.Value);
            command.Parameters.AddWithValue("@TOT_BENI_R", tmp3 ?? DBNull.Value);
            command.Parameters.AddWithValue("@TOT_ARCHIT", tmp4 ?? DBNull.Value);
            command.Parameters.AddWithValue("@TOT_ARCHIT_P", tmp5 ?? DBNull.Value);
            command.Parameters.AddWithValue("@TOT_ARCHIT_R", tmp6 ?? DBNull.Value);
            command.Parameters.AddWithValue("@TOT_ARCHEO", tmp7 ?? DBNull.Value);
            command.Parameters.AddWithValue("@TOT_ARCHEO_P", tmp8 ?? DBNull.Value);
            command.Parameters.AddWithValue("@TOT_ARCHEO_R", tmp9 ?? DBNull.Value);
            command.Parameters.AddWithValue("@TOT_PARC_GIAR", tmp10 ?? DBNull.Value);
            command.Parameters.AddWithValue("@TOT_PARC_GIAR_P", tmp11 ?? DBNull.Value);
            command.Parameters.AddWithValue("@TOT_PARC_GIAR_R", tmp12 ?? DBNull.Value);
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

        static bool alreadyExists(string comparator)
        {
            Boolean exists = false;
            using (var connection = new SqlConnection(Settings.connection_string))
            {
                using (var command = connection.CreateCommand())
                {
                    command.CommandText = "SELECT TOT_BENI, TOT_BENI_P, TOT_BENI_R, TOT_ARCHIT, TOT_ARCHIT_P, TOT_ARCHIT_R, TOT_ARCHEO, TOT_ARCHEO_P, TOT_ARCHEO_R, TOT_PARC_GIAR, TOT_PARC_GIAR_P, TOT_PARC_GIAR_R FROM tbl_13029_Stage_MIBACT";
                    connection.Open();
                    using (var reader = command.ExecuteReader())
                    {
                        var indexOfColumn1 = reader.GetOrdinal("TOT_BENI");
                        var indexOfColumn2 = reader.GetOrdinal("TOT_BENI_P");
                        var indexOfColumn3 = reader.GetOrdinal("TOT_BENI_R");
                        var indexOfColumn4 = reader.GetOrdinal("TOT_ARCHIT");
                        var indexOfColumn5 = reader.GetOrdinal("TOT_ARCHIT_P");
                        var indexOfColumn6 = reader.GetOrdinal("TOT_ARCHIT_R");
                        var indexOfColumn7 = reader.GetOrdinal("TOT_ARCHEO");
                        var indexOfColumn8 = reader.GetOrdinal("TOT_ARCHEO_P");
                        var indexOfColumn9 = reader.GetOrdinal("TOT_ARCHEO_R");
                        var indexOfColumn10 = reader.GetOrdinal("TOT_PARC_GIAR");
                        var indexOfColumn11 = reader.GetOrdinal("TOT_PARC_GIAR_P");
                        var indexOfColumn12 = reader.GetOrdinal("TOT_PARC_GIAR_R");


                        while (reader.Read())
                        {
                            string value1 = reader.GetValue(indexOfColumn1).ToString();
                            string value2 = reader.GetValue(indexOfColumn2).ToString();
                            string value3 = reader.GetValue(indexOfColumn3).ToString();
                            string value4 = reader.GetValue(indexOfColumn4).ToString();
                            string value5 = reader.GetValue(indexOfColumn5).ToString();
                            string value6 = reader.GetValue(indexOfColumn6).ToString();
                            string value7 = reader.GetValue(indexOfColumn7).ToString();
                            string value8 = reader.GetValue(indexOfColumn8).ToString();
                            string value9 = reader.GetValue(indexOfColumn9).ToString();
                            string value10 = reader.GetValue(indexOfColumn10).ToString();
                            string value11 = reader.GetValue(indexOfColumn11).ToString();
                            string value12 = reader.GetValue(indexOfColumn12).ToString();

                            if (value1.Equals("")) value1 = "--";
                            if (value2.Equals("")) value2 = "--";
                            if (value3.Equals("")) value3 = "--";
                            if (value4.Equals("")) value4 = "--";
                            if (value5.Equals("")) value5 = "--";
                            if (value6.Equals("")) value6 = "--";
                            if (value7.Equals("")) value7 = "--";
                            if (value8.Equals("")) value8 = "--";
                            if (value9.Equals("")) value9 = "--";
                            if (value10.Equals("")) value10 = "--";
                            if (value11.Equals("")) value11 = "--";
                            if (value12.Equals("")) value12 = "--";

                            string comparator2 = value1 + value2 + value3 + value4 + value5 + value6 + value7 + value8 + value9 + value10 + value11 + value12;
                            if (comparator.Equals(comparator2))
                            {
                                return true;
                            }
                            else
                            {
                                exists = false;
                            }
                        }
                    }
                    connection.Close();
                }
            }
            return exists;
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
                    object s12 = null;

                    string comparator = "";
                    for (int i = 380; i <= 391; i++)
                    {
                        comparator = comparator + workSheet.Cells[row, i].Text;
                    }

                    if (!workSheet.Cells[row, 380].Text.Equals("--")) s1 = Double.Parse(workSheet.Cells[row, 380].Text);
                    if (!workSheet.Cells[row, 381].Text.Equals("--")) s2 = Double.Parse(workSheet.Cells[row, 381].Text);
                    if (!workSheet.Cells[row, 382].Text.Equals("--")) s3 = Double.Parse(workSheet.Cells[row, 382].Text);
                    if (!workSheet.Cells[row, 383].Text.Equals("--")) s4 = Double.Parse(workSheet.Cells[row, 383].Text);
                    if (!workSheet.Cells[row, 384].Text.Equals("--")) s5 = Double.Parse(workSheet.Cells[row, 384].Text);
                    if (!workSheet.Cells[row, 385].Text.Equals("--")) s6 = Double.Parse(workSheet.Cells[row, 385].Text);
                    if (!workSheet.Cells[row, 386].Text.Equals("--")) s7 = Double.Parse(workSheet.Cells[row, 386].Text);
                    if (!workSheet.Cells[row, 387].Text.Equals("--")) s8 = Double.Parse(workSheet.Cells[row, 387].Text);
                    if (!workSheet.Cells[row, 388].Text.Equals("--")) s9 = Double.Parse(workSheet.Cells[row, 388].Text);
                    if (!workSheet.Cells[row, 389].Text.Equals("--")) s10 = Double.Parse(workSheet.Cells[row, 389].Text);
                    if (!workSheet.Cells[row, 390].Text.Equals("--")) s11 = Double.Parse(workSheet.Cells[row, 390].Text);
                    if (!workSheet.Cells[row, 391].Text.Equals("--")) s12 = Double.Parse(workSheet.Cells[row, 391].Text);

                    if (!alreadyExists(comparator))
                    {
                        insertDB(s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12);
                        Console.WriteLine("Record inserito riga " + row);
                        n_inseriti++;
                    }
                }
            }
            return n_inseriti;
        }
    }
}
