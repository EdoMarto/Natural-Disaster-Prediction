using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System.Data.SqlClient;

namespace Thesis
{
    public class File10
    {
        public static void run()
        {
            Console.WriteLine("Sto avviando File10...");
            int n_inseriti = manageExcel("Indicatori_Intero_territorio_nazionale.xlsx");
            Console.WriteLine("Esecuzione File10 terminata. Sono stati inseriti " + n_inseriti + " nuovi record");
        }

        static void insertDB(object tmp1, object tmp2, object tmp3, object tmp4, object tmp5, object tmp6, object tmp7, object tmp8, object tmp9, object tmp10, object tmp11, object tmp12, object tmp13, object tmp14, object tmp15, object tmp16, object tmp17, object tmp18, object tmp19, object tmp20, object tmp21, object tmp22, object tmp23, object tmp24, object tmp25, object tmp26, object tmp27, object tmp28, object tmp29, object tmp30, object tmp31, object tmp32, object tmp33, object tmp34, object tmp35, object tmp36, object tmp37, object tmp38, object tmp39, object tmp40, object tmp41, object tmp42, object tmp43, object tmp44, object tmp45, object tmp46, object tmp47, object tmp48, object tmp49, object tmp50, object tmp51)
        {
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            string query = "INSERT INTO tbl_13029_Stage_Ispra (PAI_AREAP4, DataInserimento, DataUltimaModifica, Deleted, PAI_AREAP3, PAI_AREAP2, PAI_AREAP1, PAI_AREAAA, PAI_AREAP3_P4, PAI_AREAAA_P, PAI_AREAAA_R, PAI_AREAP1_P, PAI_AREAP1_R, PAI_AREAP2_P, PAI_AREAP2_R, PAI_AREAP3_P, PAI_AREAP3_R, PAI_AREAP3_P4_P, PAI_AREAP3_P4_R, PAI_POPP4, PAI_POPP3, PAI_POPP2, PAI_POPP1, PAI_POPAA, PAI_POPAA_P, PAI_POPAA_R, PAI_POPP1_P, PAI_POPP1_R, PAI_POPP2_P, PAI_POPP2_R, PAI_POPP3_P, PAI_POPP3_R, PAI_POPP4_P, PAI_POPP4_R, PAI_POPP3_P4_P, PAI_POPP3_P4_R, IDR_POPP3, IDR_POPP2, IDR_POPP1, IDR_POPP1_P, IDR_POPP1_R, IDR_POPP2_P, IDR_POPP2_R, IDR_POPP3_P, IDR_POPP3_R, IDR_AREAP1, IDR_AREAP1_P, IDR_AREAP1_R, IDR_AREAP2, IDR_AREAP2_P, IDR_AREAP2_R, IDR_AREAP3, IDR_AREAP3_P, IDR_AREAP3_R, DataImportazione, ParentWeb_Id) " +
                            "VALUES(@PAI_AREAP4, @DataInserimento, @DataUltimaModifica, @Deleted, @PAI_AREAP3, @PAI_AREAP2, @PAI_AREAP1, @PAI_AREAAA, @PAI_AREAP3_P4, @PAI_AREAAA_P, @PAI_AREAAA_R, @PAI_AREAP1_P, @PAI_AREAP1_R, @PAI_AREAP2_P, @PAI_AREAP2_R, @PAI_AREAP3_P, @PAI_AREAP3_R, @PAI_AREAP3_P4_P, @PAI_AREAP3_P4_R, @PAI_POPP4, @PAI_POPP3, @PAI_POPP2, @PAI_POPP1, @PAI_POPAA, @PAI_POPAA_P, @PAI_POPAA_R, @PAI_POPP1_P, @PAI_POPP1_R, @PAI_POPP2_P, @PAI_POPP2_R, @PAI_POPP3_P, @PAI_POPP3_R, @PAI_POPP4_P, @PAI_POPP4_R, @PAI_POPP3_P4_P, @PAI_POPP3_P4_R, @IDR_POPP3, @IDR_POPP2, @IDR_POPP1, @IDR_POPP1_P, @IDR_POPP1_R, @IDR_POPP2_P, @IDR_POPP2_R, @IDR_POPP3_P, @IDR_POPP3_R, @IDR_AREAP1, @IDR_AREAP1_P, @IDR_AREAP1_R, @IDR_AREAP2, @IDR_AREAP2_P, @IDR_AREAP2_R, @IDR_AREAP3, @IDR_AREAP3_P, @IDR_AREAP3_R, @DataImportazione, @ParentWeb_Id)";
            SqlCommand command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@PAI_AREAP4", tmp1 ?? DBNull.Value);
            command.Parameters.AddWithValue("@DataInserimento", DateTime.Now);
            command.Parameters.AddWithValue("@DataUltimaModifica", DateTime.Now);
            command.Parameters.AddWithValue("@Deleted", false);
            command.Parameters.AddWithValue("@PAI_AREAP3", tmp2 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_AREAP2", tmp3 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_AREAP1", tmp4 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_AREAAA", tmp5 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_AREAP3_P4", tmp6 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_AREAAA_P", tmp7 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_AREAAA_R", tmp8 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_AREAP1_P", tmp9 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_AREAP1_R", tmp10 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_AREAP2_P", tmp11 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_AREAP2_R", tmp12 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_AREAP3_P", tmp13 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_AREAP3_R", tmp14 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_AREAP3_P4_P", tmp15 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_AREAP3_P4_R", tmp16 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_POPP4", tmp17 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_POPP3", tmp18 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_POPP2", tmp19 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_POPP1", tmp20 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_POPAA", tmp21 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_POPAA_P", tmp22 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_POPAA_R", tmp23 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_POPP1_P", tmp24 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_POPP1_R", tmp25 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_POPP2_P", tmp26 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_POPP2_R", tmp27 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_POPP3_P", tmp28 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_POPP3_R", tmp29 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_POPP4_P", tmp30 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_POPP4_R", tmp31 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_POPP3_P4_P", tmp32 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PAI_POPP3_P4_R", tmp33 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDR_POPP3", tmp34 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDR_POPP2", tmp35 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDR_POPP1", tmp36 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDR_POPP1_P", tmp37 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDR_POPP1_R", tmp38 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDR_POPP2_P", tmp39 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDR_POPP2_R", tmp40 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDR_POPP3_P", tmp41 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDR_POPP3_R", tmp42 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDR_AREAP1", tmp43 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDR_AREAP1_P", tmp44 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDR_AREAP1_R", tmp45 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDR_AREAP2", tmp46 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDR_AREAP2_P", tmp47 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDR_AREAP2_R", tmp48 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDR_AREAP3", tmp49 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDR_AREAP3_P", tmp50 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDR_AREAP3_R", tmp51 ?? DBNull.Value);
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
                    command.CommandText = "SELECT IDR_AREAP1, IDR_AREAP2, IDR_AREAP3 FROM tbl_13029_Stage_Ispra";
                    connection.Open();
                    using (var reader = command.ExecuteReader())
                    {
                        var indexOfColumn1 = reader.GetOrdinal("IDR_AREAP1");
                        var indexOfColumn2 = reader.GetOrdinal("IDR_AREAP2");
                        var indexOfColumn3 = reader.GetOrdinal("IDR_AREAP3");


                        while (reader.Read())
                        {
                            string value1 = reader.GetValue(indexOfColumn1).ToString();
                            string value2 = reader.GetValue(indexOfColumn2).ToString();

                            if (value1.Equals("")) value1 = "--";
                            if (value2.Equals("")) value2 = "--";

                            string comparator2 = value1 + value2;
                            if (comparator.Equals(comparator2))
                            {
                                connection.Close();
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
                    object s13 = null;
                    object s14 = null;
                    object s15 = null;
                    object s16 = null;
                    object s17 = null;
                    object s18 = null;
                    object s19 = null;
                    object s20 = null;
                    object s21 = null;
                    object s22 = null;
                    object s23 = null;
                    object s24 = null;
                    object s25 = null;
                    object s26 = null;
                    object s27 = null;
                    object s28 = null;
                    object s29 = null;
                    object s30 = null;
                    object s31 = null;
                    object s32 = null;
                    object s33 = null;
                    object s34 = null;
                    object s35 = null;
                    object s36 = null;
                    object s37 = null;
                    object s38 = null;
                    object s39 = null;
                    object s40 = null;
                    object s41 = null;
                    object s42 = null;
                    object s43 = null;
                    object s44 = null;
                    object s45 = null;
                    object s46 = null;
                    object s47 = null;
                    object s48 = null;
                    object s49 = null;
                    object s50 = null;
                    object s51 = null;

                    //IDR_AREAP1 + IDR_AREAP2
                    string comparator = workSheet.Cells[row, 125].Text + workSheet.Cells[row, 128].Text;

                    if (!workSheet.Cells[row, 82].Text.Equals("--")) s1 = Double.Parse(workSheet.Cells[row, 82].Text);
                    if (!workSheet.Cells[row, 83].Text.Equals("--")) s2 = Double.Parse(workSheet.Cells[row, 83].Text);
                    if (!workSheet.Cells[row, 84].Text.Equals("--")) s3 = Double.Parse(workSheet.Cells[row, 84].Text);
                    if (!workSheet.Cells[row, 85].Text.Equals("--")) s4 = Double.Parse(workSheet.Cells[row, 85].Text);
                    if (!workSheet.Cells[row, 86].Text.Equals("--")) s5 = Double.Parse(workSheet.Cells[row, 86].Text);
                    if (!workSheet.Cells[row, 87].Text.Equals("--")) s6 = Double.Parse(workSheet.Cells[row, 87].Text);
                    if (!workSheet.Cells[row, 88].Text.Equals("--")) s7 = Double.Parse(workSheet.Cells[row, 88].Text);
                    if (!workSheet.Cells[row, 89].Text.Equals("--")) s8 = Double.Parse(workSheet.Cells[row, 89].Text);
                    if (!workSheet.Cells[row, 90].Text.Equals("--")) s9 = Double.Parse(workSheet.Cells[row, 90].Text);
                    if (!workSheet.Cells[row, 91].Text.Equals("--")) s10 = Double.Parse(workSheet.Cells[row, 91].Text);
                    if (!workSheet.Cells[row, 92].Text.Equals("--")) s11 = Double.Parse(workSheet.Cells[row, 92].Text);
                    if (!workSheet.Cells[row, 93].Text.Equals("--")) s12 = Double.Parse(workSheet.Cells[row, 93].Text);
                    if (!workSheet.Cells[row, 94].Text.Equals("--")) s13 = Double.Parse(workSheet.Cells[row, 94].Text);
                    if (!workSheet.Cells[row, 95].Text.Equals("--")) s14 = Double.Parse(workSheet.Cells[row, 95].Text);
                    if (!workSheet.Cells[row, 96].Text.Equals("--")) s15 = Double.Parse(workSheet.Cells[row, 96].Text);
                    if (!workSheet.Cells[row, 97].Text.Equals("--")) s16 = Double.Parse(workSheet.Cells[row, 97].Text);
                    if (!workSheet.Cells[row, 98].Text.Equals("--")) s17 = Double.Parse(workSheet.Cells[row, 98].Text);
                    if (!workSheet.Cells[row, 99].Text.Equals("--")) s18 = Double.Parse(workSheet.Cells[row, 99].Text);
                    if (!workSheet.Cells[row, 100].Text.Equals("--")) s19 = Double.Parse(workSheet.Cells[row, 100].Text);
                    if (!workSheet.Cells[row, 101].Text.Equals("--")) s20 = Double.Parse(workSheet.Cells[row, 101].Text);
                    if (!workSheet.Cells[row, 102].Text.Equals("--")) s21 = Double.Parse(workSheet.Cells[row, 102].Text);
                    if (!workSheet.Cells[row, 103].Text.Equals("--")) s22 = Double.Parse(workSheet.Cells[row, 103].Text);
                    if (!workSheet.Cells[row, 104].Text.Equals("--")) s23 = Double.Parse(workSheet.Cells[row, 104].Text);
                    if (!workSheet.Cells[row, 105].Text.Equals("--")) s24 = Double.Parse(workSheet.Cells[row, 105].Text);
                    if (!workSheet.Cells[row, 106].Text.Equals("--")) s25 = Double.Parse(workSheet.Cells[row, 106].Text);
                    if (!workSheet.Cells[row, 107].Text.Equals("--")) s26 = Double.Parse(workSheet.Cells[row, 107].Text);
                    if (!workSheet.Cells[row, 108].Text.Equals("--")) s27 = Double.Parse(workSheet.Cells[row, 108].Text);
                    if (!workSheet.Cells[row, 109].Text.Equals("--")) s28 = Double.Parse(workSheet.Cells[row, 109].Text);
                    if (!workSheet.Cells[row, 110].Text.Equals("--")) s29 = Double.Parse(workSheet.Cells[row, 110].Text);
                    if (!workSheet.Cells[row, 111].Text.Equals("--")) s30 = Double.Parse(workSheet.Cells[row, 111].Text);
                    if (!workSheet.Cells[row, 112].Text.Equals("--")) s31 = Double.Parse(workSheet.Cells[row, 112].Text);
                    if (!workSheet.Cells[row, 114].Text.Equals("--")) s32 = Double.Parse(workSheet.Cells[row, 114].Text);
                    if (!workSheet.Cells[row, 115].Text.Equals("--")) s33 = Double.Parse(workSheet.Cells[row, 115].Text);
                    if (!workSheet.Cells[row, 116].Text.Equals("--")) s34 = Double.Parse(workSheet.Cells[row, 116].Text);
                    if (!workSheet.Cells[row, 117].Text.Equals("--")) s35 = Double.Parse(workSheet.Cells[row, 117].Text);
                    if (!workSheet.Cells[row, 118].Text.Equals("--")) s36 = Double.Parse(workSheet.Cells[row, 118].Text);
                    if (!workSheet.Cells[row, 119].Text.Equals("--")) s37 = Double.Parse(workSheet.Cells[row, 119].Text);
                    if (!workSheet.Cells[row, 120].Text.Equals("--")) s38 = Double.Parse(workSheet.Cells[row, 120].Text);
                    if (!workSheet.Cells[row, 121].Text.Equals("--")) s39 = Double.Parse(workSheet.Cells[row, 121].Text);
                    if (!workSheet.Cells[row, 122].Text.Equals("--")) s40 = Double.Parse(workSheet.Cells[row, 122].Text);
                    if (!workSheet.Cells[row, 123].Text.Equals("--")) s41 = Double.Parse(workSheet.Cells[row, 123].Text);
                    if (!workSheet.Cells[row, 124].Text.Equals("--")) s42 = Double.Parse(workSheet.Cells[row, 124].Text);
                    if (!workSheet.Cells[row, 125].Text.Equals("--")) s43 = Double.Parse(workSheet.Cells[row, 125].Text);
                    if (!workSheet.Cells[row, 126].Text.Equals("--")) s44 = Double.Parse(workSheet.Cells[row, 126].Text);
                    if (!workSheet.Cells[row, 127].Text.Equals("--")) s45 = Double.Parse(workSheet.Cells[row, 127].Text);
                    if (!workSheet.Cells[row, 128].Text.Equals("--")) s46 = Double.Parse(workSheet.Cells[row, 128].Text);
                    if (!workSheet.Cells[row, 129].Text.Equals("--")) s47 = Double.Parse(workSheet.Cells[row, 129].Text);
                    if (!workSheet.Cells[row, 130].Text.Equals("--")) s48 = Double.Parse(workSheet.Cells[row, 130].Text);
                    if (!workSheet.Cells[row, 131].Text.Equals("--")) s49 = Double.Parse(workSheet.Cells[row, 131].Text);
                    if (!workSheet.Cells[row, 132].Text.Equals("--")) s50 = Double.Parse(workSheet.Cells[row, 132].Text);
                    if (!workSheet.Cells[row, 133].Text.Equals("--")) s51 = Double.Parse(workSheet.Cells[row, 133].Text);
                    if (!alreadyExists(comparator))
                    {
                        insertDB(s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14, s15, s16, s17, s18, s19, s20, s21, s22, s23, s24, s25, s26, s27, s28, s29, s30, s31, s32, s33, s34, s35, s36, s37, s38, s39, s40, s41, s42, s43, s44, s45, s46, s47, s48, s49, s50, s51);
                        Console.WriteLine("Record inserito riga " + row);
                        n_inseriti++;
                    }
                }
            }
            return n_inseriti;
        }
    }
}
