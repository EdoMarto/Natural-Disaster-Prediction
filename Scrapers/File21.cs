using System.Data.OleDb;
using System.Data.SqlClient;

namespace Thesis
{
    public class File21
    {
        public static void run()
        {
            Console.WriteLine("Sto avviando File21...");
            int n_inseriti = manageExcel("Indicatori PM 10.xls");
            Console.WriteLine("Esecuzione File21 terminata. Sono stati inseriti " + n_inseriti + " nuovi record");
        }

        static void insertDB(object tmp1, object tmp2, object tmp3, object tmp4, object tmp5, object tmp6, object tmp7, object tmp8, object tmp9, object tmp10, object tmp11, object tmp12, object tmp13, object tmp14, object tmp15, object tmp16, object tmp17)
        {
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            string query = "INSERT INTO tbl_13029_Stage_Istat_PM_10_Capoluoghi (DataImportazione, DataInserimento, DataUltimaModifica, Deleted, Anno2003, Anno2004, Anno2005, Anno2006, Anno2007, Anno2008, Anno2009, Anno2010, Anno2011, Anno2012, Anno2013, Anno2014, Anno2015, Anno2016, Anno2017, Anno2018, ParentWeb_Id, Comune) " +
                "VALUES(@DataImportazione, @DataInserimento, @DataUltimaModifica, @Deleted, @Anno2003, @Anno2004, @Anno2005, @Anno2006, @Anno2007, @Anno2008, @Anno2009, @Anno2010, @Anno2011, @Anno2012, @Anno2013, @Anno2014, @Anno2015, @Anno2016, @Anno2017, @Anno2018, @ParentWeb_Id, @Comune)";
            SqlCommand command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@DataImportazione", DateTime.Now);
            command.Parameters.AddWithValue("@DataInserimento", DateTime.Now);
            command.Parameters.AddWithValue("@DataUltimaModifica", DateTime.Now);
            command.Parameters.AddWithValue("@Deleted", false);
            command.Parameters.AddWithValue("@Comune", tmp1 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Anno2003", tmp2 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Anno2004", tmp3 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Anno2005", tmp4 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Anno2006", tmp5 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Anno2007", tmp6 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Anno2008", tmp7 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Anno2009", tmp8 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Anno2010", tmp9 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Anno2011", tmp10 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Anno2012", tmp11 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Anno2013", tmp12 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Anno2014", tmp13 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Anno2015", tmp14 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Anno2016", tmp15 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Anno2017", tmp16 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Anno2018", tmp17 ?? DBNull.Value);
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
            SqlCommand command = new SqlCommand("SELECT COUNT(*) FROM tbl_13029_Stage_Istat_PM_10_Capoluoghi WHERE Comune=@t1", connection);
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

                        s1 = dr[1];
                        if (s1.ToString().Equals("") || s1.ToString().Contains("Fonte:") || s1.ToString().Contains("COMUNI") || s1.ToString().Contains("Note") || s1.ToString().Contains("_")) continue;
                        if (!dr[10].ToString().Equals("....")) s2 = dr[10];
                        if (!dr[11].ToString().Equals("....")) s3 = dr[11];
                        if (!dr[12].ToString().Equals("....")) s4 = dr[12];
                        if (!dr[13].ToString().Equals("....")) s5 = dr[13];
                        if (!dr[14].ToString().Equals("....")) s6 = dr[14];
                        if (!dr[15].ToString().Equals("....")) s7 = dr[15];
                        if (!dr[16].ToString().Equals("....")) s8 = dr[16];
                        if (!dr[17].ToString().Equals("....")) s9 = dr[17];
                        if (!dr[18].ToString().Equals("....")) s10 = dr[18];
                        if (!dr[19].ToString().Equals("....")) s11 = dr[19];
                        if (!dr[20].ToString().Equals("....")) s12 = dr[20];
                        if (!dr[21].ToString().Equals("....")) s13 = dr[21];
                        if (!dr[22].ToString().Equals("....")) s14 = dr[22];
                        if (!dr[23].ToString().Equals("....")) s15 = dr[23];
                        if (!dr[24].ToString().Equals("....")) s16 = dr[24];
                        if (!dr[25].ToString().Equals("....")) s17 = dr[25];

                        if (!alreadyExists(s1))
                        {
                            insertDB(s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14, s15, s16, s17);
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
