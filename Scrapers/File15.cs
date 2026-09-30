using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Data.SqlClient;
using System.IO.Compression;
using System.Net;

namespace Thesis
{
    public class File15
    {
        public static void run()
        {
            Console.WriteLine("Sto avviando File15...");
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArguments("headless");
            IWebDriver driver = new ChromeDriver(Settings.chrome_driver_path, chromeOptions);
            driver.Navigate().GoToUrl("https://annuario.isprambiente.it/sys_ind/macro/11");
            String element = driver.FindElement(By.XPath("/html/body/div[2]/div/section/div/section/div/div/div/div[1]/ul/li[12]/div/div/div/div[3]/div/div/div/a[4]")).GetAttribute("href");
            int n_inseriti = manageExcel(element);
            driver.Close();
            Console.WriteLine("Esecuzione File15 terminata. Sono stati inseriti " + n_inseriti + " nuovi record");
        }

        static void insertDB(object tmp1, object tmp2, object tmp3, object tmp4, object tmp5, object tmp6, object tmp7, object tmp8, object tmp9, object tmp10)
        {
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            string query = "INSERT INTO tbl_13029_Stage_Ispra_Popolazione_in_Aree_Allagabili (DataImportazione, DataInserimento, DataUltimaModifica, Deleted, ProvinciaID, ProvinciaNome, PopolazioneProvincia, HPHNumeroAbitanti, PERC_HPHNumeroAbitanti, MPHNumeroAbitanti, PERC_MPHAbitanti, LPHNumeroAbitanti, PERC_LPHNumeroAbitanti, ParentWeb_Id, Regione) " +
                "VALUES(@DataImportazione, @DataInserimento, @DataUltimaModifica, @Deleted, @ProvinciaID, @ProvinciaNome, @PopolazioneProvincia, @HPHNumeroAbitanti, @PERC_HPHNumeroAbitanti, @MPHNumeroAbitanti, @PERC_MPHAbitanti, @LPHNumeroAbitanti, @PERC_LPHNumeroAbitanti, @ParentWeb_Id, @Regione)";
            SqlCommand command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@DataImportazione", DateTime.Now);
            command.Parameters.AddWithValue("@DataInserimento", DateTime.Now);
            command.Parameters.AddWithValue("@DataUltimaModifica", DateTime.Now);
            command.Parameters.AddWithValue("@Deleted", false);
            command.Parameters.AddWithValue("@ProvinciaID", tmp1 ?? DBNull.Value);
            command.Parameters.AddWithValue("@ProvinciaNome", tmp2 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PopolazioneProvincia", tmp3 ?? DBNull.Value);
            command.Parameters.AddWithValue("@HPHNumeroAbitanti", tmp4 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_HPHNumeroAbitanti", tmp5 ?? DBNull.Value);
            command.Parameters.AddWithValue("@MPHNumeroAbitanti", tmp6 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_MPHAbitanti", tmp7 ?? DBNull.Value);
            command.Parameters.AddWithValue("@LPHNumeroAbitanti", tmp8 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_LPHNumeroAbitanti", tmp9 ?? DBNull.Value);
            command.Parameters.AddWithValue("@ParentWeb_Id", 13029);
            command.Parameters.AddWithValue("@Regione", tmp10);

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

        static bool alreadyExists(object pid)
        {
            int count = 0;
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            SqlCommand command = new SqlCommand("SELECT COUNT(*) FROM tbl_13029_Stage_Ispra_Popolazione_in_Aree_Allagabili WHERE ProvinciaID=@pid", connection);
            command.Parameters.AddWithValue("@pid", pid);

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

        static int manageExcel(String url)
        {
            int n_inseriti = 0;
            String filename = url.Substring(url.LastIndexOf("/") + 1);
            filename = "allegati_" + filename;
            using (WebClient wc = new WebClient())
            {
                wc.DownloadFile(
                    new System.Uri(url),
                    "../../../Excels/Automatici/" + filename
                );
            }

            using (ZipArchive zip = ZipFile.Open("../../../Excels/Automatici/" + filename, ZipArchiveMode.Read))
                foreach (ZipArchiveEntry entry in zip.Entries)
                    if (entry.Name == "Tabella1_Popolazione_in_Aree_allagabili_province.xlsx")
                        entry.ExtractToFile("../../../Excels/Automatici/Tabella1_Popolazione_in_Aree_allagabili_province.xlsx", true);

            using (var pck = new ExcelPackage(new FileInfo("../../../Excels/Automatici/Tabella1_Popolazione_in_Aree_allagabili_province.xlsx")))
            {
                ExcelWorksheet workSheet = pck.Workbook.Worksheets["popolazione_esposta_province"];
                var start = new ExcelCellAddress(3, 1);
                var end = new ExcelCellAddress(109, 10);
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
                    try
                    {
                        s1 = workSheet.Cells[row, 1].Text;
                        s2 = int.Parse(workSheet.Cells[row, 2].Text.Replace(".", ""));
                        s3 = workSheet.Cells[row, 3].Text;
                        s4 = int.Parse(workSheet.Cells[row, 4].Text.Replace(".", ""));
                        s5 = int.Parse(workSheet.Cells[row, 5].Text.Replace(".", ""));
                        s6 = Double.Parse(workSheet.Cells[row, 6].Text.Replace(".", ""));
                        s7 = int.Parse(workSheet.Cells[row, 7].Text.Replace(".", ""));
                        s8 = Double.Parse(workSheet.Cells[row, 8].Text.Replace(".", ""));
                        s9 = int.Parse(workSheet.Cells[row, 9].Text.Replace(".", ""));
                        s10 = Double.Parse(workSheet.Cells[row, 10].Text.Replace(".", ""));
                    }
                    catch (InvalidCastException e)
                    {
                        Console.WriteLine("Errore: " + e.ToString());
                    }

                    if (!alreadyExists(s2))
                    {
                        insertDB(s2, s3, s4, s5, s6, s7, s8, s9, s10, s1);
                        Console.WriteLine("Record inserito");
                        n_inseriti++;
                    }
                }
            }
            return n_inseriti;
        }
    }
}
