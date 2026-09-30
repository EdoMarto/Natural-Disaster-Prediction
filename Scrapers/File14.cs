using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Data.SqlClient;
using System.IO.Compression;
using System.Net;

namespace Thesis
{
    public class File14
    {
        public static void run()
        {
            Console.WriteLine("Sto avviando File14...");
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArguments("headless");
            IWebDriver driver = new ChromeDriver(Settings.chrome_driver_path, chromeOptions);
            driver.Navigate().GoToUrl("https://annuario.isprambiente.it/sys_ind/macro/11");
            String element = driver.FindElement(By.XPath("/html/body/div[2]/div/section/div/section/div/div/div/div[1]/ul/li[2]/div/div/div/div[3]/div/div/div/a[4]")).GetAttribute("href");
            int n_inseriti = manageExcel(element);
            driver.Close();
            Console.WriteLine("Esecuzione File14 terminata. Sono stati inseriti " + n_inseriti + " nuovi record");
        }

        static void insertDB(object tmp1, object tmp2, object tmp3, object tmp4, object tmp5, object tmp6, object tmp7, object tmp8, object tmp9, object tmp10, object tmp11, object tmp12)
        {
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            string query = "INSERT INTO tbl_13029_Stage_Ispra_Pericolosita_Frana (DataImportazione, DataInserimento, DataUltimaModifica, Deleted, CODREGcodiceregione, CODPROcodiceprovincia, Provincia, Regione, AreaProvincia, AreeapericolositadaFranamoltoelevata, AreeapericolositadaFranaelevata, AreeapericolositadaFranamedia, AreeapericolositadaFranamoderata, AreeapericolositadaFranaelevataemoltoelevata, PERC_AreeapericolositadaFranamoderata, ParentWeb_Id, AreeDiAttenzione) " +
                "VALUES(@DataImportazione, @DataInserimento, @DataUltimaModifica, @Deleted, @CODREGcodiceregione, @CODPROcodiceprovincia, @Provincia, @Regione, @AreaProvincia, @AreeapericolositadaFranamoltoelevata, @AreeapericolositadaFranaelevata, @AreeapericolositadaFranamedia, @AreeapericolositadaFranamoderata, @AreeapericolositadaFranaelevataemoltoelevata, @PERC_AreeapericolositadaFranamoderata, @ParentWeb_Id, @AreeDiAttenzione)";
            SqlCommand command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@DataImportazione", DateTime.Now);
            command.Parameters.AddWithValue("@DataInserimento", DateTime.Now);
            command.Parameters.AddWithValue("@DataUltimaModifica", DateTime.Now);
            command.Parameters.AddWithValue("@Deleted", false);
            command.Parameters.AddWithValue("@CODREGcodiceregione", tmp1 ?? DBNull.Value);
            command.Parameters.AddWithValue("@CODPROcodiceprovincia", tmp2 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Provincia", tmp3 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Regione", tmp4 ?? DBNull.Value);
            command.Parameters.AddWithValue("@AreaProvincia", tmp5 ?? DBNull.Value);
            command.Parameters.AddWithValue("@AreeapericolositadaFranamoltoelevata", tmp6 ?? DBNull.Value);
            command.Parameters.AddWithValue("@AreeapericolositadaFranaelevata", tmp7 ?? DBNull.Value);
            command.Parameters.AddWithValue("@AreeapericolositadaFranamedia", tmp8 ?? DBNull.Value);
            command.Parameters.AddWithValue("@AreeapericolositadaFranamoderata", tmp9 ?? DBNull.Value);
            command.Parameters.AddWithValue("@AreeDiAttenzione", tmp10 ?? DBNull.Value);
            command.Parameters.AddWithValue("@AreeapericolositadaFranaelevataemoltoelevata", tmp11 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_AreeapericolositadaFranamoderata", tmp12 ?? DBNull.Value);
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

        static bool alreadyExists(object pid)
        {
            int count = 0;
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            SqlCommand command = new SqlCommand("SELECT COUNT(*) FROM tbl_13029_Stage_Ispra_Pericolosita_Frana WHERE CODPROcodiceprovincia=@pid", connection);
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
                    if (entry.Name == "Tab2_Aree_pericolosita_frane_Province_ADA_ed2021_rev.xlsx")
                        entry.ExtractToFile("../../../Excels/Automatici/Tab2_Aree_pericolosita_frane_Province_ADA_ed2021_rev.xlsx", true);

            using (var pck = new ExcelPackage(new FileInfo("../../../Excels/Automatici/Tab2_Aree_pericolosita_frane_Province_ADA_ed2021_rev.xlsx")))
            {
                ExcelWorksheet workSheet = pck.Workbook.Worksheets["Aree_pericol_frane_province"];
                var start = new ExcelCellAddress(5, 1);
                var end = new ExcelCellAddress(111, 12);
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
                    try
                    {
                        s1 = int.Parse(workSheet.Cells[row, 1].Text.Replace(".", ""));
                        s2 = int.Parse(workSheet.Cells[row, 2].Text.Replace(".", ""));
                        s3 = workSheet.Cells[row, 3].Text;
                        s4 = workSheet.Cells[row, 4].Text;
                        s5 = int.Parse(workSheet.Cells[row, 5].Text.Replace(".", ""));
                        s6 = int.Parse(workSheet.Cells[row, 6].Text.Replace(".", ""));
                        s7 = int.Parse(workSheet.Cells[row, 7].Text.Replace(".", ""));
                        s8 = int.Parse(workSheet.Cells[row, 8].Text.Replace(".", ""));
                        s9 = int.Parse(workSheet.Cells[row, 9].Text.Replace(".", ""));
                        s10 = int.Parse(workSheet.Cells[row, 10].Text.Replace(".", ""));
                        s11 = int.Parse(workSheet.Cells[row, 11].Text.Replace(".", ""));
                        s12 = Double.Parse(workSheet.Cells[row, 12].Text.Replace("%", ""));
                    }
                    catch (InvalidCastException e)
                    {
                        Console.WriteLine("Errore: " + e.ToString());
                    }

                    if (!alreadyExists(s2))
                    {
                        insertDB(s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12);
                        Console.WriteLine("Record inserito");
                        n_inseriti++;
                    }
                }
            }
            return n_inseriti;
        }
    }
}
