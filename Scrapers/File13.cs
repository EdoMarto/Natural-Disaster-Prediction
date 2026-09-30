using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Data.SqlClient;
using System.IO.Compression;
using System.Net;

namespace Thesis
{
    public class File13
    {
        public static void run()
        {
            Console.WriteLine("Sto avviando File13...");
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArguments("headless");
            IWebDriver driver = new ChromeDriver(Settings.chrome_driver_path, chromeOptions);
            driver.Navigate().GoToUrl("https://annuario.isprambiente.it/sys_ind/macro/11");
            String element = driver.FindElement(By.XPath("/html/body/div[2]/div/section/div/section/div/div/div/div[1]/ul/li[7]/div/div/div/div[3]/div/div/div/a[4]")).GetAttribute("href");
            int n_inseriti = manageExcel(element);
            driver.Close();
            Console.WriteLine("Esecuzione File13 terminata. Sono stati inseriti " + n_inseriti + " nuovi record");
        }

        static void insertDB(object tmp1, object tmp2, object tmp3, object tmp4, object tmp5, object tmp6, object tmp7, object tmp8)
        {
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            string query = "INSERT INTO tbl_13029_Stage_ISPRA_Eventi_Pluviali (DataImportazione, DataInserimento, DataUltimaModifica, Deleted, Regione, Province, BacinoIdrografico, DurataComplessivaPrecipitazioneOre, PluviometroMassimaPrecipitazioneGiornaliera, PluviometroPrecipitazioneTotali, ParentWeb_Id, InizioEvento, FineEvento) " +
                "VALUES(@DataImportazione, @DataInserimento, @DataUltimaModifica, @Deleted, @Regione, @Province, @BacinoIdrografico, @DurataComplessivaPrecipitazioneOre, @PluviometroMassimaPrecipitazioneGiornaliera, @PluviometroPrecipitazioneTotali, @ParentWeb_Id, @InizioEvento, @FineEvento)";
            SqlCommand command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@DataImportazione", DateTime.Now);
            command.Parameters.AddWithValue("@DataInserimento", DateTime.Now);
            command.Parameters.AddWithValue("@DataUltimaModifica", DateTime.Now);
            command.Parameters.AddWithValue("@Deleted", false);
            command.Parameters.AddWithValue("@InizioEvento", tmp1 ?? DBNull.Value);
            command.Parameters.AddWithValue("@FineEvento", tmp2 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Regione", tmp3 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Province", tmp4 ?? DBNull.Value);
            command.Parameters.AddWithValue("@BacinoIdrografico", tmp5 ?? DBNull.Value);
            command.Parameters.AddWithValue("@DurataComplessivaPrecipitazioneOre", tmp6 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PluviometroMassimaPrecipitazioneGiornaliera", tmp7 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PluviometroPrecipitazioneTotali", tmp8 ?? DBNull.Value);
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
            SqlCommand command = new SqlCommand("SELECT COUNT(*) FROM tbl_13029_Stage_ISPRA_Eventi_Pluviali WHERE PluviometroMassimaPrecipitazioneGiornaliera=@t1", connection);
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

            using (ZipArchive zip = ZipFile.Open("../../../Excels/" + filename, ZipArchiveMode.Read))
                foreach (ZipArchiveEntry entry in zip.Entries)
                    if (entry.Name == "Tabella1_Aspetti pluviometrici degli eventi alluvionali avvenuti nel 2020.xlsx")
                        entry.ExtractToFile("../../../Excels/Tabella1_Aspetti pluviometrici degli eventi alluvionali avvenuti nel 2020.xlsx", true);

            using (var pck = new ExcelPackage(new FileInfo("../../../Excels/Tabella1_Aspetti pluviometrici degli eventi alluvionali avvenuti nel 2020.xlsx")))
            {
                ExcelWorksheet workSheet = pck.Workbook.Worksheets["Foglio1"];
                var start = new ExcelCellAddress(2, 1);
                var end = new ExcelCellAddress(18, 7);
                for (int row = start.Row; row <= end.Row; row++)
                {
                    object s0 = null;
                    object s1 = null;
                    object s2 = null;
                    object s3 = null;
                    object s4 = null;
                    object s5 = null;
                    object s6 = null;
                    object s7 = null;
                    try
                    {

                        if (!((workSheet.Cells[row, 1].Text).Equals("")))
                        {
                            if (workSheet.Cells[row, 1].Text.Contains("-"))
                            {
                                s0 = DateTime.Parse(workSheet.Cells[row, 1].Text.Substring(0, workSheet.Cells[row, 1].Text.IndexOf("-")) + workSheet.Cells[row, 1].Text.Substring(workSheet.Cells[row, 1].Text.IndexOf("/"), workSheet.Cells[row, 1].Text.Length - workSheet.Cells[row, 1].Text.IndexOf("/")));
                                s1 = DateTime.Parse(workSheet.Cells[row, 1].Text.Substring(workSheet.Cells[row, 1].Text.IndexOf("-") + 1, workSheet.Cells[row, 1].Text.IndexOf("/") - workSheet.Cells[row, 1].Text.IndexOf("-") - 1) + workSheet.Cells[row, 1].Text.Substring(workSheet.Cells[row, 1].Text.IndexOf("/"), workSheet.Cells[row, 1].Text.Length - workSheet.Cells[row, 1].Text.IndexOf("/")));
                            }
                            else if (workSheet.Cells[row, 1].Text.Contains("."))
                            {
                                s0 = DateTime.Parse(workSheet.Cells[row, 1].Text.Substring(0, workSheet.Cells[row, 1].Text.IndexOf(".")) + workSheet.Cells[row, 1].Text.Substring(workSheet.Cells[row, 1].Text.IndexOf("/"), workSheet.Cells[row, 1].Text.Length - workSheet.Cells[row, 1].Text.IndexOf("/")));
                                s1 = DateTime.Parse(workSheet.Cells[row, 1].Text.Substring(workSheet.Cells[row, 1].Text.IndexOf(".") + 1, workSheet.Cells[row, 1].Text.IndexOf("/") - workSheet.Cells[row, 1].Text.IndexOf(".") - 1) + workSheet.Cells[row, 1].Text.Substring(workSheet.Cells[row, 1].Text.IndexOf("/"), workSheet.Cells[row, 1].Text.Length - workSheet.Cells[row, 1].Text.IndexOf("/")));
                            }
                            else
                            {
                                s0 = DateTime.Parse(workSheet.Cells[row, 1].Text);
                                s1 = DateTime.Parse(workSheet.Cells[row, 1].Text);
                            }
                        }
                        if (!((workSheet.Cells[row, 2].Text).Equals(""))) s2 = workSheet.Cells[row, 2].Text;
                        if (!((workSheet.Cells[row, 3].Text).Equals(""))) s3 = workSheet.Cells[row, 3].Text;
                        if (!((workSheet.Cells[row, 4].Text).Equals(""))) s4 = workSheet.Cells[row, 4].Text;
                        if (!((workSheet.Cells[row, 5].Text).Equals(""))) s5 = workSheet.Cells[row, 5].Text;
                        if (!((workSheet.Cells[row, 6].Text).Equals(""))) s6 = workSheet.Cells[row, 6].Text;
                        if (!((workSheet.Cells[row, 7].Text).Equals(""))) s7 = workSheet.Cells[row, 7].Text;
                    }
                    catch (InvalidCastException e)
                    {
                        Console.WriteLine("Errore: " + e.ToString());
                    }

                    if (!alreadyExists(s6))
                    {
                        insertDB(s0, s1, s2, s3, s4, s5, s6, s7);
                        Console.WriteLine("Record inserito");
                        n_inseriti++;
                    }
                }
            }
            return n_inseriti;
        }
    }
}
