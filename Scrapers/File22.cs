using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Data.SqlClient;
using System.Net;

namespace Thesis
{
    public class File22
    {
        public static void run()
        {
            Console.WriteLine("Sto avviando File22...");
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArguments("headless");
            IWebDriver driver = new ChromeDriver(Settings.chrome_driver_path, chromeOptions);
            driver.Navigate().GoToUrl("https://www.istat.it/it/archivio/263811");
            String element = driver.FindElement(By.XPath("/html/body/div[3]/div/div/div/div[3]/div[4]/div[2]/ul/li[2]/span[1]/a")).GetAttribute("href");
            int n_inseriti = manageExcel(element);
            driver.Close();
            Console.WriteLine("Esecuzione File22 terminata. Sono stati inseriti " + n_inseriti + " nuovi record");
        }

        static void insertDB(object comuni, object tmpmedia1, object tmpmedia2, object tmpmedia3, object tmpmedia4)
        {
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            string query = "INSERT INTO tbl_13029_Stage_Istat_Precipitazioni (DataImportazione, DataInserimento, DataUltimaModifica, Deleted, COMUNI, PrecipitazioneTotaleDifferenza2019dalValoreMedio2007_2016, PrecipitazionetotaleValoreMedio2007_2016, PrecipitazioneTotaleAnomalia2019dalValoreClimatico1971_2000, PrecipitazioneTotaleValoreClimatico1971_2000, ParentWeb_Id) " +
                "VALUES(@DataImportazione, @DataInserimento, @DataUltimaModifica, @Deleted, @COMUNI, @PrecipitazioneTotaleDifferenza2019dalValoreMedio2007_2016, @PrecipitazionetotaleValoreMedio2007_2016, @PrecipitazioneTotaleAnomalia2019dalValoreClimatico1971_2000, @PrecipitazioneTotaleValoreClimatico1971_2000, @ParentWeb_Id)";
            SqlCommand command = new SqlCommand(query, connection);

            if (comuni != null)
            {
                command.Parameters.AddWithValue("@COMUNI", comuni);
            }
            command.Parameters.AddWithValue("@DataImportazione", DateTime.Now);
            command.Parameters.AddWithValue("@DataInserimento", DateTime.Now);
            command.Parameters.AddWithValue("@DataUltimaModifica", DateTime.Now);
            command.Parameters.AddWithValue("@Deleted", false);
            command.Parameters.AddWithValue("@PrecipitazioneTotaleDifferenza2019dalValoreMedio2007_2016", tmpmedia1 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PrecipitazionetotaleValoreMedio2007_2016", tmpmedia2 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PrecipitazioneTotaleAnomalia2019dalValoreClimatico1971_2000", tmpmedia3 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PrecipitazioneTotaleValoreClimatico1971_2000", tmpmedia4 ?? DBNull.Value);
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

        static bool alreadyExists(object comune)
        {
            int count = 0;
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            SqlCommand command = new SqlCommand("SELECT COUNT(*) FROM tbl_13029_Stage_Istat_Precipitazioni WHERE COMUNI=@comune", connection);
            command.Parameters.AddWithValue("@comune", (string)comune);

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
            using (WebClient wc = new WebClient())
            {
                wc.DownloadFile(
                    new System.Uri(url),
                    "../../../Excels/Automatici/" + filename
                );
            }

            using (var pck = new ExcelPackage(new FileInfo("../../../Excels/Automatici/" + filename)))
            {
                ExcelWorksheet workSheet = pck.Workbook.Worksheets["Tavola_2"];
                var start = new ExcelCellAddress(4, 1);
                var end = new ExcelCellAddress(27, 5);
                for (int row = start.Row; row <= end.Row; row++)
                {
                    object s1 = null;
                    object s2 = null;
                    object s3 = null;
                    object s4 = null;
                    object s5 = null;
                    try
                    {
                        if (!((workSheet.Cells[row, 1].Text).Equals("...."))) s1 = workSheet.Cells[row, 1].Text;
                        if (!((workSheet.Cells[row, 2].Text).Equals("...."))) s2 = Double.Parse(workSheet.Cells[row, 2].Text);
                        if (!((workSheet.Cells[row, 3].Text).Equals("...."))) s3 = Double.Parse(workSheet.Cells[row, 3].Text);
                        if (!((workSheet.Cells[row, 4].Text).Equals("...."))) s4 = Double.Parse(workSheet.Cells[row, 4].Text);
                        if (!((workSheet.Cells[row, 5].Text).Equals("...."))) s5 = Double.Parse(workSheet.Cells[row, 5].Text);
                    }
                    catch (InvalidCastException e)
                    {
                        Console.WriteLine("Errore: " + e.ToString());
                    }

                    if (!alreadyExists(s1))
                    {
                        insertDB(s1, s2, s3, s4, s5);
                        Console.WriteLine("Record inserito");
                        n_inseriti++;
                    }
                }
            }
            return n_inseriti;
        }
    }
}
