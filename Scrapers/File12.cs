using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Data.SqlClient;
using System.IO.Compression;
using System.Net;

namespace Thesis
{
    public class File12
    {
        public static void run()
        {
            Console.WriteLine("Sto avviando File12...");
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArguments("headless");
            IWebDriver driver = new ChromeDriver(Settings.chrome_driver_path, chromeOptions);
            driver.Navigate().GoToUrl("https://annuario.isprambiente.it/sys_ind/macro/11");
            String element = driver.FindElement(By.XPath("/html/body/div[2]/div/section/div/section/div/div/div/div[1]/ul/li[8]/div/div/div/div[3]/div/div/div/a[4]")).GetAttribute("href");
            int n_inseriti = manageExcel(element);
            driver.Close();
            Console.WriteLine("Esecuzione File12 terminata. Sono stati inseriti " + n_inseriti + " nuovi record");
        }

        static void insertDB(object tmp1, object tmp2, object tmp3, object tmp4, object tmp5, object tmp6, object tmp7, object tmp8, object tmp9, object tmp10, object tmp11, object tmp12, object tmp13, object tmp14)
        {
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            string query = "INSERT INTO tbl_13029_Stage_Ispra_Eventi_Franosi (DataInserimento, DataUltimaModifica, Deleted, DataImportazione, Id_Evento, Data, Localita, Comune, Provincia, Regione, Descrizione, PersoneEvacuate, Feriti, MortieDispersi, DanniEdificiBeniCulturaliPaesaggistici, DanniInfrastrutturediComunicazionePrimarie, DanniInfrastruttureeRetidiServizi, Ordinanze, ParentWeb_Id) " +
                "VALUES(@DataInserimento, @DataUltimaModifica, @Deleted, @DataImportazione, @Id_Evento, @Data, @Localita, @Comune, @Provincia, @Regione, @Descrizione, @PersoneEvacuate, @Feriti, @MortieDispersi, @DanniEdificiBeniCulturaliPaesaggistici, @DanniInfrastrutturediComunicazionePrimarie, @DanniInfrastruttureeRetidiServizi, @Ordinanze, @ParentWeb_Id)";
            SqlCommand command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@DataInserimento", DateTime.Now);
            command.Parameters.AddWithValue("@DataUltimaModifica", DateTime.Now);
            command.Parameters.AddWithValue("@Deleted", false);
            command.Parameters.AddWithValue("@DataImportazione", DateTime.Now);
            command.Parameters.AddWithValue("@Id_Evento", tmp1);
            command.Parameters.AddWithValue("@Data", DateTime.Parse((string)tmp2));
            command.Parameters.AddWithValue("@Localita", tmp3 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Comune", tmp4 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Provincia", tmp5 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Regione", tmp6 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Descrizione", tmp7 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PersoneEvacuate", tmp8 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Feriti", tmp9 ?? DBNull.Value);
            command.Parameters.AddWithValue("@MortieDispersi", tmp10 ?? DBNull.Value);
            command.Parameters.AddWithValue("@DanniEdificiBeniCulturaliPaesaggistici", tmp11 ?? DBNull.Value);
            command.Parameters.AddWithValue("@DanniInfrastrutturediComunicazionePrimarie", tmp12 ?? DBNull.Value);
            command.Parameters.AddWithValue("@DanniInfrastruttureeRetidiServizi", tmp13 ?? DBNull.Value);
            command.Parameters.AddWithValue("@Ordinanze", tmp14 ?? DBNull.Value);
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
            SqlCommand command = new SqlCommand("SELECT COUNT(*) FROM tbl_13029_Stage_Ispra_Eventi_Franosi WHERE Id_Evento=@pid", connection);
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
                    if (entry.Name == "Tab1_Eventi_franosi_principali_2020_ADA_ed2021.xlsx")
                        entry.ExtractToFile("../../../Excels/Automatici/Tab1_Eventi_franosi_principali_2020_ADA_ed2021.xlsx", true);

            using (var pck = new ExcelPackage(new FileInfo("../../../Excels/Automatici/Tab1_Eventi_franosi_principali_2020_ADA_ed2021.xlsx")))
            {
                ExcelWorksheet workSheet = pck.Workbook.Worksheets["Eventi_franosi_2020"];
                var start = new ExcelCellAddress(2, 1);
                var end = new ExcelCellAddress(123, 14);
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
                    try
                    {
                        s1 = workSheet.Cells[row, 1].Text;
                        if (!((workSheet.Cells[row, 2].Text).Equals("-"))) s2 = workSheet.Cells[row, 2].Text;
                        if (!((workSheet.Cells[row, 3].Text).Equals(""))) s3 = workSheet.Cells[row, 3].Text;
                        if (!((workSheet.Cells[row, 4].Text).Equals("-"))) s4 = workSheet.Cells[row, 4].Text;
                        if (!((workSheet.Cells[row, 5].Text).Equals("-"))) s5 = workSheet.Cells[row, 5].Text;
                        if (!((workSheet.Cells[row, 6].Text).Equals("-"))) s6 = workSheet.Cells[row, 6].Text;
                        if (!((workSheet.Cells[row, 7].Text).Equals("-"))) s7 = workSheet.Cells[row, 7].Text;
                        if (!((workSheet.Cells[row, 8].Text).Equals("-"))) s8 = workSheet.Cells[row, 8].Text;
                        if (!((workSheet.Cells[row, 9].Text).Equals("-"))) s9 = workSheet.Cells[row, 9].Text;
                        if (!((workSheet.Cells[row, 10].Text).Equals("-"))) s10 = workSheet.Cells[row, 10].Text;
                        if (!((workSheet.Cells[row, 11].Text).Equals("-"))) s11 = workSheet.Cells[row, 11].Text;
                        if (!((workSheet.Cells[row, 12].Text).Equals("-"))) s12 = workSheet.Cells[row, 12].Text;
                        if (!((workSheet.Cells[row, 13].Text).Equals("-"))) s13 = workSheet.Cells[row, 13].Text;
                        if (!((workSheet.Cells[row, 14].Text).Equals(""))) s14 = workSheet.Cells[row, 14].Text;
                    }
                    catch (InvalidCastException e)
                    {
                        Console.WriteLine("Errore: " + e.ToString());
                    }

                    if (!alreadyExists(s1))
                    {
                        insertDB(s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14);
                        Console.WriteLine("Record inserito");
                        n_inseriti++;
                    }
                }
            }
            return n_inseriti;
        }
    }
}
