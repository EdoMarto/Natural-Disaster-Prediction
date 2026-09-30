using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Data.SqlClient;
using System.Net;

namespace Thesis
{
    public class File17
    {
        public static void run()
        {
            Console.WriteLine("Sto avviando File17...");
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArguments("headless");
            IWebDriver driver = new ChromeDriver(Settings.chrome_driver_path, chromeOptions);
            driver.Navigate().GoToUrl("https://www.istat.it/it/archivio/263811");
            String element = driver.FindElement(By.XPath("/html/body/div[3]/div/div/div/div[3]/div[4]/div[2]/ul/li[2]/span[1]/a")).GetAttribute("href");
            int n_inseriti = manageExcel(element);
            driver.Close();
            Console.WriteLine("Esecuzione File17 terminata. Sono stati inseriti " + n_inseriti + " nuovi record");
        }

        static void insertDB(object tmp1, object tmp2, object tmp3, object tmp4, object tmp5, object tmp6, object tmp7, object tmp8, object tmp9, object tmp10, object tmp11, object tmp12, object tmp13, object tmp14, object tmp15, object tmp16, object tmp17, object tmp18, object tmp19, object tmp20, object tmp21, object tmp22, object tmp23, object tmp24, object tmp25, object tmp26, object tmp27, object tmp28, object tmp29, object tmp30, object tmp31, object tmp32, object tmp33, object tmp34, object tmp35, object tmp36, object tmp37, object tmp38, object tmp39, object tmp40, object tmp41, object tmp42, object tmp43, object tmp44, object tmp45)
        {
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            string query = "INSERT INTO tbl_13029_Stage_Istat_Estremi_Precipitazioni (DataImportazione, DataInserimento, DataUltimaModifica, Deleted, Comuni, ParentWeb_Id, GiorniSenzaPioggiaR0Anomalia2020dalValoreClimatico1971_2000, GiorniSenzaPioggiaR0Anomalia2006_2015dalValoreClimatico1971_2000, GiorniSenzaPioggiaR0ValoreClimatico1971_2000, GiorniSenzaPioggiaR0Differenza2020dalValoreMedio2006_2015, GiorniSenzaPioggiaR0ValoreMedio2006_2015, GiorniconPrecipitazionemag1mmR1Anomalia2020dalValoreClimatico1971_2000, GiorniconPrecipitazionemag1mmR1Anomalia2006_2015dalValoreClimatico1971_2000, GiorniconPrecipitazionemag1mmR1ValoreClimatico1971_2000, GiorniconPrecipitazionemag1mmR1Differenza2020dalValoremedio2006_2015, GiorniconPrecipitazionemag1mmR1ValoreMedio2006_2015, GiorniconPrecipitazionemag10mmR10Anomalia2020dalValoreClimatico1971_2000, GiorniconPrecipitazionemag10mmR10Anomalia2006_2015dalValoreClimatico1971_2000, GiorniconPrecipitazionemag10mmR10ValoreClimatico1971_2000, GiorniconPrecipitazionemag10mmR10Differenza2020dalValoreMedio2006_2015, GiorniconPrecipitazionemag10mmR10ValoreMedio2006_2015, GiorniconPrecipitazionemag20mmR20Anomalia2020dalValoreClimatico1971_2000, GiorniconPrecipitazionemag20mmR20Anomalia2006_2015dalValoreClimatico1971_2000, GiorniconPrecipitazionemag20mmR20ValoreClimatico1971_2000, GiorniconPrecipitazionemag20mmR20Differenza2020dalValoreMedio2006_2015, GiorniconPrecipitazionemag20mmR20ValoreMedio2006_2015, GiorniconPrecipitazionemag50mmR50Anomalia2020dalValoreClimatico1971_2000, GiorniconPrecipitazionemag50mmR50Anomalia2006_2015dalValoreClimatico1971_2000, GiorniconPrecipitazionemag50mmR50ValoreClimatico1971_2000, GiorniconPrecipitazionemag50mmR50Differenza2020dalValoreMedio2006_2015, GiorniconPrecipitazionemag50mmR50ValoreMedio2006_2015, GiorniConsecutiviconPioggiaCWDAnomalia2020dalValoreClimatico1971_2000, GiorniConsecutiviconPioggiaCWDAnomalia2006_2015dalValoreClimatico1971_2000, GiorniConsecutiviconPioggiaCWDValoreClimatico1971_2000, GiorniConsecutiviconPioggiaCWDDifferenza2020dalValoreMedio2006_2015, GiorniConsecutiviconPioggiaCWDValoreMedio2006_2015, GiorniConsecutiviSenzaPioggiaCDDAnomalia2020dalValoreClimatico1971_2000, GiorniConsecutiviSenzaPioggiaCDDAnomalia2006_2015dalValoreClimatico1971_2000, GiorniConsecutiviSenzaPioggiaCDDValoreClimatico1971_2000, GiorniConsecutiviSenzaPioggiaCDDDifferenza2020dalValoreMedio2006_2015, GiorniConsecutiviSenzaPioggiaCDDValoreMedio2006_2015, IntensitadiPioggiaGiornalieraSDIImmAnomalia2020dalValoreClimatico1971_2000, IntensitadiPioggiaGiornalieraSDIImmAnomalia2006_2015dalValoreClimatico1971_2000, IntensitadiPioggiaGiornalieraSDIImmValoreClimatico1971_2000, IntensitadiPioggiaGiornalieraSDIImmDifferenza2020dalValoreMedio2006_2015, IntensitadiPioggiaGiornalieraSDIImmValoreMedio2006_2015, PrecipitazioneneigiorniMoltoPiovosiR95PmmAnomalia2020dalValoreClimatico1971_2000, PrecipitazioneneigiorniMoltoPiovosiR95PmmAnomalia2006_2015dalValoreClimatico1971_2000, PrecipitazioneneigiorniMoltoPiovosiR95PmmPrecipitazioneneigiorniMoltoPiovosiR95Pmm, PrecipitazioneneigiorniMoltoPiovosiR95PmmValoreMedio2006_2015) " +
                "VALUES(@DataImportazione, @DataInserimento, @DataUltimaModifica, @Deleted, @Comuni, @ParentWeb_Id, @GiorniSenzaPioggiaR0Anomalia2020dalValoreClimatico1971_2000, @GiorniSenzaPioggiaR0Anomalia2006_2015dalValoreClimatico1971_2000, @GiorniSenzaPioggiaR0ValoreClimatico1971_2000, @GiorniSenzaPioggiaR0Differenza2020dalValoreMedio2006_2015, @GiorniSenzaPioggiaR0ValoreMedio2006_2015, @GiorniconPrecipitazionemag1mmR1Anomalia2020dalValoreClimatico1971_2000, @GiorniconPrecipitazionemag1mmR1Anomalia2006_2015dalValoreClimatico1971_2000, @GiorniconPrecipitazionemag1mmR1ValoreClimatico1971_2000, @GiorniconPrecipitazionemag1mmR1Differenza2020dalValoremedio2006_2015, @GiorniconPrecipitazionemag1mmR1ValoreMedio2006_2015, @GiorniconPrecipitazionemag10mmR10Anomalia2020dalValoreClimatico1971_2000, @GiorniconPrecipitazionemag10mmR10Anomalia2006_2015dalValoreClimatico1971_2000, @GiorniconPrecipitazionemag10mmR10ValoreClimatico1971_2000, @GiorniconPrecipitazionemag10mmR10Differenza2020dalValoreMedio2006_2015, @GiorniconPrecipitazionemag10mmR10ValoreMedio2006_2015, @GiorniconPrecipitazionemag20mmR20Anomalia2020dalValoreClimatico1971_2000, @GiorniconPrecipitazionemag20mmR20Anomalia2006_2015dalValoreClimatico1971_2000, @GiorniconPrecipitazionemag20mmR20ValoreClimatico1971_2000, @GiorniconPrecipitazionemag20mmR20Differenza2020dalValoreMedio2006_2015, @GiorniconPrecipitazionemag20mmR20ValoreMedio2006_2015, @GiorniconPrecipitazionemag50mmR50Anomalia2020dalValoreClimatico1971_2000, @GiorniconPrecipitazionemag50mmR50Anomalia2006_2015dalValoreClimatico1971_2000, @GiorniconPrecipitazionemag50mmR50ValoreClimatico1971_2000, @GiorniconPrecipitazionemag50mmR50Differenza2020dalValoreMedio2006_2015, @GiorniconPrecipitazionemag50mmR50ValoreMedio2006_2015, @GiorniConsecutiviconPioggiaCWDAnomalia2020dalValoreClimatico1971_2000, @GiorniConsecutiviconPioggiaCWDAnomalia2006_2015dalValoreClimatico1971_2000, @GiorniConsecutiviconPioggiaCWDValoreClimatico1971_2000, @GiorniConsecutiviconPioggiaCWDDifferenza2020dalValoreMedio2006_2015, @GiorniConsecutiviconPioggiaCWDValoreMedio2006_2015, @GiorniConsecutiviSenzaPioggiaCDDAnomalia2020dalValoreClimatico1971_2000, @GiorniConsecutiviSenzaPioggiaCDDAnomalia2006_2015dalValoreClimatico1971_2000, @GiorniConsecutiviSenzaPioggiaCDDValoreClimatico1971_2000, @GiorniConsecutiviSenzaPioggiaCDDDifferenza2020dalValoreMedio2006_2015, @GiorniConsecutiviSenzaPioggiaCDDValoreMedio2006_2015, @IntensitadiPioggiaGiornalieraSDIImmAnomalia2020dalValoreClimatico1971_2000, @IntensitadiPioggiaGiornalieraSDIImmAnomalia2006_2015dalValoreClimatico1971_2000, @IntensitadiPioggiaGiornalieraSDIImmValoreClimatico1971_2000, @IntensitadiPioggiaGiornalieraSDIImmDifferenza2020dalValoreMedio2006_2015, @IntensitadiPioggiaGiornalieraSDIImmValoreMedio2006_2015, @PrecipitazioneneigiorniMoltoPiovosiR95PmmAnomalia2020dalValoreClimatico1971_2000, @PrecipitazioneneigiorniMoltoPiovosiR95PmmAnomalia2006_2015dalValoreClimatico1971_2000, @PrecipitazioneneigiorniMoltoPiovosiR95PmmPrecipitazioneneigiorniMoltoPiovosiR95Pmm, @PrecipitazioneneigiorniMoltoPiovosiR95PmmValoreMedio2006_2015)";
            SqlCommand command = new SqlCommand(query, connection);

            if (tmp1 != null)
            {
                command.Parameters.AddWithValue("@Comuni", tmp1);
            }
            command.Parameters.AddWithValue("@DataImportazione", DateTime.Now);
            command.Parameters.AddWithValue("@DataInserimento", DateTime.Now);
            command.Parameters.AddWithValue("@DataUltimaModifica", DateTime.Now);
            command.Parameters.AddWithValue("@Deleted", false);
            command.Parameters.AddWithValue("@ParentWeb_Id", 13029);
            command.Parameters.AddWithValue("@GiorniSenzaPioggiaR0Anomalia2020dalValoreClimatico1971_2000", tmp2 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniSenzaPioggiaR0Anomalia2006_2015dalValoreClimatico1971_2000", tmp3 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniSenzaPioggiaR0ValoreClimatico1971_2000", tmp4 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniSenzaPioggiaR0Differenza2020dalValoreMedio2006_2015", tmp5 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniSenzaPioggiaR0ValoreMedio2006_2015", tmp6 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniconPrecipitazionemag1mmR1Anomalia2020dalValoreClimatico1971_2000", tmp7 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniconPrecipitazionemag1mmR1Anomalia2006_2015dalValoreClimatico1971_2000", tmp8 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniconPrecipitazionemag1mmR1ValoreClimatico1971_2000", tmp9 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniconPrecipitazionemag1mmR1Differenza2020dalValoremedio2006_2015", tmp10 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniconPrecipitazionemag1mmR1ValoreMedio2006_2015", tmp11 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniconPrecipitazionemag10mmR10Anomalia2020dalValoreClimatico1971_2000", tmp12 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniconPrecipitazionemag10mmR10Anomalia2006_2015dalValoreClimatico1971_2000", tmp13 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniconPrecipitazionemag10mmR10ValoreClimatico1971_2000", tmp14 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniconPrecipitazionemag10mmR10Differenza2020dalValoreMedio2006_2015", tmp15 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniconPrecipitazionemag10mmR10ValoreMedio2006_2015", tmp16 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniconPrecipitazionemag20mmR20Anomalia2020dalValoreClimatico1971_2000", tmp17 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniconPrecipitazionemag20mmR20Anomalia2006_2015dalValoreClimatico1971_2000", tmp18 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniconPrecipitazionemag20mmR20ValoreClimatico1971_2000", tmp19 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniconPrecipitazionemag20mmR20Differenza2020dalValoreMedio2006_2015", tmp20 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniconPrecipitazionemag20mmR20ValoreMedio2006_2015", tmp21 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniconPrecipitazionemag50mmR50Anomalia2020dalValoreClimatico1971_2000", tmp22 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniconPrecipitazionemag50mmR50Anomalia2006_2015dalValoreClimatico1971_2000", tmp23 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniconPrecipitazionemag50mmR50ValoreClimatico1971_2000", tmp24 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniconPrecipitazionemag50mmR50Differenza2020dalValoreMedio2006_2015", tmp25 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniconPrecipitazionemag50mmR50ValoreMedio2006_2015", tmp26 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniConsecutiviconPioggiaCWDAnomalia2020dalValoreClimatico1971_2000", tmp27 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniConsecutiviconPioggiaCWDAnomalia2006_2015dalValoreClimatico1971_2000", tmp28 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniConsecutiviconPioggiaCWDValoreClimatico1971_2000", tmp29 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniConsecutiviconPioggiaCWDDifferenza2020dalValoreMedio2006_2015", tmp30 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniConsecutiviconPioggiaCWDValoreMedio2006_2015", tmp31 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniConsecutiviSenzaPioggiaCDDAnomalia2020dalValoreClimatico1971_2000", tmp32 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniConsecutiviSenzaPioggiaCDDAnomalia2006_2015dalValoreClimatico1971_2000", tmp33 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniConsecutiviSenzaPioggiaCDDValoreClimatico1971_2000", tmp34 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniConsecutiviSenzaPioggiaCDDDifferenza2020dalValoreMedio2006_2015", tmp35 ?? DBNull.Value);
            command.Parameters.AddWithValue("@GiorniConsecutiviSenzaPioggiaCDDValoreMedio2006_2015", tmp36 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IntensitadiPioggiaGiornalieraSDIImmAnomalia2020dalValoreClimatico1971_2000", tmp37 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IntensitadiPioggiaGiornalieraSDIImmAnomalia2006_2015dalValoreClimatico1971_2000", tmp38 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IntensitadiPioggiaGiornalieraSDIImmValoreClimatico1971_2000", tmp39 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IntensitadiPioggiaGiornalieraSDIImmDifferenza2020dalValoreMedio2006_2015", tmp40 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IntensitadiPioggiaGiornalieraSDIImmValoreMedio2006_2015", tmp41 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PrecipitazioneneigiorniMoltoPiovosiR95PmmAnomalia2020dalValoreClimatico1971_2000", tmp42 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PrecipitazioneneigiorniMoltoPiovosiR95PmmAnomalia2006_2015dalValoreClimatico1971_2000", tmp43 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PrecipitazioneneigiorniMoltoPiovosiR95PmmPrecipitazioneneigiorniMoltoPiovosiR95Pmm", tmp44 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PrecipitazioneneigiorniMoltoPiovosiR95PmmValoreMedio2006_2015", tmp45 ?? DBNull.Value);

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
            SqlCommand command = new SqlCommand("SELECT COUNT(*) FROM tbl_13029_Stage_Istat_Estremi_Precipitazioni WHERE Comuni=@comune", connection);
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
            Console.WriteLine(filename);
            using (WebClient wc = new WebClient())
            {
                wc.DownloadFile(
                    new System.Uri(url),
                    "../../../Excels/Automatici/" + filename
                );
            }

            using (var pck = new ExcelPackage(new FileInfo("../../../Excels/Automatici/" + filename)))
            {
                ExcelWorksheet workSheet = pck.Workbook.Worksheets["Tavola_5"];
                var start = new ExcelCellAddress(5, 1);
                var end = new ExcelCellAddress(28, 45);
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
                    try
                    {
                        if (!((workSheet.Cells[row, 1].Text).Equals("…."))) s1 = workSheet.Cells[row, 1].Text.Replace("-+", "");
                        if (!((workSheet.Cells[row, 2].Text).Equals("…."))) s2 = Double.Parse(workSheet.Cells[row, 2].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 3].Text).Equals("…."))) s3 = Double.Parse(workSheet.Cells[row, 3].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 4].Text).Equals("…."))) s4 = Double.Parse(workSheet.Cells[row, 4].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 5].Text).Equals("…."))) s5 = Double.Parse(workSheet.Cells[row, 5].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 6].Text).Equals("…."))) s6 = Double.Parse(workSheet.Cells[row, 6].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 7].Text).Equals("…."))) s7 = Double.Parse(workSheet.Cells[row, 7].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 8].Text).Equals("…."))) s8 = Double.Parse(workSheet.Cells[row, 8].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 9].Text).Equals("…."))) s9 = Double.Parse(workSheet.Cells[row, 9].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 10].Text).Equals("…."))) s10 = Double.Parse(workSheet.Cells[row, 10].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 11].Text).Equals("…."))) s11 = Double.Parse(workSheet.Cells[row, 11].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 12].Text).Equals("…."))) s12 = Double.Parse(workSheet.Cells[row, 12].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 13].Text).Equals("…."))) s13 = Double.Parse(workSheet.Cells[row, 13].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 14].Text).Equals("…."))) s14 = Double.Parse(workSheet.Cells[row, 14].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 15].Text).Equals("…."))) s15 = Double.Parse(workSheet.Cells[row, 15].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 16].Text).Equals("…."))) s16 = Double.Parse(workSheet.Cells[row, 16].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 17].Text).Equals("…."))) s17 = Double.Parse(workSheet.Cells[row, 17].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 18].Text).Equals("…."))) s18 = Double.Parse(workSheet.Cells[row, 18].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 19].Text).Equals("…."))) s19 = Double.Parse(workSheet.Cells[row, 19].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 20].Text).Equals("…."))) s20 = Double.Parse(workSheet.Cells[row, 20].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 21].Text).Equals("…."))) s21 = Double.Parse(workSheet.Cells[row, 21].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 22].Text).Equals("…."))) s22 = Double.Parse(workSheet.Cells[row, 22].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 23].Text).Equals("…."))) s23 = Double.Parse(workSheet.Cells[row, 23].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 24].Text).Equals("…."))) s24 = Double.Parse(workSheet.Cells[row, 24].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 25].Text).Equals("…."))) s25 = Double.Parse(workSheet.Cells[row, 25].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 26].Text).Equals("…."))) s26 = Double.Parse(workSheet.Cells[row, 26].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 27].Text).Equals("…."))) s27 = Double.Parse(workSheet.Cells[row, 27].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 28].Text).Equals("…."))) s28 = Double.Parse(workSheet.Cells[row, 28].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 29].Text).Equals("…."))) s29 = Double.Parse(workSheet.Cells[row, 29].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 30].Text).Equals("…."))) s30 = Double.Parse(workSheet.Cells[row, 30].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 31].Text).Equals("…."))) s31 = Double.Parse(workSheet.Cells[row, 31].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 32].Text).Equals("…."))) s32 = Double.Parse(workSheet.Cells[row, 32].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 33].Text).Equals("…."))) s33 = Double.Parse(workSheet.Cells[row, 33].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 34].Text).Equals("…."))) s34 = Double.Parse(workSheet.Cells[row, 34].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 35].Text).Equals("…."))) s35 = Double.Parse(workSheet.Cells[row, 35].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 36].Text).Equals("…."))) s36 = Double.Parse(workSheet.Cells[row, 36].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 37].Text).Equals("…."))) s37 = Double.Parse(workSheet.Cells[row, 37].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 38].Text).Equals("…."))) s38 = Double.Parse(workSheet.Cells[row, 38].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 39].Text).Equals("…."))) s39 = Double.Parse(workSheet.Cells[row, 39].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 40].Text).Equals("…."))) s40 = Double.Parse(workSheet.Cells[row, 40].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 41].Text).Equals("…."))) s41 = Double.Parse(workSheet.Cells[row, 41].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 42].Text).Equals("…."))) s42 = Double.Parse(workSheet.Cells[row, 42].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 43].Text).Equals("…."))) s43 = Double.Parse(workSheet.Cells[row, 43].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 44].Text).Equals("…."))) s44 = Double.Parse(workSheet.Cells[row, 44].Text.Replace("-+", ""));
                        if (!((workSheet.Cells[row, 45].Text).Equals("…."))) s45 = Double.Parse(workSheet.Cells[row, 45].Text.Replace("-+", ""));
                    }
                    catch (InvalidCastException e)
                    {
                        Console.WriteLine("Errore: " + e.ToString());
                    }

                    if (!alreadyExists(s1))
                    {
                        insertDB(s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14, s15, s16, s17, s18, s19, s20, s21, s22, s23, s24, s25, s26, s27, s28, s29, s30, s31, s32, s33, s34, s35, s36, s37, s38, s39, s40, s41, s42, s43, s44, s45);
                        Console.WriteLine("Record inserito");
                        n_inseriti++;
                    }
                }
            }
            return n_inseriti;
        }
    }
}
