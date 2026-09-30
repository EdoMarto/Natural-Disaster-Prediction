using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Collections.ObjectModel;
using System.Data.SqlClient;

namespace Thesis
{
    class File9
    {
        public static void run()
        {
            Console.WriteLine("Sto avviando File9...");
            int n_inseriti = 0;
            var chromeOptions = new ChromeOptions();
            //chromeOptions.AddArguments("headless");
            IWebDriver driver = new ChromeDriver(Settings.chrome_driver_path, chromeOptions);

            string url = "http://terremoti.ingv.it/events?starttime=1985-01-01%2B00%253A00%253A00&endtime=2022-04-05%2B23%253A59%253A59&last_nd=-2&minmag=2&maxmag=10&mindepth=-10&maxdepth=1000&minlat=35&maxlat=49&minlon=5&maxlon=20&minversion=100&limit=30&orderby=ot-desc&lat=0&lon=0&maxradiuskm=-1&wheretype=area&box_search=Italia&page=1";

            driver.Navigate().GoToUrl(url);

            ReadOnlyCollection<IWebElement> page_tags = driver.FindElements(By.ClassName("page-item"));
            IWebElement lastElement = page_tags.ElementAt(page_tags.Count - 2);
            int number_of_pages = Int32.Parse(lastElement.Text);

            for (int i = 1; i <= number_of_pages; i++)
            {
                url = "http://terremoti.ingv.it/events?starttime=1985-01-01%2B00%253A00%253A00&endtime=2022-04-05%2B23%253A59%253A59&last_nd=-2&minmag=2&maxmag=10&mindepth=-10&maxdepth=1000&minlat=35&maxlat=49&minlon=5&maxlon=20&minversion=100&limit=30&orderby=ot-desc&lat=0&lon=0&maxradiuskm=-1&wheretype=area&box_search=Italia&page=" + i;
                while (true)
                {
                    driver.Navigate().GoToUrl(url);
                    if (!driver.PageSource.Contains("Server Error"))
                    {
                        break;
                    }
                }
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                IWebElement table = wait.Until(driver => driver.FindElement(By.TagName("tbody")));
                var elements = driver.FindElements(By.XPath("/html/body/div[1]/div[2]/div[6]/div/div/div/table/tbody/tr"));
                foreach (IWebElement e in elements)
                {
                    ReadOnlyCollection<IWebElement> infos = e.FindElements(By.TagName("td"));
                    DateTime dataeora = DateTime.Parse(infos.ElementAt(0).Text);
                    double magnitudo = double.Parse(infos.ElementAt(1).Text.Substring(infos.ElementAt(1).Text.IndexOf(" ") + 1).Replace(".", ","));
                    String zona = infos.ElementAt(2).Text;
                    double profondita = double.Parse(infos.ElementAt(3).Text);
                    double latitudine = double.Parse(infos.ElementAt(4).Text.Replace(".", ","));
                    double longitudine = double.Parse(infos.ElementAt(5).Text.Replace(".", ","));
                    if (!existingInfo(dataeora))
                    {
                        insertInfo(dataeora, magnitudo, zona, profondita, latitudine, longitudine);
                        Console.WriteLine("Record inserito");
                        n_inseriti++;
                    }
                }
            }
            driver.Close();
            driver.Quit();
            Console.WriteLine("Esecuzione File9 terminata. Sono stati inseriti " + n_inseriti + " nuovi record");

        }

        public static void insertInfo(DateTime dataeora, double magnitudo, String zona, double profondita, double latitudine, double longitudine)
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(Settings.connection_string);

            using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
            {
                SqlCommand command = new SqlCommand(
                    "INSERT INTO tbl_13029_STAGE_INGV_Sismici (DataImportazione, DataInserimento, DataUltimaModifica, Deleted, DataeOra, Magnitudo, Zona, ProfonditaKm, Latitudine, Longitudine, ParentWeb_Id) " +
                    "VALUES (@DataImportazione, @DataInserimento, @DataUltimaModifica, @Deleted, @DataeOra, @Magnitudo, @Zona, @ProfonditaKm, @Latitudine, @Longitudine, @ParentWeb_Id)", connection);

                command.Parameters.AddWithValue("@DataImportazione", System.DateTime.Now);
                command.Parameters.AddWithValue("@DataInserimento", System.DateTime.Now);
                command.Parameters.AddWithValue("@DataUltimaModifica", System.DateTime.Now);
                command.Parameters.AddWithValue("@Deleted", false);
                command.Parameters.AddWithValue("@DataeOra", dataeora);
                command.Parameters.AddWithValue("@Magnitudo", magnitudo);
                command.Parameters.AddWithValue("@Zona", zona);
                command.Parameters.AddWithValue("@ProfonditaKm", profondita);
                command.Parameters.AddWithValue("@Latitudine", latitudine);
                command.Parameters.AddWithValue("@Longitudine", longitudine);
                command.Parameters.AddWithValue("@ParentWeb_Id", 13029);
                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (SqlException e)
                {
                    Console.WriteLine(e.ToString());
                }
                finally
                {
                    connection.Close();
                }
            }
        }

        public static Boolean existingInfo(DateTime dataeora)
        {
            int count = 0;
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(Settings.connection_string);


            using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
            {
                SqlCommand command = new SqlCommand("SELECT COUNT(*) FROM tbl_13029_Stage_INGV_Sismici WHERE DataeOra=@dataeora", connection);

                command.Parameters.AddWithValue("@dataeora", dataeora);
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
            }
            return count != 0;
        }
    }
}