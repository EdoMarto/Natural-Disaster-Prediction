using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System.Data.SqlClient;

namespace Thesis
{
    class File2
    {
        public static void run()
        {
            Console.WriteLine("Sto avviando File2...");
            // maltempo, meteo, frana, terremoto, grandine, alluvione, magnitudo
            int n_inseriti = getArticles("maltempo");
            Console.WriteLine("Esecuzione File2 terminata. Sono stati inseriti " + n_inseriti + " nuovi record");
        }

        public static void insertArticle(string link, string html, string title, string subtitle, string description, string luogo, string data, string tipoevento)
        {
            string connectionString = Settings.connection_string;

            using (SqlConnection connection = new SqlConnection(@connectionString))
            {
                SqlCommand command = new SqlCommand(
                    "INSERT INTO tbl_13029_STAGE_Articoli (Link, DataInserimento, DataUltimaModifica, Deleted, HTML, Titolo, Descrizione, DataImportazione, Sorgente_Id, ParentWeb_Id, Sottotitolo, Comune, DataArticolo, TipoEvento) " +
                    "VALUES (@Link, @DataInserimento, @DataUltimaModifica, @Deleted, @HTML, @Titolo, @Descrizione, @DataImportazione, @Sorgente_Id, @ParentWeb_Id, @Sottotitolo, @Comune, @DataArticolo, @TipoEvento)", connection);

                command.Parameters.AddWithValue("@Link", link);
                command.Parameters.AddWithValue("@DataInserimento", System.DateTime.Now);
                command.Parameters.AddWithValue("@DataUltimaModifica", System.DateTime.Now);
                command.Parameters.AddWithValue("@Deleted", false);
                command.Parameters.AddWithValue("@HTML", html);
                command.Parameters.AddWithValue("@Titolo", title);
                command.Parameters.AddWithValue("@Descrizione", description);
                command.Parameters.AddWithValue("@DataImportazione", System.DateTime.Now);
                command.Parameters.AddWithValue("@Sorgente_Id", 1);
                command.Parameters.AddWithValue("@ParentWeb_Id", 13029);
                command.Parameters.AddWithValue("@Sottotitolo", subtitle);
                command.Parameters.AddWithValue("@Comune", luogo);
                command.Parameters.AddWithValue("@DataArticolo", DateTime.Parse(data));
                command.Parameters.AddWithValue("@TipoEvento", tipoevento);
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

        public static Boolean existingArticle(string article_link, string titolo, string sottotitolo)
        {
            string connectionString = Settings.connection_string;

            int count = 0;
            using (SqlConnection connection = new SqlConnection(@connectionString))
            {
                SqlCommand command = new SqlCommand("SELECT COUNT(*) FROM tbl_13029_STAGE_Articoli WHERE Link=@link OR Titolo=@titolo OR Sottotitolo=@sottotitolo", connection);

                command.Parameters.AddWithValue("@link", article_link);
                command.Parameters.AddWithValue("@titolo", titolo);
                command.Parameters.AddWithValue("@sottotitolo", sottotitolo);
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

        public static int getArticles(string keyword)
        {
            string link = "https://www.ansa.it/ricerca/ansait/search.shtml?any=" + keyword;
            int n_inseriti = 0;

            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArguments("headless");
            IWebDriver driver = new ChromeDriver(Settings.chrome_driver_path, chromeOptions);
            IWebDriver driver2 = new ChromeDriver(Settings.chrome_driver_path, chromeOptions);
            HtmlDocument htmlSnippet = new HtmlDocument();

            while (true)
            {
                try
                {
                    driver.Navigate().GoToUrl(link);
                    htmlSnippet.LoadHtml(driver.PageSource);
                    HtmlNodeCollection asd = htmlSnippet.DocumentNode.SelectNodes("//h3[@class='search-title']//a[@href]");
                    break;
                }
                catch (Exception e)
                {
                    Thread.Sleep(5000);
                }
            }

            IWebElement periodo = driver.FindElement(By.Name("periodo"));
            SelectElement selectElement = new SelectElement(periodo);
            selectElement.SelectByText("Da sempre");

            int rowCount = driver.PageSource.Split("facet_iptc_domain_label_it_").Length - 1;

            Console.WriteLine("CONTEGGIO CATEGORIE: " + rowCount);

            for (int i = 1; i <= rowCount; i++)
            {
                Console.WriteLine("NUMERO ATTUALE: " + i);
                Boolean done = false;

                IWebElement l = driver.FindElement(By.Id("facet_iptc_domain_label_it_" + i));

                Actions a = new Actions(driver);
                a.MoveToElement(l).Perform();
                string categoria = l.Text;

                Console.WriteLine("Analizzando categoria: " + categoria);
                IJavaScriptExecutor executor = (IJavaScriptExecutor)driver;
                executor.ExecuteScript("arguments[0].click();", l);
                HtmlNodeCollection asd;

                while (!done)
                {
                    while(true)
                    {
                        htmlSnippet.LoadHtml(driver.PageSource);
                        asd = htmlSnippet.DocumentNode.SelectNodes("//h3[@class='search-title']//a[@href]");
                        if(asd != null)
                        {
                            break;
                        }
                        else
                        {
                            Thread.Sleep(5000);
                            Console.WriteLine("\n\n\n\nSONO NELL ERRORE\n\n\n\n");
                            driver.Navigate().Refresh();
                        }
                    }
                    foreach (HtmlNode cc in asd)
                    {
                        HtmlAttribute att = cc.Attributes["href"];
                        string article_link = "https://www.ansa.it" + att.Value;
                        driver2.Navigate().GoToUrl(article_link);
                        String html_page = driver2.PageSource;
                        HtmlDocument html_snippet = new HtmlDocument();
                        html_snippet.LoadHtml(html_page);

                        HtmlNode? node_luogo = html_snippet.DocumentNode.SelectSingleNode("//span[@class='location']");
                        string? luogo = node_luogo?.InnerText;
                        HtmlNode? node_data = html_snippet.DocumentNode.SelectSingleNode("//time");
                        string? data = node_data?.InnerText;
                        HtmlNode? node_titolo = html_snippet.DocumentNode.SelectSingleNode("//h1[@class='news-title']");
                        string? titolo = node_titolo?.InnerText;
                        HtmlNode? node_sottotitolo = html_snippet.DocumentNode.SelectSingleNode("//h2[@class='news-stit']");
                        string? sottotitolo = node_sottotitolo?.InnerText;
                        HtmlNode? node_desc = html_snippet.DocumentNode.SelectSingleNode("//div[@class='news-txt']");
                        string? desc = node_desc?.InnerText;

                        if (titolo != null && sottotitolo != null && desc != null && luogo != null && data != null && !titolo.Equals("") && !sottotitolo.Equals("") && !desc.Equals("") && !categoria.Equals("") && !luogo.Equals("") && !data.Equals(""))
                        {
                            if (!existingArticle(article_link, titolo, sottotitolo))
                            {
                                desc = desc.Replace("&nbsp;", "");
                                desc = desc.Replace("  ", "");
                                data = data.Substring(0, data.IndexOf(":") - 2) + " " + data.Substring(data.IndexOf(":") - 2);
                                insertArticle(article_link, html_page, titolo, sottotitolo, desc, luogo, data, categoria);
                                n_inseriti++;
                            }
                        }
                    }

                    try
                    {
                        IWebElement el_page = driver.FindElement(By.ClassName("next"));
                        string next_page_function = el_page.FindElement(By.TagName("a")).GetAttribute("onclick");

                        IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                        string fun = next_page_function.Substring(next_page_function.IndexOf(" ") + 1) + ";";
                        js.ExecuteScript(fun);
                    }
                    catch (NoSuchElementException ex)
                    {
                        Console.WriteLine(ex.Message);
                        Console.WriteLine("FINITE PAGINE DI QUESTA CATEGORIA");
                        IWebElement lx = driver.FindElement(By.Name("rmv_facet_filter"));
                        IJavaScriptExecutor categ = (IJavaScriptExecutor)driver;
                        categ.ExecuteScript("arguments[0].click();", lx);
                        done = true;
                    }
                }
            }
            driver.Close();
            driver2.Close();
            return n_inseriti;
        }
    }
}