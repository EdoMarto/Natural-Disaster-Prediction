namespace Thesis
{
    public class Settings
    {
        // Folder containing chromedriver.exe (download it from https://chromedriver.chromium.org)
        public static String chrome_driver_path =
            Environment.GetEnvironmentVariable("CHROMEDRIVER_PATH") ?? Path.GetFullPath("../../../Settings/");

        // SQL Server connection string, e.g. "Data Source=localhost;Initial Catalog=ClimateGeckoDB;Integrated Security=True;"
        public static String connection_string =
            Environment.GetEnvironmentVariable("CLIMATEGECKO_DB_CONNECTION")
            ?? "Data Source=localhost;Initial Catalog=ClimateGeckoDB;Integrated Security=True;";
    }
}
