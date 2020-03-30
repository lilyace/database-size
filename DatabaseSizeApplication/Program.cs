using Npgsql;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace DatabaseSizeApplication
{
    class Program
    {
        const int kilobytesInGigabyte = 1073741824;
        static void Main(string[] args)
        {
            Timer timer = new Timer(new TimerCallback(UpdateSpreadsheet), null, 2000, 90000);
            Console.WriteLine("Для остановки автообновления нажмите q");
            ConsoleKeyInfo key;
            do
            {
                key = Console.ReadKey();
            } while (key.KeyChar != 'q');
            timer.Dispose();
            Console.WriteLine("\nАвтообновление файла остановлено");
        }

        private static void UpdateSpreadsheet(object source)
        {
            var spreadsheetId = ConfigurationManager.AppSettings["spreadsheetId"];
            if (spreadsheetId == null)
            {
                spreadsheetId = GoogleSpreadsheetsHelper.CreateSpreadsheet();
                var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var settings = config.AppSettings.Settings;
                settings.Add("spreadsheetId", spreadsheetId);
                config.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection(config.AppSettings.SectionInformation.Name);
                Console.WriteLine("Таблица успешно создана");
            }
            else GoogleSpreadsheetsHelper.SetSpreadsheetId(spreadsheetId);

            var sheetsName = GoogleSpreadsheetsHelper.GetSheetsNames();
            var servers = ConfigurationManager.ConnectionStrings;

            foreach (ConnectionStringSettings server in servers)
            {
                if (server.ProviderName != "Npgsql")
                    continue;
                if (!sheetsName.Contains(server.Name))
                    GoogleSpreadsheetsHelper.CreateSheet(server.Name);
                int linesCount = GoogleSpreadsheetsHelper.GetLinesCount(server.Name);
                using (var connection = new NpgsqlConnection(server.ConnectionString))
                {
                    connection.Open();
                    var getAllDatabaseCommand = new NpgsqlCommand("SELECT datname, pg_database_size(datname) from pg_database WHERE datistemplate=false;", connection);
                    var reader = getAllDatabaseCommand.ExecuteReader();
                    var dataForWriting = new List<IList<object>>();
                    while (reader.Read())
                    {
                        dataForWriting.Add(new List<object>() { server.Name, reader[0], string.Format("{0:0.##}", Convert.ToDouble(reader[1]) / kilobytesInGigabyte), DateTime.Now.ToShortDateString() });
                    }
                    GoogleSpreadsheetsHelper.WriteNewData(dataForWriting, linesCount, server.Name);
                }
            }
        }
    }
}
