using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace DatabaseSizeApplication
{
    public static class GoogleSpreadsheetsHelper
    {
        private static SheetsService _service;
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        private static string _spreadsheetId;
        static GoogleSpreadsheetsHelper()
        {
            UserCredential credential;

            var googleAccountConnectionParameters = ConfigurationManager.AppSettings["googleAccountConnectionParameters"];
            using (var stream =
                new MemoryStream(Encoding.UTF8.GetBytes(googleAccountConnectionParameters)))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            _service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "appname",
            });
        }

        public static string CreateSpreadsheet()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.Properties = new SpreadsheetProperties();
            spreadsheet.Properties.Title = "Databases information";
            try
            {
                var spreadsheetWithId = _service.Spreadsheets.Create(spreadsheet).Execute();
                _spreadsheetId = spreadsheetWithId.SpreadsheetId;
                return _spreadsheetId;
            }
            catch (Exception ex)
            {
                throw new Exception($"Произошла непредвиденная ошибка: {ex.Message}");
            }
        }

        public static void SetSpreadsheetId(string spreadsheetId)
        {
            _spreadsheetId = spreadsheetId;
        }

        public static IEnumerable<string> GetSheetsNames()
        {
            return _service.Spreadsheets.Get(_spreadsheetId).Execute().Sheets.Select(x => x.Properties.Title);
        }

        public static void WriteNewData(List<IList<object>> dataForWriting, int linesCount, string listName)
        {
            if (linesCount == 1)
                linesCount++;
            var driveSizeString = ((NameValueCollection)ConfigurationManager.GetSection("driveSize"))[$"{listName}Size"];
            int driveSize = Convert.ToInt32(driveSizeString);
            double totalDbSize = dataForWriting.Select(x => Convert.ToDouble(x[2])).Sum();
            dataForWriting.ForEach(x =>  x[2] = x[2].ToString().Replace(",","."));
            
            dataForWriting.Add(new List<object>() { listName, "Свободно", (driveSize - totalDbSize).ToString("0.##", CultureInfo.CreateSpecificCulture("en-US")), DateTime.Now.ToShortDateString() });
            var range = $"{listName}!A{linesCount}:D{linesCount + dataForWriting.Count}";
            var valueRange = new ValueRange() { Values = dataForWriting };
            var updateRequest = _service.Spreadsheets.Values.Update(valueRange, _spreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            updateRequest.Execute();
        }

        public static void CreateSheet(string sheetName)
        {
            var addSheetRequest = new AddSheetRequest();
            addSheetRequest.Properties = new SheetProperties();
            addSheetRequest.Properties.Title = sheetName;

            var updateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>()
            };
            updateSpreadsheetRequest.Requests.Add(
                new Request
                {
                    AddSheet = addSheetRequest,
                });

            _service.Spreadsheets.BatchUpdate(updateSpreadsheetRequest, _spreadsheetId).Execute();

            var valueRange = new ValueRange() { Values = new List<IList<object>>() { new List<object>() { "Сервер", "База данных", "Размер в ГБ", "Дата обновления" } } };
            var updateRequest = _service.Spreadsheets.Values.Update(valueRange, _spreadsheetId, $"{sheetName}!A1:4");
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            updateRequest.Execute();
        }

        public static int GetLinesCount(string sheetName)
        {
            string range = $"{sheetName}!A1:D";
            var lineWithValues = _service.Spreadsheets.Values.Get(_spreadsheetId, range).Execute().Values;
            return lineWithValues.Count;
        }
    }
}
