using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;

namespace GitHubActionsReportGenerator.SheetGenerators
{
    public interface ISheetUpdater
    {
        Task Update();
    }

    public abstract class BaseSheetUpdater : ISheetUpdater
    {
        public string SheetName { get; set; }

        private readonly string _applicationName = "GHA CI/CD Stats Reporting Tool";
        protected readonly string _spreadsheetId = "";
        private readonly GoogleCredential _credential;
        private readonly string _credentialPath = "";
        protected SheetsService _sheetsService;

        public BaseSheetUpdater()
        {
            _credential = GoogleCredential.FromFile(_credentialPath).CreateScoped(SheetsService.Scope.Spreadsheets);
            _sheetsService = CreateSheetsService();
        }

        public abstract Task Update();

        private SheetsService CreateSheetsService()
        {
            return new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = _credential,
                ApplicationName = _applicationName,
            });
        }

        /// <summary>
        /// The date we want report data until
        /// </summary>
        protected DateTime GetReportEndDate()
        {
            // we always use the next Monday for this date (if today is Monday, we use today)
            DateTime today = DateTime.Now.Date;
            int daysUntilMonday = ((int)DayOfWeek.Monday - (int)today.DayOfWeek + 7) % 7;

            // If today is Monday, return today's date
            if (daysUntilMonday == 0)
            {
                return today;
            }

            // Otherwise, return the date of the next Monday
            return today.AddDays(daysUntilMonday);
        }

        protected int GetSheetId()
        {
            var spreadsheet = _sheetsService.Spreadsheets.Get(_spreadsheetId).Execute();
            var sheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title == SheetName);

            if (sheet == null)
            {
                throw new Exception($"Sheet '{SheetName}' not found in spreadsheet.");
            }

            return sheet.Properties.SheetId.Value;
        }
    }
}
