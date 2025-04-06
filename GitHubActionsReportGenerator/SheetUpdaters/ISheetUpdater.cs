using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;

namespace GitHubActionsReportGenerator.SheetGenerators
{
    public interface ISheetUpdater
    {
        Task Update();
    }

    public abstract class BaseSheetGenerator : ISheetUpdater
    {
        public string SheetName { get; set; }

        private readonly string _applicationName = "GHA CI/CD Stats Reporting Tool";
        protected readonly string _spreadsheetId = "";
        private readonly GoogleCredential _credential;
        private readonly string _credentialPath = "";

        public BaseSheetGenerator()
        {
            _credential = GoogleCredential.FromFile(_credentialPath).CreateScoped(SheetsService.Scope.Spreadsheets);
        }

        public abstract Task Update();

        protected SheetsService CreateSheetsService()
        {
            return new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = _credential,
                ApplicationName = _applicationName,
            });
        }
    }
}
