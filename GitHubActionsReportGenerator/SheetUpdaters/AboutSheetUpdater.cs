using GitHubActionsReportGenerator.SheetGenerators;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Sheets.v4;

namespace GitHubActionsReportGenerator.SheetUpdaters
{
    public class AboutSheetUpdater : BaseSheetUpdater
    {
        public AboutSheetUpdater()
        {
            SheetName = "About";
        }

        public override async Task Update()
        {
            // Get the data from the sheet
            string range = $"{SheetName}!A:A";
            var request = _sheetsService.Spreadsheets.Values.Get(_spreadsheetId, range);
            var response = await request.ExecuteAsync();
            var values = response.Values;

            int lastUpdatedRowIndex = -1;

            // Find the cell containing "Last Updated:"
            if (values != null)
            {
                for (int i = 0; i < values.Count; i++)
                {
                    if (values[i].Count > 0 && values[i][0].ToString().Contains("Last Updated:"))
                    {
                        lastUpdatedRowIndex = i + 1; // Convert to 1-based index for sheet API
                        break;
                    }
                }
            }

            if (lastUpdatedRowIndex > 0)
            {
                // Update the cell with current UTC time
                string updateValue = $"Last Updated: {DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss")} UTC";

                var updateRange = $"{SheetName}!A{lastUpdatedRowIndex}";
                var valueRange = new ValueRange
                {
                    Values = new List<IList<object>> { new List<object> { updateValue } }
                };

                var updateRequest = _sheetsService.Spreadsheets.Values.Update(valueRange, _spreadsheetId, updateRange);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                await updateRequest.ExecuteAsync();
            }
        }
    }
}
