using GitHubActionsReportGenerator.Repositories;
using GitHubActionsReportGenerator.SheetGenerators;
using Google.Apis.Sheets.v4.Data;

namespace GitHubActionsReportGenerator.SheetUpdaters
{
    public class SlowestTestsSheetUpdater : BaseSheetGenerator
    {
        private readonly SlowestTestsRepository _repository;

        public SlowestTestsSheetUpdater(SlowestTestsRepository repository)
        {
            SheetName = "Slowest Tests";
            _repository = repository;
        }

        public override async Task Update()
        {
            var endDate = DateTime.UtcNow.Date; //GetReportEndDate();

            var slowTests = await _repository.GetSlowTests(endDate);

            var requestBody = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request>()
            };

            // so first we clean out the existing data
            var clearRequest = _sheetsService.Spreadsheets.Values.Clear(
                new ClearValuesRequest(),
                _spreadsheetId,
                $"{SheetName}!A2:Z");

            await clearRequest.ExecuteAsync();

            var sheetId = GetSheetId();

            // then we get the new data
            foreach (var slowTest in slowTests.OrderBy(t => t.MinDurationSecondsLastWeek))
            {
                // request to insert a row
                requestBody.Requests.Add(new Request
                {
                    InsertDimension = new InsertDimensionRequest
                    {
                        Range = new DimensionRange
                        {
                            SheetId = sheetId,
                            Dimension = "ROWS",
                            StartIndex = 1,  // 0-based index, so 1 is the second row
                            EndIndex = 2     // Insert 1 row (end is exclusive)
                        },
                        InheritFromBefore = false
                    }
                });

                // Request to update the newly inserted row with data
                var rowData = new RowData
                {
                    Values = new List<CellData>
                    {
                        new CellData { UserEnteredValue = new ExtendedValue { StringValue = slowTest.TestName } },
                        new CellData { UserEnteredValue = new ExtendedValue { NumberValue = slowTest.AvgDurationSecondsLastWeek } },
                        new CellData { UserEnteredValue = new ExtendedValue { NumberValue = slowTest.MaxDurationSecondsLastWeek } },
                        new CellData { UserEnteredValue = new ExtendedValue { NumberValue = slowTest.MinDurationSecondsLastWeek } },
                        new CellData { UserEnteredValue = new ExtendedValue { NumberValue = slowTest.AvgDurationSecondsLast4Weeks } },
                        new CellData { UserEnteredValue = new ExtendedValue { NumberValue = slowTest.MaxDurationSecondsLast4Weeks } },
                        new CellData { UserEnteredValue = new ExtendedValue { NumberValue = slowTest.MinDurationSecondsLast4Weeks } }
                    }
                };

                // Request to update the newly inserted row with data
                requestBody.Requests.Add(new Request
                {
                    UpdateCells = new UpdateCellsRequest
                    {
                        Start = new GridCoordinate
                        {
                            SheetId = sheetId,
                            RowIndex = 1,   // Second row (0-based index)
                            ColumnIndex = 0  // First column (0-based index)
                        },
                        Rows = new List<RowData> { rowData },
                        Fields = "userEnteredValue,userEnteredFormat.numberFormat"
                    }
                });
            }

            // Execute the batch update
            var batchUpdateRequest = _sheetsService.Spreadsheets.BatchUpdate(requestBody, _spreadsheetId);
            var batchUpdateResponse = batchUpdateRequest.Execute();
        }
    }
}
