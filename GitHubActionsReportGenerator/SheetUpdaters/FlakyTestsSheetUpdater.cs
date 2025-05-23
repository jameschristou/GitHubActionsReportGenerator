﻿using GitHubActionsReportGenerator.Repositories;
using GitHubActionsReportGenerator.SheetGenerators;
using Google.Apis.Sheets.v4.Data;

namespace GitHubActionsReportGenerator.SheetUpdaters
{
    public class FlakyTestsSheetUpdater : BaseSheetUpdater
    {
        private readonly FlakyTestsRepository _repository;

        public FlakyTestsSheetUpdater(FlakyTestsRepository repository)
        {
            SheetName = "Tests Likely To Be Flaky";
            _repository = repository;
        }

        public override async Task Update()
        {
            var endDate = DateTime.UtcNow.Date; //GetReportEndDate();

            // get the flaky tests
            // we only look at test results from the previous 4 weeks and a test must have failed at least once during the last week to be considered flaky
            var flakyTests = (await _repository.GetFlakyTests(endDate)).OrderBy(f => f.FailureCount);

            // how do we keep notes and jira numbers??
            // maybe we read the existing data and try and match any additional data added to the sheet by the name of the test?
            // TODO: we'll do this as an improvement

            // we no longer keep all the data for this. It's not really useful to see and hard to keep it up to date
            // we keep refreshing this sheet with the latest data

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
            foreach (var flakyTest in flakyTests)
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
                        new CellData { UserEnteredValue = new ExtendedValue { StringValue = flakyTest.TestName } },
                        new CellData { UserEnteredValue = new ExtendedValue { NumberValue = flakyTest.AvgDurationSeconds } },
                        new CellData { UserEnteredValue = new ExtendedValue { NumberValue = flakyTest.MaxDurationSeconds } },
                        new CellData { UserEnteredValue = new ExtendedValue { NumberValue = flakyTest.MinDurationSeconds } },
                        new CellData { UserEnteredValue = new ExtendedValue { NumberValue = flakyTest.FailureCount } },
                        new CellData { UserEnteredValue = new ExtendedValue { NumberValue = flakyTest.SuccessCount } },
                        new CellData { UserEnteredValue = new ExtendedValue { NumberValue = flakyTest.FailurePercentage } },
                        new CellData { UserEnteredValue = new ExtendedValue { NumberValue = flakyTest.NumberOfRunsImpacted } },
                        new CellData { UserEnteredValue = new ExtendedValue { NumberValue = flakyTest.FailureCountLast4Weeks } }
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
