using GitHubActionsReportGenerator.Repositories;
using GitHubActionsReportGenerator.SheetGenerators;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace GitHubActionsReportGenerator.SheetUpdaters
{
    public class FlakyTestFailuresSheetUpdater : BaseSheetUpdater
    {
        private readonly FlakyTestFailuresRepository _repository;

        public FlakyTestFailuresSheetUpdater(FlakyTestFailuresRepository repository)
        {
            SheetName = "Flaky Test Failures";
            _repository = repository;
        }

        public override async Task Update()
        {
            var endDate = DateTime.UtcNow.Date; //GetReportEndDate();

            // get the flaky tests
            // we only look at test results from the previous 4 weeks and a test must have failed at least once during the last week to be considered flaky
            var flakyTestFailures = await _repository.GetFlakyTestFailures(endDate);

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
            foreach (var flakyTestFailure in flakyTestFailures.OrderBy(t => t.StartedAtUtc))
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
                        new CellData { UserEnteredValue = new ExtendedValue { StringValue = flakyTestFailure.TestName } },
                        new CellData { UserEnteredValue = new ExtendedValue { NumberValue = flakyTestFailure.RunId } },
                        new CellData { UserEnteredValue = new ExtendedValue { NumberValue = flakyTestFailure.RunAttempt } },
                        new CellData { UserEnteredValue = new ExtendedValue { StringValue = flakyTestFailure.JobUrl } },
                        new CellData {
                            UserEnteredValue = new ExtendedValue { 
                                // Convert DateTime to Google Sheets serial number
                                NumberValue = (flakyTestFailure.StartedAtUtc - new DateTime(1899, 12, 30)).TotalDays
                            },
                            UserEnteredFormat = new CellFormat {
                                NumberFormat = new NumberFormat {
                                    Type = "DATE_TIME",
                                    Pattern = "yyyy-MM-dd HH:mm:ss.SSS"
                                }
                            }
                        },
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
