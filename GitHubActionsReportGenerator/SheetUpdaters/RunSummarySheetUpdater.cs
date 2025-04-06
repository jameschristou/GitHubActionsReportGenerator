using GitHubActionsReportGenerator.Repositories;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GitHubActionsReportGenerator.SheetGenerators
{
    public class RunSummarySheetUpdater : BaseSheetGenerator
    {
        private readonly IRunSummaryRepository _runSummaryRepository;

        public RunSummarySheetUpdater(IRunSummaryRepository runSummaryRepository)
        {
            SheetName = "Run Summary";
            _runSummaryRepository = runSummaryRepository;
        }

        public override async Task Update()
        {
            // then check whether we are inserting a new record or updating an existing one
            var lastWeekStarting = GetWeekStartingOfMostRecentRecordOnReport();

            // ok now we need to loop through the data and insert records where needed or update existing records
            // first get the data required from the DB
            var runSummary = await _runSummaryRepository.GetRunSummary(DateTime.Now.AddDays(-7));

            // we're only going to update/insert those weeks greater than/equal to the last week starting date
            var recordsToInsert = runSummary.ToList();

            // update the first record
            var recordToUpdate = recordsToInsert.FirstOrDefault(x => x.WeekStarting == lastWeekStarting);

            var requestBody = new Google.Apis.Sheets.v4.Data.BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Google.Apis.Sheets.v4.Data.Request>()
            };

            // Get sheet ID
            var service = CreateSheetsService();

            var spreadsheet = service.Spreadsheets.Get(_spreadsheetId).Execute();
            var sheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title == SheetName);

            if (sheet == null)
            {
                throw new Exception($"Sheet '{SheetName}' not found in spreadsheet.");
            }

            var sheetId = sheet.Properties.SheetId.Value;

            try
            {
                // Create data to be updated in the second row (row 2)
                var rowData = new Google.Apis.Sheets.v4.Data.RowData
                {
                    Values = new List<Google.Apis.Sheets.v4.Data.CellData>
                    {
                        new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { StringValue = recordToUpdate.WeekStarting.ToString("yyyy-MM-dd") } },
                        new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { NumberValue = recordToUpdate.AvgDurationForSuccessfulRuns } },
                        new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { NumberValue = recordToUpdate.MaxDurationForSuccessfulRuns } },
                        new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { NumberValue = recordToUpdate.MinDurationForSuccessfulRuns } },
                        new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { NumberValue = recordToUpdate.AvgAttemptsForSuccessfulRuns } },
                        new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { NumberValue = recordToUpdate.FailureCount } },
                        new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { NumberValue = recordToUpdate.SuccessCount } },
                        new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { NumberValue = recordToUpdate.SuccessPercentage } },
                        new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { NumberValue = recordToUpdate.FlakyTestCount } },
                        new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { NumberValue = recordToUpdate.MinutesWastedDueToFailingTests } },
                        new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { NumberValue = recordToUpdate.TotalMinutesWasted } }
                    }
                };

                requestBody.Requests.Add(new Google.Apis.Sheets.v4.Data.Request
                {
                    UpdateCells = new Google.Apis.Sheets.v4.Data.UpdateCellsRequest
                    {
                        Start = new Google.Apis.Sheets.v4.Data.GridCoordinate
                        {
                            SheetId = sheetId,
                            RowIndex = 1,   // Second row (0-based index)
                            ColumnIndex = 0  // First column (0-based index)
                        },
                        Rows = new List<Google.Apis.Sheets.v4.Data.RowData> { rowData },
                        Fields = "userEnteredValue"
                    }
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating record: {ex.Message}");
            }


            foreach (var item in recordsToInsert.Where(x => x.WeekStarting > lastWeekStarting).OrderBy(r => r.WeekStarting))
            {
                try
                {
                    // Prepare the data to be inserted
                    var rowData = new Google.Apis.Sheets.v4.Data.RowData
                    {
                        Values = new List<Google.Apis.Sheets.v4.Data.CellData>
                        {
                            new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { StringValue = item.WeekStarting.ToString("yyyy-MM-dd") } },
                            new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { NumberValue = item.AvgDurationForSuccessfulRuns } },
                            new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { NumberValue = item.MaxDurationForSuccessfulRuns } },
                            new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { NumberValue = item.MinDurationForSuccessfulRuns } },
                            new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { NumberValue = item.AvgAttemptsForSuccessfulRuns } },
                            new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { NumberValue = item.FailureCount } },
                            new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { NumberValue = item.SuccessCount } },
                            new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { NumberValue = item.SuccessPercentage } },
                            new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { NumberValue = item.FlakyTestCount } },
                            new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { NumberValue = item.MinutesWastedDueToFailingTests } },
                            new Google.Apis.Sheets.v4.Data.CellData { UserEnteredValue = new Google.Apis.Sheets.v4.Data.ExtendedValue { NumberValue = item.TotalMinutesWasted } }
                        }
                    };

                    // request to insert a row
                    requestBody.Requests.Add(new Google.Apis.Sheets.v4.Data.Request
                    {
                        InsertDimension = new Google.Apis.Sheets.v4.Data.InsertDimensionRequest
                        {
                            Range = new Google.Apis.Sheets.v4.Data.DimensionRange
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
                    requestBody.Requests.Add(new Google.Apis.Sheets.v4.Data.Request
                    {
                        UpdateCells = new Google.Apis.Sheets.v4.Data.UpdateCellsRequest
                        {
                            Start = new Google.Apis.Sheets.v4.Data.GridCoordinate
                            {
                                SheetId = sheetId,
                                RowIndex = 1,   // Second row (0-based index)
                                ColumnIndex = 0  // First column (0-based index)
                            },
                            Rows = new List<Google.Apis.Sheets.v4.Data.RowData> { rowData },
                            Fields = "userEnteredValue"
                        }
                    });

                    Console.WriteLine($"Inserted and populated new row with data for {item.WeekStarting.ToString("yyyy-MM-dd")}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error inserting record: {ex.Message}");
                }
            }

            // Execute the batch update
            var batchUpdateRequest = service.Spreadsheets.BatchUpdate(requestBody, _spreadsheetId);
            var batchUpdateResponse = batchUpdateRequest.Execute();

            // populates data into the appropriate row
            return;
        }


        private DateTime GetWeekStartingOfMostRecentRecordOnReport()
        {
            try
            {
                // Create Google Sheets API service
                var service = CreateSheetsService();

                // Define the range to read (A2 - first column of second row)
                string range = $"{SheetName}!A2:A2";

                // Request data from the spreadsheet
                var request = service.Spreadsheets.Values.Get(_spreadsheetId, range);
                var response = request.Execute();
                var values = response.Values;

                // Check if we have data in the response
                if (values != null && values.Count > 0 && values[0].Count > 0)
                {
                    // Parse the date from the first column
                    if (DateTime.TryParse(values[0][0].ToString(), out DateTime weekStarting))
                    {
                        return weekStarting;
                    }
                }

                // Return a default date if no valid date was found
                return DateTime.MinValue;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading date from sheet: {ex.Message}");
                return DateTime.MinValue;
            }

        }
    }
}
