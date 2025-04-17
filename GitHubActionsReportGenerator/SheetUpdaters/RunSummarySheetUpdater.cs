using GitHubActionsReportGenerator.Repositories;
using Google.Apis.Sheets.v4.Data;

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

            var endDate = GetNextMonday();

            // ok now we need to loop through the data and insert records where needed or update existing records
            // first get the data required from the DB
            var runSummary = await _runSummaryRepository.GetRunSummary(endDate);

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

            requestBody.Requests.Add(CreateUpdateRowRequest(recordToUpdate, sheetId));

            foreach (var item in recordsToInsert.Where(x => x.WeekStarting > lastWeekStarting).OrderBy(r => r.WeekStarting))
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
                requestBody.Requests.Add(CreateUpdateRowRequest(item, sheetId));
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

        private DateTime GetNextMonday()
        {
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

        private Request CreateUpdateRowRequest(WeeklyRunSummary data, int sheetId)
        {
            // Prepare the data to be inserted
            var rowData = new RowData
            {
                Values = new List<CellData>
                {
                    new CellData {
                        UserEnteredValue = new ExtendedValue { 
                            // Convert DateTime to Google Sheets serial number
                            NumberValue = (data.WeekStarting - new DateTime(1899, 12, 30)).TotalDays
                        },
                        UserEnteredFormat = new CellFormat {
                           NumberFormat = new NumberFormat {
                               Type = "DATE",
                               Pattern = "yyyy-MM-dd"
                           }
                       }
                    },
                    new CellData { UserEnteredValue = new ExtendedValue { NumberValue = data.AvgDurationForSuccessfulRuns } },
                    new CellData { UserEnteredValue = new ExtendedValue { NumberValue = data.MaxDurationForSuccessfulRuns } },
                    new CellData { UserEnteredValue = new ExtendedValue { NumberValue = data.MinDurationForSuccessfulRuns } },
                    new CellData { UserEnteredValue = new ExtendedValue { NumberValue = data.AvgAttemptsForSuccessfulRuns } },
                    new CellData { UserEnteredValue = new ExtendedValue { NumberValue = data.FailureCount } },
                    new CellData { UserEnteredValue = new ExtendedValue { NumberValue = data.SuccessCount } },
                    new CellData { UserEnteredValue = new ExtendedValue { NumberValue = data.SuccessPercentage } },
                    new CellData { UserEnteredValue = new ExtendedValue { NumberValue = data.FlakyTestCount } },
                    new CellData { UserEnteredValue = new ExtendedValue { NumberValue = data.MinutesWastedDueToFailingTests } },
                    new CellData { UserEnteredValue = new ExtendedValue { NumberValue = data.TotalMinutesWasted } }
                }
            };

            // Request to update the newly inserted row with data
            return new Request
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
            };
        }
    }
}
