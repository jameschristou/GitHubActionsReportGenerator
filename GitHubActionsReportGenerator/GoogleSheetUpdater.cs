using GitHubActionsReportGenerator.Entities;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using System.Linq;

namespace GitHubActionsReportGenerator
{
    public interface IGoogleSheetUpdater
    {
        Task UpdateSheet();
    }

    public class GoogleSheetUpdater : IGoogleSheetUpdater
    {
        private readonly string _applicationName = "GHA CI/CD Stats Reporting Tool";
        private readonly string _spreadsheetId = "";
        private readonly GoogleCredential _credential;
        private readonly string _credentialPath = "";

        public GoogleSheetUpdater()
        {
            _credential = GoogleCredential.FromFile(_credentialPath).CreateScoped(SheetsService.Scope.Spreadsheets);
        }

        public async Task UpdateSheet()
        {
            var sheetName = "Run Summary";

            Console.WriteLine($"Updating sheet '{sheetName}'");

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = _credential,
                ApplicationName = _applicationName,
            });

            InsertEmptyNewRow(service, sheetName);


            //// Prepare data for the sheet
            //var rows = new List<IList<object>>();

            //// Add header row
            //rows.Add(new List<object> { "Name", "DateCreatedUtc", "DateStatusChangedUtc", "Creation Jira", "Created By", "Removal Jira" });

            //// Add data rows
            //foreach (var flag in featureFlags.Where(sheetConfig.Predicate).OrderBy(ff => ff.DateFFStatusChangedUtc))
            //{
            //    rows.Add(new List<object>
            //    {
            //        flag.Name,
            //        flag.DateFFCreatedUtc.ToString("yyyy-MM-dd HH:mm:ss"),
            //        flag.DateFFStatusChangedUtc.ToString("yyyy-MM-dd HH:mm:ss"),
            //        flag.CreationJiraIssueKey,
            //        flag.CreationJiraIssueAssignedTo,
            //        flag.RemovalJiraIssueKey
            //    });
            //}

            //// Create the value range object
            //var valueRange = new ValueRange
            //{
            //    Values = rows
            //};

            //// Write data to the sheet
            //var updateRequest = service.Spreadsheets.Values.Update(
            //    valueRange,
            //    _spreadsheetId,
            //    $"{sheetName}!A1");
            //updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

            //await updateRequest.ExecuteAsync();

            //var sheetId = await GetSheetId(service, sheetName);

            //// Apply formatting to the header row
            //var requests = new List<Request>
            //{
            //    // Bold text for header row
            //    new Request
            //    {
            //        RepeatCell = new RepeatCellRequest
            //        {
            //            Range = new GridRange
            //            {
            //                SheetId = sheetId,
            //                StartRowIndex = 0,
            //                EndRowIndex = 1
            //            },
            //            Cell = new CellData
            //            {
            //                UserEnteredFormat = new CellFormat
            //                {
            //                    TextFormat = new TextFormat
            //                    {
            //                        Bold = true
            //                    },
            //                    BackgroundColor = new Color
            //                    {
            //                        Red = 0.827f,
            //                        Green = 0.827f,
            //                        Blue = 0.827f
            //                    }
            //                }
            //            },
            //            Fields = "userEnteredFormat(textFormat,backgroundColor)"
            //        }
            //    },
            //    // Freeze the header row
            //    new Request
            //    {
            //        UpdateSheetProperties = new UpdateSheetPropertiesRequest
            //        {
            //            Properties = new SheetProperties
            //            {
            //                SheetId = sheetId,
            //                GridProperties = new GridProperties
            //                {
            //                    FrozenRowCount = 1
            //                }
            //            },
            //            Fields = "gridProperties.frozenRowCount"
            //        }
            //    }
            //};

            //var batchUpdateRequest = new BatchUpdateSpreadsheetRequest
            //{
            //    Requests = requests
            //};

            //await service.Spreadsheets.BatchUpdate(batchUpdateRequest, _spreadsheetId).ExecuteAsync();

            //await SetColumnWidths(service, sheetName);
        }

        private async Task SetColumnWidths(SheetsService service, string sheetName)
        {
            try
            {
                int sheetId = await GetSheetId(service, sheetName);

                // Define column widths (in pixels)
                var columnWidths = new[]
                {
                    350,  // Name (A) - wider for potentially long feature flag names
                    150,  // DateCreatedUtc (B)
                    160,  // DateStatusChangedUtc (C)
                    110,  // Creation Jira (D)
                    200,  // Created By (E)
                    110   // Removal Jira (F)
                };

                var requests = new List<Request>();

                // Add dimension properties for each column
                for (int i = 0; i < columnWidths.Length; i++)
                {
                    requests.Add(new Request
                    {
                        UpdateDimensionProperties = new UpdateDimensionPropertiesRequest
                        {
                            Range = new DimensionRange
                            {
                                SheetId = sheetId,
                                Dimension = "COLUMNS",
                                StartIndex = i,
                                EndIndex = i + 1
                            },
                            Properties = new DimensionProperties
                            {
                                PixelSize = columnWidths[i]
                            },
                            Fields = "pixelSize"
                        }
                    });
                }

                // Apply the column width settings
                var batchUpdateRequest = new BatchUpdateSpreadsheetRequest
                {
                    Requests = requests
                };

                await service.Spreadsheets.BatchUpdate(batchUpdateRequest, _spreadsheetId).ExecuteAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error setting column widths for sheet '{sheetName}': {ex.Message}");
                throw;
            }
        }

        private async Task<int> GetSheetId(SheetsService service, string sheetName)
        {
            try
            {
                var spreadsheet = await service.Spreadsheets.Get(_spreadsheetId).ExecuteAsync();
                var sheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title == sheetName);

                if (sheet == null)
                {
                    throw new Exception($"Sheet '{sheetName}' not found in the spreadsheet.");
                }

                return sheet.Properties.SheetId.Value;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving sheet ID for '{sheetName}': {ex.Message}");
                throw;
            }
        }

        private async Task EnsureSheetExistsAndIsEmpty(SheetsService service, string sheetName)
        {
            // Clear existing content
            // Ensure the sheet exists
            if (!await SheetExists(service, sheetName))
            {
                await CreateSheet(service, sheetName);
                return;
            }

            var clearRequest = service.Spreadsheets.Values.Clear(
                new ClearValuesRequest(),
                _spreadsheetId,
                $"{sheetName}!A2:Z");
            await clearRequest.ExecuteAsync();
        }

        private async Task<bool> SheetExists(SheetsService service, string sheetName)
        {
            try
            {
                var spreadsheet = await service.Spreadsheets.Get(_spreadsheetId).ExecuteAsync();
                return spreadsheet.Sheets.Any(s => s.Properties.Title == sheetName);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking if sheet exists: {ex.Message}");
                return false;
            }
        }

        private async Task CreateSheet(SheetsService service, string sheetName)
        {
            try
            {
                var batchUpdateRequest = new BatchUpdateSpreadsheetRequest
                {
                    Requests = new List<Request>
                    {
                        new Request
                        {
                            AddSheet = new AddSheetRequest
                            {
                                Properties = new SheetProperties
                                {
                                    Title = sheetName
                                }
                            }
                        }
                    }
                };

                var batchUpdateResponse = await service.Spreadsheets.BatchUpdate(batchUpdateRequest, _spreadsheetId).ExecuteAsync();

                Console.WriteLine($"Created new sheet: {sheetName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating sheet {sheetName}: {ex.Message}");
                throw;
            }
        }

        private async Task InsertEmptyNewRow(SheetsService service, string sheetName)
        {
            try
            {
                int sheetId = await GetSheetId(service, sheetName);

                var requests = new List<Request>
                {
                    new Request
                    {
                        InsertDimension = new InsertDimensionRequest
                        {
                            Range = new DimensionRange
                            {
                                SheetId = sheetId,
                                Dimension = "ROWS",
                                StartIndex = 1,  // 0-based index, so 1 means the second row (above row 2)
                                EndIndex = 2     // Insert 1 row (EndIndex - StartIndex = number of rows)
                            },
                            InheritFromBefore = false
                        }
                    }
                };

                var batchUpdateRequest = new BatchUpdateSpreadsheetRequest
                {
                    Requests = requests
                };

                await service.Spreadsheets.BatchUpdate(batchUpdateRequest, _spreadsheetId).ExecuteAsync();

                Console.WriteLine($"Inserted empty row at position 2 in sheet '{sheetName}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error inserting empty row in sheet '{sheetName}': {ex.Message}");
                throw;
            }
        }
    }
}
