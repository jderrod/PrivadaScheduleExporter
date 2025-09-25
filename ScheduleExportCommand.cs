using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ScheduleExtractor
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ScheduleExportCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Show dialog to get custom output path
                string customOutputPath = ShowOutputPathDialog();
                
                // Create output directory
                string outputDir;
                if (!string.IsNullOrWhiteSpace(customOutputPath) && Directory.Exists(Path.GetDirectoryName(customOutputPath)))
                {
                    outputDir = Path.Combine(customOutputPath, "ScheduleExports");
                }
                else
                {
                    // Fallback to default location
                    outputDir = Path.Combine(Path.GetDirectoryName(doc.PathName) ?? Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "ScheduleExports");
                }
                
                Directory.CreateDirectory(outputDir);

                // Create debug log file
                string debugLogPath = Path.Combine(outputDir, "debug_log.txt");
                var debugLines = new List<string>();
                debugLines.Add($"=== Schedule Export Debug Log - {DateTime.Now} ===");
                debugLines.Add($"Document: {doc.PathName ?? "Unsaved Document"}");
                debugLines.Add("");

                // Get all view schedules in the document
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                ICollection<Element> schedules = collector.OfClass(typeof(ViewSchedule)).ToElements();

                int exportedCount = 0;
                int totalSchedulesFound = schedules.Count;
                var debugInfo = new List<string>();
                
                // Dictionary to store data for the combined JSONs
                var combinedDataFull = new Dictionary<string, object>();
                var combinedDataDPS = new Dictionary<string, object>();
                
                var allTargetSchedules = new HashSet<string> 
                { 
                    "Head_Rails_Segments", 
                    "Privada_Door", 
                    "Privada_Fascia", 
                    "Privada_Head_Rail", 
                    "Privada_Panel", 
                    "Privada_Pedestal" 
                };
                
                var dpsSchedules = new HashSet<string>
                {
                    "Privada_Door", 
                    "Privada_Panel", 
                    "Privada_Fascia"
                };

                foreach (ViewSchedule schedule in schedules.Cast<ViewSchedule>())
                {
                    try
                    {
                        // Debug information
                        string scheduleInfo = $"Found schedule: '{schedule.Name}' - Type: {schedule.ViewType} - Category: {schedule.Definition.CategoryId}";
                        debugInfo.Add(scheduleInfo);
                        debugLines.Add(scheduleInfo);
                        System.Diagnostics.Debug.WriteLine(scheduleInfo);

                        // Only skip schedules that are clearly not data schedules
                        // Remove the restrictive CategoryId filter - let's try to export all schedules
                        
                        // Export the schedule
                        var scheduleData = ExtractScheduleData(schedule, debugLines);
                        if (scheduleData.Count > 0)
                        {
                            // Apply custom formatting for specific schedules
                            scheduleData = ApplyCustomFormatting(schedule.Name, scheduleData);
                            
                            // Clean the data (remove header row and fix column names)
                            var cleanedData = CleanScheduleData(scheduleData);
                            
                            // Add to combined data if it's one of the target schedules
                            if (allTargetSchedules.Contains(schedule.Name))
                            {
                                combinedDataFull[schedule.Name] = cleanedData;
                            }
                            
                            if (dpsSchedules.Contains(schedule.Name))
                            {
                                combinedDataDPS[schedule.Name] = cleanedData;
                            }
                            
                            string fileName = SanitizeFileName(schedule.Name) + ".json";
                            string filePath = Path.Combine(outputDir, fileName);

                            string jsonString = JsonConvert.SerializeObject(scheduleData, Formatting.Indented);
                            File.WriteAllText(filePath, jsonString);

                            exportedCount++;
                        }
                        else
                        {
                            debugInfo.Add($"  -> No data found for schedule: '{schedule.Name}'");
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log error for individual schedule but continue with others
                        string errorInfo = $"Error processing schedule '{schedule.Name}': {ex.Message}";
                        debugInfo.Add($"  -> ERROR: {errorInfo}");
                        System.Diagnostics.Debug.WriteLine(errorInfo);
                    }
                }

                // Create combined JSON files with timestamps
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm");
                
                // Full combined JSON
                if (combinedDataFull.Count > 0)
                {
                    string fullFilePath = Path.Combine(outputDir, $"combined_json_full_{timestamp}.json");
                    string fullJsonString = JsonConvert.SerializeObject(combinedDataFull, Formatting.Indented);
                    File.WriteAllText(fullFilePath, fullJsonString);
                    
                    debugLines.Add($"Created combined_json_full with {combinedDataFull.Count} schedules: {string.Join(", ", combinedDataFull.Keys)}");
                }
                
                // DPS combined JSON
                if (combinedDataDPS.Count > 0)
                {
                    string dpsFilePath = Path.Combine(outputDir, $"combined_json_DPS_{timestamp}.json");
                    string dpsJsonString = JsonConvert.SerializeObject(combinedDataDPS, Formatting.Indented);
                    File.WriteAllText(dpsFilePath, dpsJsonString);
                    
                    debugLines.Add($"Created combined_json_DPS with {combinedDataDPS.Count} schedules: {string.Join(", ", combinedDataDPS.Keys)}");
                }

                // Write debug log to file
                debugLines.Add("");
                debugLines.Add($"=== SUMMARY ===");
                debugLines.Add($"Total schedules found: {totalSchedulesFound}");
                debugLines.Add($"Successfully exported: {exportedCount}");
                debugLines.Add($"Combined full schedules: {combinedDataFull.Count}");
                debugLines.Add($"Combined DPS schedules: {combinedDataDPS.Count}");
                File.WriteAllLines(debugLogPath, debugLines);

                // Show success message with debug info
                string resultMessage = $"Found {totalSchedulesFound} total schedules.\nSuccessfully exported {exportedCount} schedules to:\n{outputDir}";
                
                if (combinedDataFull.Count > 0 || combinedDataDPS.Count > 0)
                {
                    resultMessage += $"\n\n📦 Combined JSONs created:";
                    if (combinedDataFull.Count > 0)
                        resultMessage += $"\n• combined_json_full_{timestamp}.json ({combinedDataFull.Count} schedules)";
                    if (combinedDataDPS.Count > 0)
                        resultMessage += $"\n• combined_json_DPS_{timestamp}.json ({combinedDataDPS.Count} schedules)";
                }
                
                resultMessage += "\n\nDebug log written to: debug_log.txt";
                
                // Add debug info if no schedules were exported
                if (exportedCount == 0 && debugInfo.Count > 0)
                {
                    resultMessage += "\n\nFirst few debug entries:\n" + string.Join("\n", debugInfo.Take(5)); // Show first 5 entries
                }
                
                TaskDialog.Show("Schedule Export Complete", resultMessage);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error: {ex.Message}";
                return Result.Failed;
            }
        }

        private List<Dictionary<string, object>> ExtractScheduleData(ViewSchedule schedule, List<string> debugLines = null)
        {
            var scheduleData = new List<Dictionary<string, object>>();

            try
            {
                // Try to access the schedule data more carefully
                TableData tableData = null;
                try
                {
                    tableData = schedule.GetTableData();
                }
                catch (Exception ex)
                {
                    string msg = $"  -> Cannot get table data for schedule '{schedule.Name}': {ex.Message}";
                    System.Diagnostics.Debug.WriteLine(msg);
                    debugLines?.Add(msg);
                    return scheduleData;
                }

                if (tableData == null)
                {
                    string msg = $"  -> No table data for schedule: {schedule.Name}";
                    System.Diagnostics.Debug.WriteLine(msg);
                    debugLines?.Add(msg);
                    return scheduleData;
                }

                TableSectionData sectionData = tableData.GetSectionData(SectionType.Body);
                if (sectionData == null)
                {
                    string msg = $"  -> No body section data for schedule: {schedule.Name}";
                    System.Diagnostics.Debug.WriteLine(msg);
                    debugLines?.Add(msg);
                    return scheduleData;
                }

                string rowColMsg = $"  -> Schedule '{schedule.Name}' has {sectionData.NumberOfRows} rows and {sectionData.NumberOfColumns} columns";
                System.Diagnostics.Debug.WriteLine(rowColMsg);
                debugLines?.Add(rowColMsg);

                if (sectionData.NumberOfRows == 0) // No data rows
                {
                    string msg = $"  -> Schedule '{schedule.Name}' has no data rows";
                    System.Diagnostics.Debug.WriteLine(msg);
                    debugLines?.Add(msg);
                    return scheduleData;
                }

                // Get column headers - try multiple approaches
                var headers = new List<string>();
                int numberOfColumns = sectionData.NumberOfColumns;
                
                // First try to get headers from the header section
                TableSectionData headerData = null;
                try
                {
                    headerData = tableData.GetSectionData(SectionType.Header);
                }
                catch (Exception ex)
                {
                    string msg = $"  -> Could not get header section for '{schedule.Name}': {ex.Message}";
                    System.Diagnostics.Debug.WriteLine(msg);
                    debugLines?.Add(msg);
                }

                for (int col = 0; col < numberOfColumns; col++)
                {
                    string headerText = "";
                    
                    // Try different methods to get header text
                    if (headerData != null && headerData.NumberOfRows > 0)
                    {
                        try
                        {
                            headerText = schedule.GetCellText(SectionType.Header, 0, col);
                        }
                        catch (Exception)
                        {
                            // If header section fails, try getting from schedule definition
                        }
                    }
                    
                    // If still no header text, try from the schedule definition fields
                    if (string.IsNullOrWhiteSpace(headerText))
                    {
                        try
                        {
                            var definition = schedule.Definition;
                            if (col < definition.GetFieldCount())
                            {
                                var field = definition.GetField(col);
                                headerText = field.GetName();
                            }
                        }
                        catch (Exception)
                        {
                            // Fall back to generic column name
                            headerText = $"Column_{col}";
                        }
                    }
                    
                    if (string.IsNullOrWhiteSpace(headerText))
                        headerText = $"Column_{col}";
                    headers.Add(headerText);
                }
                
                string headerMsg = $"  -> Headers for '{schedule.Name}': {string.Join(", ", headers)}";
                System.Diagnostics.Debug.WriteLine(headerMsg);
                debugLines?.Add(headerMsg);

                // Extract data rows
                for (int row = 0; row < sectionData.NumberOfRows; row++)
                {
                    var rowData = new Dictionary<string, object>();
                    bool hasData = false;
                    var cellValues = new List<string>();

                    for (int col = 0; col < numberOfColumns; col++)
                    {
                        string cellText = "";
                        try
                        {
                            cellText = schedule.GetCellText(SectionType.Body, row, col) ?? "";
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error getting cell text for '{schedule.Name}' row {row}, col {col}: {ex.Message}");
                            cellText = "";
                        }
                        
                        cellValues.Add($"'{cellText}'");
                        
                        if (!string.IsNullOrEmpty(cellText) && !string.IsNullOrWhiteSpace(cellText))
                        {
                            hasData = true;
                        }

                        // Try to parse numeric values
                        object cellValue = cellText;
                        if (!string.IsNullOrEmpty(cellText))
                        {
                            cellValue = ParseNumericValue(cellText);
                        }

                        rowData[headers[col]] = cellValue;
                    }

                    string rowMsg = $"    Row {row}: [{string.Join(", ", cellValues)}] - HasData: {hasData}";
                    System.Diagnostics.Debug.WriteLine(rowMsg);
                    debugLines?.Add(rowMsg);

                    // Only add rows that have some data
                    if (hasData)
                    {
                        scheduleData.Add(rowData);
                    }
                }
                
                string extractMsg = $"  -> Extracted {scheduleData.Count} data rows from '{schedule.Name}'";
                System.Diagnostics.Debug.WriteLine(extractMsg);
                debugLines?.Add(extractMsg);
                debugLines?.Add(""); // Add blank line between schedules
            }
            catch (Exception ex)
            {
                string errorMsg = $"  -> ERROR extracting data from schedule '{schedule.Name}': {ex.Message}";
                System.Diagnostics.Debug.WriteLine(errorMsg);
                debugLines?.Add(errorMsg);
            }

            return scheduleData;
        }

        private string SanitizeFileName(string fileName)
        {
            string sanitized = fileName;
            char[] invalidChars = Path.GetInvalidFileNameChars();

            foreach (char c in invalidChars)
            {
                sanitized = sanitized.Replace(c, '_');
            }

            // Also replace some additional characters that might cause issues
            sanitized = sanitized.Replace(" ", "_")
                                .Replace("(", "")
                                .Replace(")", "")
                                .Replace("[", "")
                                .Replace("]", "")
                                .Replace("{", "")
                                .Replace("}", "");

            return sanitized;
        }

        private List<Dictionary<string, object>> ApplyCustomFormatting(string scheduleName, List<Dictionary<string, object>> originalData)
        {
            try
            {
                // Handle Head_Rails_Segments - convert to single object with column headers as keys
                if (scheduleName == "Head_Rails_Segments")
                {
                    return FormatHeadRailSegments(originalData);
                }
                
                // Handle Privada_Pedestal - create single object with family name, height, and quantity
                if (scheduleName == "Privada_Pedestal")
                {
                    return FormatPedestal(originalData);
                }
                
                // Handle Privada_Head_Rail - modify last object to show quantity
                if (scheduleName == "Privada_Head_Rail")
                {
                    return FormatHeadRail(originalData);
                }
                
                // Return original data for other schedules
                return originalData;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error applying custom formatting to {scheduleName}: {ex.Message}");
                return originalData; // Return original data if formatting fails
            }
        }

        private List<Dictionary<string, object>> FormatHeadRailSegments(List<Dictionary<string, object>> originalData)
        {
            // Create single object where column headers are keys and data row values are values
            if (originalData.Count >= 2)
            {
                var headerRow = originalData[0];
                var dataRow = originalData[1];
                var result = new Dictionary<string, object>();

                foreach (var kvp in headerRow)
                {
                    if (dataRow.ContainsKey(kvp.Key))
                    {
                        result[kvp.Value.ToString()] = dataRow[kvp.Key];
                    }
                }

                return new List<Dictionary<string, object>> { result };
            }
            return originalData;
        }

        private List<Dictionary<string, object>> FormatPedestal(List<Dictionary<string, object>> originalData)
        {
            // Create single object: family name (col 1), pedestal height (col 2), quantity: 7
            if (originalData.Count >= 2)
            {
                var dataRow = originalData[1]; // Second row has the actual data
                var result = new Dictionary<string, object>();

                // Get family name (first column) and pedestal height (second column)
                var keys = dataRow.Keys.ToList();
                if (keys.Count >= 2)
                {
                    result["family_name"] = dataRow[keys[0]];
                    result["pedestal_height"] = dataRow[keys[1]];
                    result["quantity"] = 7; // As specified by user
                }

                return new List<Dictionary<string, object>> { result };
            }
            return originalData;
        }

        private List<Dictionary<string, object>> FormatHeadRail(List<Dictionary<string, object>> originalData)
        {
            // Modify last object to show quantity: 6
            if (originalData.Count > 0)
            {
                var result = new List<Dictionary<string, object>>(originalData);
                var lastObject = new Dictionary<string, object>();
                
                // Create a summary object for the head rail quantity
                lastObject["head_rail_quantity"] = 6; // As specified by user
                lastObject["family_name"] = "Privada-Head_rail";
                
                // Replace the last object with our custom one
                result[result.Count - 1] = lastObject;
                
                return result;
            }
            return originalData;
        }

        private List<Dictionary<string, object>> CleanScheduleData(List<Dictionary<string, object>> originalData)
        {
            try
            {
                if (originalData == null || originalData.Count <= 1)
                    return originalData;

                // Remove the first row (header row) and process the remaining data
                var cleanedData = new List<Dictionary<string, object>>();
                
                for (int i = 1; i < originalData.Count; i++) // Skip first row (index 0)
                {
                    var row = originalData[i];
                    var cleanedRow = new Dictionary<string, object>();
                    
                    bool isFirstColumn = true;
                    foreach (var kvp in row)
                    {
                        if (isFirstColumn)
                        {
                            // Rename first column to "component_ID"
                            cleanedRow["component_ID"] = kvp.Value;
                            isFirstColumn = false;
                        }
                        else
                        {
                            // Keep other columns as-is
                            cleanedRow[kvp.Key] = kvp.Value;
                        }
                    }
                    
                    cleanedData.Add(cleanedRow);
                }
                
                return cleanedData;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error cleaning schedule data: {ex.Message}");
                return originalData; // Return original data if cleaning fails
            }
        }

        private object ParseNumericValue(string cellText)
        {
            try
            {
                // Clean the text to extract numeric values
                string cleanText = cellText;
                
                // Remove surrounding quotes
                if (cleanText.StartsWith("\"") && cleanText.EndsWith("\""))
                {
                    cleanText = cleanText.Substring(1, cleanText.Length - 2);
                }
                
                // Remove common unit symbols and trailing quotes
                cleanText = cleanText.Replace("\"\"", ""); // Remove double quotes at end (inches symbol)
                cleanText = cleanText.Replace("\"", "");   // Remove single quotes
                cleanText = cleanText.Replace("'", "");    // Remove feet symbol
                cleanText = cleanText.Replace("°", "");    // Remove degree symbol
                cleanText = cleanText.Replace("mm", "");   // Remove millimeters
                cleanText = cleanText.Replace("cm", "");   // Remove centimeters
                cleanText = cleanText.Replace("m", "");    // Remove meters (be careful with this one)
                cleanText = cleanText.Replace("in", "");   // Remove inches
                cleanText = cleanText.Replace("ft", "");   // Remove feet
                
                // Trim whitespace
                cleanText = cleanText.Trim();
                
                // Try to parse as double first (handles decimals)
                if (double.TryParse(cleanText, out double doubleValue))
                {
                    // Return as double for decimal values, int for whole numbers
                    if (doubleValue == Math.Floor(doubleValue) && doubleValue <= int.MaxValue && doubleValue >= int.MinValue)
                    {
                        return (int)doubleValue;
                    }
                    return doubleValue;
                }
                
                // If numeric parsing fails, return original text
                return cellText;
            }
            catch (Exception)
            {
                // If any error occurs, return original text
                return cellText;
            }
        }

        private string ShowOutputPathDialog()
        {
            // Create a TaskDialog for path input
            TaskDialog pathDialog = new TaskDialog("Select Output Folder");
            pathDialog.MainInstruction = "Choose Output Folder for Schedule Export";
            pathDialog.MainContent = "Enter the path to the folder where you want to save the exported JSON files.\n\n" +
                                   "📁 How to get folder path:\n" +
                                   "1. Create a new folder on your Desktop (or anywhere)\n" +
                                   "2. Open the folder in File Explorer\n" +
                                   "3. Click in the address bar at the top (where it shows the path)\n" +
                                   "4. Copy the full path (Ctrl+C)\n" +
                                   "5. Paste it in the input below\n\n" +
                                   "Example: C:\\Users\\YourName\\Desktop\\MyScheduleExports\n\n" +
                                   "Leave empty to use default location (next to Revit project file).";

            pathDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Browse for Folder", "Open folder browser dialog");
            pathDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Enter Path Manually", "Type the folder path");
            pathDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Use Default Location", "Save next to Revit project file");

            TaskDialogResult result = pathDialog.Show();

            switch (result)
            {
                case TaskDialogResult.CommandLink1:
                    return ShowFolderBrowserDialog();
                case TaskDialogResult.CommandLink2:
                    return ShowTextInputDialog();
                case TaskDialogResult.CommandLink3:
                default:
                    return null; // Use default location
            }
        }

        private string ShowFolderBrowserDialog()
        {
            try
            {
                // Use Windows Forms FolderBrowserDialog
                using (var folderDialog = new System.Windows.Forms.FolderBrowserDialog())
                {
                    folderDialog.Description = "Select folder for schedule exports";
                    folderDialog.ShowNewFolderButton = true;
                    
                    if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        return folderDialog.SelectedPath;
                    }
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Could not open folder browser: {ex.Message}");
            }
            
            return null;
        }

        private string ShowTextInputDialog()
        {
            // Create a simple text input dialog using TaskDialog
            TaskDialog inputDialog = new TaskDialog("Enter Folder Path");
            inputDialog.MainInstruction = "Enter the full path to your output folder";
            inputDialog.MainContent = "Paste the folder path here:\n\n" +
                                    "Example: C:\\Users\\YourName\\Desktop\\MyScheduleExports\n\n" +
                                    "📋 Tip: Right-click and paste (Ctrl+V) if you copied from File Explorer";
            
            inputDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Use Desktop", "Save to Desktop/ScheduleExports");
            inputDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Use Documents", "Save to Documents/ScheduleExports"); 
            inputDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Cancel", "Use default location");

            TaskDialogResult result = inputDialog.Show();

            switch (result)
            {
                case TaskDialogResult.CommandLink1:
                    return Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                case TaskDialogResult.CommandLink2:
                    return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                case TaskDialogResult.CommandLink3:
                default:
                    return null;
            }
        }
    }
}
