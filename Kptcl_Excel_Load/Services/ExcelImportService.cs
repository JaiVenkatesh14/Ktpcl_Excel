using ClosedXML.Excel;
using KPTCL_API_STAGG.Models;
using Oracle.ManagedDataAccess.Client;

namespace KPTCL_API_STAGG.Services
{
    public class ExcelImportService
    {
        private readonly IConfiguration _configuration;

        public ExcelImportService(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        // =========================================================
        // MAIN IMPORT METHOD
        // =========================================================

        public async Task<ExcelImportResult> ImportExcelAsync(IFormFile file)
        {
            var result = new ExcelImportResult
            {
                Success = true
            };

            // =====================================================
            // FILE VALIDATION
            // =====================================================

            if (file == null || file.Length == 0)
            {
                result.Success = false;
                result.Message = "Please select an Excel file.";
                return result;
            }

            var extension = Path.GetExtension(file.FileName);

            if (!extension.Equals(
                    ".xlsx",
                    StringComparison.OrdinalIgnoreCase))
            {
                result.Success = false;
                result.Message =
                    "Only .xlsx Excel files are supported.";

                return result;
            }

            const long maxFileSize = 10 * 1024 * 1024;

            if (file.Length > maxFileSize)
            {
                result.Success = false;
                result.Message =
                    "Excel file size cannot exceed 10 MB.";

                return result;
            }

            // =====================================================
            // READ EXCEL
            // =====================================================

            using var stream = new MemoryStream();

            await file.CopyToAsync(stream);

            stream.Position = 0;

            using var workbook = new XLWorkbook(stream);

            // =====================================================
            // REQUIRED SHEETS
            // =====================================================

            var requiredSheets = new[]
            {
                "Combined_work_details",
                "Station_details",
                "Line_names_with_line_codes"
            };

            foreach (var sheetName in requiredSheets)
            {
                if (!workbook.Worksheets.Contains(sheetName))
                {
                    result.Success = false;

                    result.Message =
                        $"Required sheet '{sheetName}' was not found.";

                    return result;
                }
            }

            var combinedSheet =
                workbook.Worksheet("Combined_work_details");

            var stationSheet =
                workbook.Worksheet("Station_details");

            var lineSheet =
                workbook.Worksheet("Line_names_with_line_codes");

            // =====================================================
            // 1. EXTRACT WORK ORDERS
            // =====================================================

            var workOrders =
                ExtractWorkOrders(
                    combinedSheet,
                    result);

            // =====================================================
            // 2. BUILD STATION -> SUB WORK MAPPING
            //
            // Station Code -> one or more Sub Work Codes
            // =====================================================

            var stationWorkCodeMap =
                BuildStationWorkCodeMap(
                    combinedSheet,
                    result);

            // =====================================================
            // 3. BUILD LINE -> SUB WORK MAPPING
            //
            // Line Code -> one or more Sub Work Codes
            // =====================================================

            var lineWorkCodeMap =
                BuildLineWorkCodeMap(
                    combinedSheet,
                    result);

            // =====================================================
            // 4. EXTRACT ALL STATIONS
            //
            // IMPORTANT:
            // Do NOT filter stations.
            //
            // Every row in Station_details will be loaded.
            //
            // Referenced stations receive their Sub Work Code.
            // Non-referenced stations receive NULL.
            // =====================================================

            var stations =
                ExtractStations(
                    stationSheet,
                    stationWorkCodeMap,
                    result);

            // =====================================================
            // 5. EXTRACT ALL LINES
            //
            // IMPORTANT:
            // Do NOT filter lines.
            //
            // Every row in Line_names_with_line_codes will be loaded.
            //
            // Referenced lines receive their Sub Work Code.
            // Non-referenced lines receive NULL.
            // =====================================================

            var lines =
                ExtractLines(
                    lineSheet,
                    lineWorkCodeMap,
                    result);

            // =====================================================
            // 6. EXCEL VALIDATION
            //
            // Validate that every Station/Line referenced from
            // Combined_work_details exists in the corresponding
            // Excel sheet.
            //
            // Oracle is NOT checked here.
            // =====================================================

            ValidateWorkOrderReferences(
                combinedSheet,
                stationSheet,
                lineSheet,
                result);

            if (result.Errors.Count > 0)
            {
                result.Success = false;
                result.Message =
                    "Excel validation failed.";

                result.WorkOrders = workOrders;
                result.Stations = stations;
                result.Lines = lines;

                return result;
            }

            // =====================================================
            // 7. DUPLICATE VALIDATION
            // =====================================================

            ValidateExcelDuplicates(
                workOrders,
                stations,
                lines,
                result);

            if (result.Errors.Count > 0)
            {
                result.Success = false;

                result.Message =
                    "Duplicate records were found in the Excel file.";

                result.WorkOrders = workOrders;
                result.Stations = stations;
                result.Lines = lines;

                return result;
            }

            // =====================================================
            // STORE RESULT DATA
            // =====================================================

            result.WorkOrders = workOrders;
            result.Stations = stations;
            result.Lines = lines;

            // =====================================================
            // 8. INSERT EVERYTHING INTO ORACLE
            // =====================================================

            try
            {
                await InsertIntoOracleAsync(
                    workOrders,
                    stations,
                    lines);
            }
            catch (Exception ex)
            {
                result.Success = false;

                result.Message =
                    "Excel validation succeeded, " +
                    "but database insertion failed.";

                result.Errors.Add(
                    new ImportError
                    {
                        Sheet = "Oracle Database",
                        Row = 0,
                        Message = ex.Message
                    });

                return result;
            }

            // =====================================================
            // SUCCESS
            // =====================================================

            result.Success = true;

            result.Message =
                "Excel data validated and inserted successfully into Oracle.";

            return result;
        }


        // =========================================================
        // WORK ORDERS
        // =========================================================

        private List<WorkOrderImportModel> ExtractWorkOrders(
            IXLWorksheet worksheet,
            ExcelImportResult result)
        {
            var records =
                new List<WorkOrderImportModel>();

            var rows =
                worksheet
                    .RowsUsed()
                    .Skip(1);

            string? currentWorkCode = null;
            string? currentWorkName = null;

            foreach (var row in rows)
            {
                var workCode =
                    GetString(row.Cell(4));

                var workName =
                    GetString(row.Cell(6));

                var subWorkCode =
                    GetString(row.Cell(7));

                var subWorkName =
                    GetString(row.Cell(8));

                // -------------------------------------------------
                // Ignore section/header rows
                // -------------------------------------------------

                if (string.IsNullOrWhiteSpace(subWorkCode))
                {
                    continue;
                }

                // -------------------------------------------------
                // Parent Work Code
                // -------------------------------------------------

                if (!string.IsNullOrWhiteSpace(workCode))
                {
                    currentWorkCode =
                        workCode.Trim();

                    currentWorkName =
                        workName?.Trim();
                }

                // -------------------------------------------------
                // Work Code required
                // -------------------------------------------------

                if (string.IsNullOrWhiteSpace(currentWorkCode))
                {
                    result.AddError(
                        "Combined_work_details",
                        row.RowNumber(),
                        $"Work Code is missing for Sub Work Code '{subWorkCode}'."
                    );

                    continue;
                }

                // -------------------------------------------------
                // Sub Work Code
                // -------------------------------------------------

                if (string.IsNullOrWhiteSpace(subWorkCode))
                {
                    result.AddError(
                        "Combined_work_details",
                        row.RowNumber(),
                        "Sub Work Code is missing."
                    );

                    continue;
                }

                subWorkCode =
                    subWorkCode.Trim();

                // -------------------------------------------------
                // Sub Work Name
                // -------------------------------------------------

                if (string.IsNullOrWhiteSpace(subWorkName))
                {
                    result.AddError(
                        "Combined_work_details",
                        row.RowNumber(),
                        $"Sub Work Name is missing for Sub Work Code '{subWorkCode}'."
                    );

                    continue;
                }

                subWorkName =
                    subWorkName.Trim();

                // -------------------------------------------------
                // Length validation
                // -------------------------------------------------

                if (currentWorkCode.Length > 100)
                {
                    result.AddError(
                        "Combined_work_details",
                        row.RowNumber(),
                        "Work Code exceeds 100 characters."
                    );

                    continue;
                }

                if (currentWorkName != null &&
                    currentWorkName.Length > 500)
                {
                    result.AddError(
                        "Combined_work_details",
                        row.RowNumber(),
                        "Name of Work exceeds 500 characters."
                    );

                    continue;
                }

                if (subWorkCode.Length > 100)
                {
                    result.AddError(
                        "Combined_work_details",
                        row.RowNumber(),
                        "Sub Work Code exceeds 100 characters."
                    );

                    continue;
                }

                if (subWorkName.Length > 500)
                {
                    result.AddError(
                        "Combined_work_details",
                        row.RowNumber(),
                        "Name of Subwork exceeds 500 characters."
                    );

                    continue;
                }

                // -------------------------------------------------
                // Survey Users 1 - 9
                //
                // Columns 11 - 19
                // -------------------------------------------------

                var users =
                    new List<string>();

                for (int i = 11; i <= 19; i++)
                {
                    var user =
                        GetString(row.Cell(i));

                    if (!string.IsNullOrWhiteSpace(user))
                    {
                        users.Add(user.Trim());
                    }
                }

                // -------------------------------------------------
                // At least one user
                // -------------------------------------------------

                if (users.Count == 0)
                {
                    result.AddError(
                        "Combined_work_details",
                        row.RowNumber(),
                        $"No Survey User found for Sub Work Code '{subWorkCode}'."
                    );

                    continue;
                }

                // -------------------------------------------------
                // One WORK_ORDER record per user
                // -------------------------------------------------

                foreach (var user in users)
                {
                    if (user.Length > 100)
                    {
                        result.AddError(
                            "Combined_work_details",
                            row.RowNumber(),
                            $"Survey User '{user}' exceeds 100 characters."
                        );

                        continue;
                    }

                    records.Add(
                        new WorkOrderImportModel
                        {
                            WorkCode =
                                currentWorkCode,

                            NameOfWork =
                                currentWorkName,

                            SubWorkCode =
                                subWorkCode,

                            NameOfSubWork =
                                subWorkName,

                            UserName =
                                user,

                            // STATUS intentionally NULL
                            Status = 0
                        });
                }
            }

            return records;
        }


        // =========================================================
        // BUILD STATION MAPPING
        //
        // Station Code -> Sub Work Code(s)
        // =========================================================

        private Dictionary<string, List<string>>
            BuildStationWorkCodeMap(
                IXLWorksheet worksheet,
                ExcelImportResult result)
        {
            var map =
                new Dictionary<string, List<string>>(
                    StringComparer.OrdinalIgnoreCase);

            var rows =
                worksheet
                    .RowsUsed()
                    .Skip(1);

            foreach (var row in rows)
            {
                var subWorkCode =
                    GetString(row.Cell(7));

                var stationCode =
                    GetString(row.Cell(9));

                var lineCode =
                    GetString(row.Cell(10));

                // -------------------------------------------------
                // Ignore section/header rows
                // -------------------------------------------------

                if (string.IsNullOrWhiteSpace(subWorkCode))
                {
                    continue;
                }

                subWorkCode =
                    subWorkCode.Trim();

                // -------------------------------------------------
                // Station + Line cannot both be empty
                // -------------------------------------------------

                if (string.IsNullOrWhiteSpace(stationCode) &&
                    string.IsNullOrWhiteSpace(lineCode))
                {
                    result.AddError(
                        "Combined_work_details",
                        row.RowNumber(),
                        $"Sub Work '{subWorkCode}' must have either " +
                        "a Station Code or a Line Code."
                    );

                    continue;
                }

                // -------------------------------------------------
                // Station mapping
                // -------------------------------------------------

                if (!string.IsNullOrWhiteSpace(stationCode))
                {
                    stationCode =
                        stationCode.Trim();

                    if (!map.ContainsKey(stationCode))
                    {
                        map[stationCode] =
                            new List<string>();
                    }

                    if (!map[stationCode].Any(
                            x => string.Equals(
                                x,
                                subWorkCode,
                                StringComparison.OrdinalIgnoreCase)))
                    {
                        map[stationCode]
                            .Add(subWorkCode);
                    }
                }
            }

            return map;
        }


        // =========================================================
        // BUILD LINE MAPPING
        //
        // Line Code -> Sub Work Code(s)
        // =========================================================

        private Dictionary<string, List<string>>
            BuildLineWorkCodeMap(
                IXLWorksheet worksheet,
                ExcelImportResult result)
        {
            var map =
                new Dictionary<string, List<string>>(
                    StringComparer.OrdinalIgnoreCase);

            var rows =
                worksheet
                    .RowsUsed()
                    .Skip(1);

            foreach (var row in rows)
            {
                var subWorkCode =
                    GetString(row.Cell(7));

                var stationCode =
                    GetString(row.Cell(9));

                var lineCode =
                    GetString(row.Cell(10));

                // -------------------------------------------------
                // Ignore section/header rows
                // -------------------------------------------------

                if (string.IsNullOrWhiteSpace(subWorkCode))
                {
                    continue;
                }

                subWorkCode =
                    subWorkCode.Trim();

                // -------------------------------------------------
                // Station + Line cannot both be empty
                // -------------------------------------------------

                if (string.IsNullOrWhiteSpace(stationCode) &&
                    string.IsNullOrWhiteSpace(lineCode))
                {
                    continue;
                }

                // -------------------------------------------------
                // Line mapping
                // -------------------------------------------------

                if (!string.IsNullOrWhiteSpace(lineCode))
                {
                    lineCode =
                        lineCode.Trim();

                    if (!map.ContainsKey(lineCode))
                    {
                        map[lineCode] =
                            new List<string>();
                    }

                    if (!map[lineCode].Any(
                            x => string.Equals(
                                x,
                                subWorkCode,
                                StringComparison.OrdinalIgnoreCase)))
                    {
                        map[lineCode]
                            .Add(subWorkCode);
                    }
                }
            }

            return map;
        }


        // =========================================================
        // EXTRACT ALL STATIONS
        //
        // IMPORTANT:
        // Every station in Station_details is returned.
        //
        // If referenced from Work Details:
        //     WORK_CODE = Sub Work Code
        //
        // If not referenced:
        //     WORK_CODE = NULL
        // =========================================================

        private List<StationImportModel> ExtractStations(
            IXLWorksheet worksheet,
            Dictionary<string, List<string>> stationWorkCodeMap,
            ExcelImportResult result)
        {
            var records =
                new List<StationImportModel>();

            var rows =
                worksheet
                    .RowsUsed()
                    .Skip(1);

            foreach (var row in rows)
            {
                var stationCode =
                    GetString(row.Cell(6));

                // Empty row
                if (string.IsNullOrWhiteSpace(stationCode))
                {
                    continue;
                }

                stationCode =
                    stationCode.Trim();

                if (stationCode.Length > 100)
                {
                    result.AddError(
                        "Station_details",
                        row.RowNumber(),
                        $"Station Code '{stationCode}' exceeds 100 characters."
                    );

                    continue;
                }

                var stationName =
                    GetString(row.Cell(5));

                if (stationName != null &&
                    stationName.Length > 500)
                {
                    result.AddError(
                        "Station_details",
                        row.RowNumber(),
                        $"Substation Name for '{stationCode}' exceeds 500 characters."
                    );

                    continue;
                }

                string? workCode = null;

                // -------------------------------------------------
                // Apply Sub Work Code only when this station is
                // referenced in Combined_work_details.
                // -------------------------------------------------

                if (stationWorkCodeMap.TryGetValue(
                        stationCode,
                        out var subWorkCodes))
                {
                    workCode =
                        string.Join(
                            ", ",
                            subWorkCodes);
                }

                records.Add(
                    new StationImportModel
                    {
                        Zone =
                            GetString(row.Cell(1)),

                        Circle =
                            GetString(row.Cell(2)),

                        Division =
                            GetString(row.Cell(3)),

                        VoltageClass =
                            GetString(row.Cell(4)),

                        SubstationName =
                            stationName,

                        StationCode =
                            stationCode,

                        TypeOfSubstation =
                            GetString(row.Cell(7)),

                        DateOfCommissioning =
                            GetString(row.Cell(8)),

                        WorkCode =
                            workCode
                    });
            }

            return records;
        }


        // =========================================================
        // EXTRACT ALL LINES
        //
        // IMPORTANT:
        // Every line in Line_names_with_line_codes is returned.
        //
        // If referenced from Work Details:
        //     WORK_CODE = Sub Work Code
        //
        // If not referenced:
        //     WORK_CODE = NULL
        // =========================================================

        private List<LineCodeImportModel> ExtractLines(
            IXLWorksheet worksheet,
            Dictionary<string, List<string>> lineWorkCodeMap,
            ExcelImportResult result)
        {
            var records =
                new List<LineCodeImportModel>();

            var rows =
                worksheet
                    .RowsUsed()
                    .Skip(1);

            foreach (var row in rows)
            {
                var lineCode =
                    GetString(row.Cell(6));

                // Empty row
                if (string.IsNullOrWhiteSpace(lineCode))
                {
                    continue;
                }

                lineCode =
                    lineCode.Trim();

                if (lineCode.Length > 100)
                {
                    result.AddError(
                        "Line_names_with_line_codes",
                        row.RowNumber(),
                        $"Line Code '{lineCode}' exceeds 100 characters."
                    );

                    continue;
                }

                var lineName =
                    GetString(row.Cell(7));

                if (lineName != null &&
                    lineName.Length > 500)
                {
                    result.AddError(
                        "Line_names_with_line_codes",
                        row.RowNumber(),
                        $"Line Name for '{lineCode}' exceeds 500 characters."
                    );

                    continue;
                }

                string? workCode = null;

                // -------------------------------------------------
                // Apply Sub Work Code only when this line is
                // referenced in Combined_work_details.
                // -------------------------------------------------

                if (lineWorkCodeMap.TryGetValue(
                        lineCode,
                        out var subWorkCodes))
                {
                    workCode =
                        string.Join(
                            ", ",
                            subWorkCodes);
                }

                records.Add(
                    new LineCodeImportModel
                    {
                        Zone =
                            GetString(row.Cell(2)),

                        Circle =
                            GetString(row.Cell(3)),

                        Division =
                            GetString(row.Cell(4)),

                        VoltageClass =
                            GetString(row.Cell(5)),

                        LineCode =
                            lineCode,

                        LineName =
                            lineName,

                        ZoneCode =
                            GetString(row.Cell(8)),

                        LineType =
                            GetString(row.Cell(9)),

                        WorkCode =
                            workCode
                    });
            }

            return records;
        }


        // =========================================================
        // VALIDATE WORK ORDER REFERENCES AGAINST EXCEL
        //
        // Station Code must exist in Station_details.
        //
        // Line Code must exist in
        // Line_names_with_line_codes.
        //
        // Oracle is NOT queried here.
        // =========================================================

        private void ValidateWorkOrderReferences(
            IXLWorksheet combinedSheet,
            IXLWorksheet stationSheet,
            IXLWorksheet lineSheet,
            ExcelImportResult result)
        {
            // =====================================================
            // BUILD STATION CODE SET
            // =====================================================

            var stationCodes =
                new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);

            foreach (var row in stationSheet.RowsUsed().Skip(1))
            {
                var stationCode =
                    GetString(row.Cell(6));

                if (!string.IsNullOrWhiteSpace(stationCode))
                {
                    stationCodes.Add(
                        stationCode.Trim());
                }
            }

            // =====================================================
            // BUILD LINE CODE SET
            // =====================================================

            var lineCodes =
                new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);

            foreach (var row in lineSheet.RowsUsed().Skip(1))
            {
                var lineCode =
                    GetString(row.Cell(6));

                if (!string.IsNullOrWhiteSpace(lineCode))
                {
                    lineCodes.Add(
                        lineCode.Trim());
                }
            }

            // =====================================================
            // CHECK WORK DETAILS
            // =====================================================

            foreach (var row in combinedSheet.RowsUsed().Skip(1))
            {
                var subWorkCode =
                    GetString(row.Cell(7));

                var stationCode =
                    GetString(row.Cell(9));

                var lineCode =
                    GetString(row.Cell(10));

                // -------------------------------------------------
                // Ignore rows without Sub Work
                // -------------------------------------------------

                if (string.IsNullOrWhiteSpace(subWorkCode))
                {
                    continue;
                }

                subWorkCode =
                    subWorkCode.Trim();

                // -------------------------------------------------
                // Both Station and Line empty
                // -------------------------------------------------

                if (string.IsNullOrWhiteSpace(stationCode) &&
                    string.IsNullOrWhiteSpace(lineCode))
                {
                    result.AddError(
                        "Combined_work_details",
                        row.RowNumber(),
                        $"Sub Work '{subWorkCode}' must have either " +
                        "a Station Code or a Line Code."
                    );

                    continue;
                }

                // -------------------------------------------------
                // Validate Station Code
                // -------------------------------------------------

                if (!string.IsNullOrWhiteSpace(stationCode))
                {
                    stationCode =
                        stationCode.Trim();

                    if (!stationCodes.Contains(
                            stationCode))
                    {
                        result.AddError(
                            "Combined_work_details",
                            row.RowNumber(),
                            $"Station Code '{stationCode}' referenced by " +
                            $"Sub Work '{subWorkCode}' does not exist in " +
                            "the Station_details Excel sheet."
                        );
                    }
                }

                // -------------------------------------------------
                // Validate Line Code
                // -------------------------------------------------

                if (!string.IsNullOrWhiteSpace(lineCode))
                {
                    lineCode =
                        lineCode.Trim();

                    if (!lineCodes.Contains(
                            lineCode))
                    {
                        result.AddError(
                            "Combined_work_details",
                            row.RowNumber(),
                            $"Line Code '{lineCode}' referenced by " +
                            $"Sub Work '{subWorkCode}' does not exist in " +
                            "the Line_names_with_line_codes Excel sheet."
                        );
                    }
                }
            }
        }


        // =========================================================
        // DUPLICATE VALIDATION
        // =========================================================

        private void ValidateExcelDuplicates(
            List<WorkOrderImportModel> workOrders,
            List<StationImportModel> stations,
            List<LineCodeImportModel> lines,
            ExcelImportResult result)
        {
            // =====================================================
            // WORK ORDER DUPLICATES
            // =====================================================

            var duplicateWorkOrders =
                workOrders
                    .GroupBy(x => new
                    {
                        WorkCode =
                            Normalize(x.WorkCode),

                        SubWorkCode =
                            Normalize(x.SubWorkCode),

                        UserName =
                            Normalize(x.UserName)
                    })
                    .Where(g => g.Count() > 1);

            foreach (var group in duplicateWorkOrders)
            {
                result.AddError(
                    "Combined_work_details",
                    0,
                    $"Duplicate Work Order found: " +
                    $"{group.Key.WorkCode} / " +
                    $"{group.Key.SubWorkCode} / " +
                    $"{group.Key.UserName}"
                );
            }

            // =====================================================
            // STATION DUPLICATES
            // =====================================================

            var duplicateStations =
                stations
                    .GroupBy(
                        x =>
                            Normalize(x.StationCode),
                        StringComparer.OrdinalIgnoreCase)
                    .Where(g => g.Count() > 1);

            foreach (var group in duplicateStations)
            {
                result.AddError(
                    "Station_details",
                    0,
                    $"Duplicate Station Code found: " +
                    $"{group.Key}"
                );
            }

            // =====================================================
            // LINE DUPLICATES
            // =====================================================

            var duplicateLines =
                lines
                    .GroupBy(
                        x =>
                            Normalize(x.LineCode),
                        StringComparer.OrdinalIgnoreCase)
                    .Where(g => g.Count() > 1);

            foreach (var group in duplicateLines)
            {
                result.AddError(
                    "Line_names_with_line_codes",
                    0,
                    $"Duplicate Line Code found: " +
                    $"{group.Key}"
                );
            }
        }


        // =========================================================
        // INSERT EVERYTHING INTO ORACLE
        //
        // WORK_ORDER
        //     INSERT
        //
        // STATION_DETAILS
        //     INSERT all rows
        //     WORK_CODE gets mapped Sub Work Code or NULL
        //
        // LINE_CODE_DETAILS
        //     INSERT all rows
        //     WORK_CODE gets mapped Sub Work Code or NULL
        // =========================================================

        private async Task InsertIntoOracleAsync(
            List<WorkOrderImportModel> workOrders,
            List<StationImportModel> stations,
            List<LineCodeImportModel> lines)
        {
            var connectionString =
                _configuration.GetConnectionString(
                    "OracleConnection");

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new Exception(
                    "Oracle connection string was not found.");
            }

            await using var connection =
                new OracleConnection(connectionString);

            await connection.OpenAsync();

            await using var transaction =
                connection.BeginTransaction();

            try
            {
                // =================================================
                // 1. WORK_ORDER
                // =================================================

                foreach (var work in workOrders)
                {
                    const string sql = @"
                        INSERT INTO SYSTEM.WORK_ORDER
                        (
                            WORK_CODE,
                            NAME_OF_WORK,
                            SUB_WORKCODE,
                            NAME_OF_SUBWORK,
                            USER_NAME,
                            STATUS
                        )
                        VALUES
                        (
                            :WORK_CODE,
                            :NAME_OF_WORK,
                            :SUB_WORKCODE,
                            :NAME_OF_SUBWORK,
                            :USER_NAME,
                            :STATUS
                        )";

                    await using var command =
    new OracleCommand(
        sql,
        connection);

                    command.BindByName = true;

                    command.Transaction =
                        transaction;
                    command.Parameters.Add(
                        ":WORK_CODE",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)work.WorkCode ??
                        DBNull.Value;

                    command.Parameters.Add(
                        ":NAME_OF_WORK",
                        OracleDbType.Varchar2,
                        500).Value =
                        (object?)work.NameOfWork ??
                        DBNull.Value;

                    command.Parameters.Add(
                        ":SUB_WORKCODE",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)work.SubWorkCode ??
                        DBNull.Value;

                    command.Parameters.Add(
                        ":NAME_OF_SUBWORK",
                        OracleDbType.Varchar2,
                        500).Value =
                        (object?)work.NameOfSubWork ??
                        DBNull.Value;

                    command.Parameters.Add(
                        ":USER_NAME",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)work.UserName ??
                        DBNull.Value;

                    // STATUS remains NULL
                    command.Parameters.Add(
                        ":STATUS",
                        OracleDbType.Int32).Value =
                        DBNull.Value;

                    await command.ExecuteNonQueryAsync();
                }


                // =================================================
                // 2. STATION_DETAILS
                //
                // INSERT ALL STATIONS
                //
                // If Station Code already exists:
                // update WORK_CODE
                //
                // Otherwise:
                // insert complete station record.
                // =================================================

                foreach (var station in stations)
                {
                    const string sql = @"
                        MERGE INTO SYSTEM.STATION_DETAILS target
                        USING
                        (
                            SELECT
                                :STATION_CODE AS STATION_CODE
                            FROM DUAL
                        ) source
                        ON
                        (
                            TRIM(UPPER(target.STATION_CODE)) =
                            TRIM(UPPER(source.STATION_CODE))
                        )
                        WHEN MATCHED THEN
                            UPDATE SET
                                target.WORK_CODE = :WORK_CODE
                        WHEN NOT MATCHED THEN
                            INSERT
                            (
                                ZONE,
                                CIRCLE,
                                DIVISION,
                                VOLTAGE_CLASS,
                                SUBSTATION_NAME,
                                STATION_CODE,
                                TYPE_OF_SUBSTATION,
                                DATE_OF_COMMISSIONING,
                                WORK_CODE
                            )
                            VALUES
                            (
                                :ZONE,
                                :CIRCLE,
                                :DIVISION,
                                :VOLTAGE_CLASS,
                                :SUBSTATION_NAME,
                                :STATION_CODE,
                                :TYPE_OF_SUBSTATION,
                                :DATE_OF_COMMISSIONING,
                                :WORK_CODE
                            )";

                    await using var command =
    new OracleCommand(
        sql,
        connection);

                    command.BindByName = true;

                    command.Transaction =
                        transaction;

                    command.Parameters.Add(
                        ":STATION_CODE",
                        OracleDbType.Varchar2,
                        100).Value =
                        station.StationCode!.Trim();

                    command.Parameters.Add(
                        ":WORK_CODE",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)station.WorkCode ??
                        DBNull.Value;

                    command.Parameters.Add(
                        ":ZONE",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)station.Zone ??
                        DBNull.Value;

                    command.Parameters.Add(
                        ":CIRCLE",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)station.Circle ??
                        DBNull.Value;

                    command.Parameters.Add(
                        ":DIVISION",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)station.Division ??
                        DBNull.Value;

                    command.Parameters.Add(
                        ":VOLTAGE_CLASS",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)station.VoltageClass ??
                        DBNull.Value;

                    command.Parameters.Add(
                        ":SUBSTATION_NAME",
                        OracleDbType.Varchar2,
                        500).Value =
                        (object?)station.SubstationName ??
                        DBNull.Value;

                    command.Parameters.Add(
                        ":TYPE_OF_SUBSTATION",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)station.TypeOfSubstation ??
                        DBNull.Value;

                    command.Parameters.Add(
                        ":DATE_OF_COMMISSIONING",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)station.DateOfCommissioning ??
                        DBNull.Value;

                    await command.ExecuteNonQueryAsync();
                }


                // =================================================
                // 3. LINE_CODE_DETAILS
                //
                // INSERT ALL LINES
                //
                // If Line Code already exists:
                // update WORK_CODE
                //
                // Otherwise:
                // insert complete line record.
                // =================================================

                foreach (var line in lines)
                {
                    const string sql = @"
                        MERGE INTO SYSTEM.LINE_CODE_DETAILS target
                        USING
                        (
                            SELECT
                                :LINE_CODE AS LINE_CODE
                            FROM DUAL
                        ) source
                        ON
                        (
                            TRIM(UPPER(target.LINE_CODE)) =
                            TRIM(UPPER(source.LINE_CODE))
                        )
                        WHEN MATCHED THEN
                            UPDATE SET
                                target.WORK_CODE = :WORK_CODE
                        WHEN NOT MATCHED THEN
                            INSERT
                            (
                                ZONE,
                                CIRCLE,
                                DIVISION,
                                VOLTAGECLASS,
                                LINE_CODE,
                                LINE_NAME,
                                ZONE_CODE,
                                LINE_TYPE,
                                WORK_CODE
                            )
                            VALUES
                            (
                                :ZONE,
                                :CIRCLE,
                                :DIVISION,
                                :VOLTAGECLASS,
                                :LINE_CODE,
                                :LINE_NAME,
                                :ZONE_CODE,
                                :LINE_TYPE,
                                :WORK_CODE
                            )";

                    await using var command =
    new OracleCommand(
        sql,
        connection);

                    command.BindByName = true;

                    command.Transaction =
                        transaction;

                    command.Parameters.Add(
                        ":LINE_CODE",
                        OracleDbType.Varchar2,
                        100).Value =
                        line.LineCode!.Trim();

                    command.Parameters.Add(
                        ":WORK_CODE",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)line.WorkCode ??
                        DBNull.Value;

                    command.Parameters.Add(
                        ":ZONE",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)line.Zone ??
                        DBNull.Value;

                    command.Parameters.Add(
                        ":CIRCLE",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)line.Circle ??
                        DBNull.Value;

                    command.Parameters.Add(
                        ":DIVISION",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)line.Division ??
                        DBNull.Value;

                    // IMPORTANT:
                    // Oracle column is VOLTAGECLASS
                    // NOT VOLTAGE_CLASS
                    command.Parameters.Add(
                        ":VOLTAGECLASS",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)line.VoltageClass ??
                        DBNull.Value;

                    command.Parameters.Add(
                        ":LINE_NAME",
                        OracleDbType.Varchar2,
                        500).Value =
                        (object?)line.LineName ??
                        DBNull.Value;

                    command.Parameters.Add(
                        ":ZONE_CODE",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)line.ZoneCode ??
                        DBNull.Value;

                    command.Parameters.Add(
                        ":LINE_TYPE",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)line.LineType ??
                        DBNull.Value;

                    await command.ExecuteNonQueryAsync();
                }


                // =================================================
                // COMMIT
                // =================================================

                await transaction.CommitAsync();
            }
            catch
            {
                // =================================================
                // ROLLBACK EVERYTHING
                // =================================================

                await transaction.RollbackAsync();

                throw;
            }
        }


        // =========================================================
        // NORMALIZE
        // =========================================================

        private static string Normalize(
            string? value)
        {
            return value?.Trim() ?? string.Empty;
        }


        // =========================================================
        // GET STRING
        // =========================================================

        private static string? GetString(
            IXLCell cell)
        {
            if (cell.IsEmpty())
            {
                return null;
            }

            var value =
                cell.GetValue<string>();

            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            return value.Trim();
        }
    }
}