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

        public async Task<ExcelImportResult> ImportExcelAsync(IFormFile file)
        {
            var result = new ExcelImportResult
            {
                Success = true
            };

            // ==========================================
            // FILE VALIDATION
            // ==========================================

            if (file == null || file.Length == 0)
            {
                result.Success = false;
                result.Message = "Please select an Excel file.";
                return result;
            }

            var extension = Path.GetExtension(file.FileName);

            if (!extension.Equals(".xlsx",
                StringComparison.OrdinalIgnoreCase))
            {
                result.Success = false;
                result.Message = "Only .xlsx Excel files are supported.";
                return result;
            }

            const long maxFileSize = 10 * 1024 * 1024;

            if (file.Length > maxFileSize)
            {
                result.Success = false;
                result.Message = "Excel file size cannot exceed 10 MB.";
                return result;
            }


            // ==========================================
            // READ EXCEL
            // ==========================================

            using var stream = new MemoryStream();

            await file.CopyToAsync(stream);

            stream.Position = 0;

            using var workbook = new XLWorkbook(stream);


            // ==========================================
            // REQUIRED SHEETS
            // ==========================================

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


            // ==========================================
            // EXTRACT
            // ==========================================

            var workOrders = ExtractWorkOrders(
                workbook.Worksheet("Combined_work_details"),
                result);

            var stations = ExtractStations(
                workbook.Worksheet("Station_details"),
                result);

            var lines = ExtractLines(
                workbook.Worksheet("Line_names_with_line_codes"),
                result);


            // ==========================================
            // VALIDATION ERRORS
            // ==========================================

            if (result.Errors.Count > 0)
            {
                result.Success = false;

                result.Message =
                    "Excel validation failed.";

                return result;
            }


            // ==========================================
            // DUPLICATE VALIDATION INSIDE EXCEL
            // ==========================================

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

                return result;
            }

            // ==========================================
            // CHECK EXISTING ORACLE RECORDS
            // ==========================================

            await ValidateExistingOracleCodesAsync(
                stations,
                lines,
                result);


            if (result.Errors.Count > 0)
            {
                result.Success = false;

                result.Message =
                    "Excel validation failed. Some records already exist in Oracle.";

                return result;
            }
            // ==========================================
            // STORE EXTRACTED DATA
            // ==========================================

            result.WorkOrders = workOrders;
            result.Stations = stations;
            result.Lines = lines;


            // ==========================================
            // INSERT INTO ORACLE
            // ==========================================

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
                    "Excel validation succeeded, but database insertion failed.";

                result.Errors.Add(new ImportError
                {
                    Sheet = "Oracle Database",
                    Row = 0,
                    Message = ex.Message
                });

                return result;
            }


            // ==========================================
            // SUCCESS
            // ==========================================

            result.Success = true;

            result.Message =
                "Excel data validated and inserted successfully into Oracle.";

            return result;
        }


        // ==========================================
        // WORK ORDERS
        // ==========================================

        private List<WorkOrderImportModel> ExtractWorkOrders(
            IXLWorksheet worksheet,
            ExcelImportResult result)
        {
            var records = new List<WorkOrderImportModel>();

            var rows = worksheet
                .RowsUsed()
                .Skip(1);

            string? currentWorkCode = null;
            string? currentWorkName = null;

            foreach (var row in rows)
            {
                var workCode = GetString(row.Cell(4));
                var workName = GetString(row.Cell(6));

                var subWorkCode = GetString(row.Cell(7));
                var subWorkName = GetString(row.Cell(8));

                // Survey users 1-9
                var users = new List<string>();

                for (int i = 11; i <= 19; i++)
                {
                    var user = GetString(row.Cell(i));

                    if (!string.IsNullOrWhiteSpace(user))
                    {
                        users.Add(user);
                    }
                }

                // --------------------------------
                // Ignore completely empty rows
                // --------------------------------

                if (string.IsNullOrWhiteSpace(workCode) &&
                    string.IsNullOrWhiteSpace(subWorkCode) &&
                    string.IsNullOrWhiteSpace(workName) &&
                    string.IsNullOrWhiteSpace(subWorkName))
                {
                    continue;
                }

                // --------------------------------
                // New parent work
                // --------------------------------

                if (!string.IsNullOrWhiteSpace(workCode))
                {
                    currentWorkCode = workCode;
                    currentWorkName = workName;
                }

                // --------------------------------
                // Ignore rows without work/subwork
                // --------------------------------

                if (string.IsNullOrWhiteSpace(currentWorkCode))
                {
                    continue;
                }

                // --------------------------------
                // Need subwork information
                // --------------------------------

                if (string.IsNullOrWhiteSpace(subWorkCode) &&
                    string.IsNullOrWhiteSpace(subWorkName))
                {
                    continue;
                }

                // --------------------------------
                // Validation
                // --------------------------------

                if (currentWorkCode.Length > 100)
                {
                    result.AddError(
                        "Combined_work_details",
                        row.RowNumber(),
                        "Work Code exceeds 100 characters.");

                    continue;
                }

                if (currentWorkName != null &&
                    currentWorkName.Length > 500)
                {
                    result.AddError(
                        "Combined_work_details",
                        row.RowNumber(),
                        "Name of Work exceeds 500 characters.");

                    continue;
                }

                if (subWorkCode != null &&
                    subWorkCode.Length > 100)
                {
                    result.AddError(
                        "Combined_work_details",
                        row.RowNumber(),
                        "Sub Work Code exceeds 100 characters.");

                    continue;
                }

                if (subWorkName != null &&
                    subWorkName.Length > 500)
                {
                    result.AddError(
                        "Combined_work_details",
                        row.RowNumber(),
                        "Name of Subwork exceeds 500 characters.");

                    continue;
                }

                // --------------------------------
                // USER_NAME
                // --------------------------------

                var userName = users.Count > 0
                    ? string.Join(", ", users)
                    : null;

                if (userName != null &&
                    userName.Length > 100)
                {
                    result.AddError(
                        "Combined_work_details",
                        row.RowNumber(),
                        "Combined Survey User value exceeds 100 characters.");

                    continue;
                }

                records.Add(new WorkOrderImportModel
                {
                    WorkCode = currentWorkCode,
                    NameOfWork = currentWorkName,
                    SubWorkCode = subWorkCode,
                    NameOfSubWork = subWorkName,
                    UserName = userName,
                    Status = 1
                });
            }

            return records;
        }


        // ==========================================
        // STATIONS
        // ==========================================

        private List<StationImportModel> ExtractStations(
            IXLWorksheet worksheet,
            ExcelImportResult result)
        {
            var records = new List<StationImportModel>();

            var rows = worksheet
                .RowsUsed()
                .Skip(1);

            foreach (var row in rows)
            {
                var stationCode = GetString(row.Cell(6));

                // Empty row
                if (string.IsNullOrWhiteSpace(stationCode))
                    continue;

                if (stationCode.Length > 100)
                {
                    result.AddError(
                        "Station_details",
                        row.RowNumber(),
                        "Station Code exceeds 100 characters.");

                    continue;
                }

                var stationName = GetString(row.Cell(5));

                if (stationName != null &&
                    stationName.Length > 500)
                {
                    result.AddError(
                        "Station_details",
                        row.RowNumber(),
                        "Substation Name exceeds 500 characters.");

                    continue;
                }

                records.Add(new StationImportModel
                {
                    Zone = GetString(row.Cell(1)),
                    Circle = GetString(row.Cell(2)),
                    Division = GetString(row.Cell(3)),
                    VoltageClass = GetString(row.Cell(4)),
                    SubstationName = stationName,
                    StationCode = stationCode,
                    TypeOfSubstation = GetString(row.Cell(7)),
                    DateOfCommissioning =
                        GetString(row.Cell(8)),
                    WorkCode = null
                });
            }

            return records;
        }


        // ==========================================
        // LINE CODES
        // ==========================================

        private List<LineCodeImportModel> ExtractLines(
            IXLWorksheet worksheet,
            ExcelImportResult result)
        {
            var records = new List<LineCodeImportModel>();

            var rows = worksheet
                .RowsUsed()
                .Skip(1);

            foreach (var row in rows)
            {
                var lineCode = GetString(row.Cell(6));

                if (string.IsNullOrWhiteSpace(lineCode))
                    continue;

                if (lineCode.Length > 100)
                {
                    result.AddError(
                        "Line_names_with_line_codes",
                        row.RowNumber(),
                        "Line Code exceeds 100 characters.");

                    continue;
                }

                var lineName = GetString(row.Cell(7));

                if (lineName != null &&
                    lineName.Length > 500)
                {
                    result.AddError(
                        "Line_names_with_line_codes",
                        row.RowNumber(),
                        "Line Name exceeds 500 characters.");

                    continue;
                }

                records.Add(new LineCodeImportModel
                {
                    Zone = GetString(row.Cell(2)),
                    Circle = GetString(row.Cell(3)),
                    Division = GetString(row.Cell(4)),
                    VoltageClass = GetString(row.Cell(5)),
                    LineCode = lineCode,
                    LineName = lineName,
                    ZoneCode = GetString(row.Cell(8)),
                    LineType = GetString(row.Cell(9)),
                    WorkCode = null
                });
            }

            return records;
        }


        // ==========================================
        // HELPER
        // ==========================================

        private static string? GetString(IXLCell cell)
        {
            if (cell.IsEmpty())
                return null;

            var value = cell.GetValue<string>();

            if (string.IsNullOrWhiteSpace(value))
                return null;

            return value.Trim();
        }

        private void ValidateExcelDuplicates(
    List<WorkOrderImportModel> workOrders,
    List<StationImportModel> stations,
    List<LineCodeImportModel> lines,
    ExcelImportResult result)
        {
            // ==========================================
            // WORK ORDER DUPLICATES
            // ==========================================

            var duplicateWorkOrders = workOrders
                .GroupBy(x => new
                {
                    x.WorkCode,
                    x.SubWorkCode
                })
                .Where(g => g.Count() > 1);

            foreach (var group in duplicateWorkOrders)
            {
                result.AddError(
                    "Combined_work_details",
                    0,
                    $"Duplicate Work Order found: " +
                    $"{group.Key.WorkCode} / {group.Key.SubWorkCode}");
            }


            // ==========================================
            // STATION CODE DUPLICATES
            // ==========================================

            var duplicateStations = stations
                .GroupBy(x => x.StationCode)
                .Where(g => g.Count() > 1);

            foreach (var group in duplicateStations)
            {
                result.AddError(
                    "Station_details",
                    0,
                    $"Duplicate Station Code found: {group.Key}");
            }


            // ==========================================
            // LINE CODE DUPLICATES
            // ==========================================

            var duplicateLines = lines
                .GroupBy(x => x.LineCode)
                .Where(g => g.Count() > 1);

            foreach (var group in duplicateLines)
            {
                result.AddError(
                    "Line_names_with_line_codes",
                    0,
                    $"Duplicate Line Code found: {group.Key}");
            }
        }

        private async Task InsertIntoOracleAsync(
    List<WorkOrderImportModel> workOrders,
    List<StationImportModel> stations,
    List<LineCodeImportModel> lines)
        {
            var connectionString =
                _configuration.GetConnectionString("OracleConnection");

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
                // ==========================================
                // WORK_ORDER
                // ==========================================

                foreach (var work in workOrders)
                {
                    const string sql = @"
                INSERT INTO WORK_ORDER
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
                        new OracleCommand(sql, connection);

                    command.Transaction = transaction;

                    command.Parameters.Add(
                        ":WORK_CODE",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)work.WorkCode ?? DBNull.Value;

                    command.Parameters.Add(
                        ":NAME_OF_WORK",
                        OracleDbType.Varchar2,
                        500).Value =
                        (object?)work.NameOfWork ?? DBNull.Value;

                    command.Parameters.Add(
                        ":SUB_WORKCODE",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)work.SubWorkCode ?? DBNull.Value;

                    command.Parameters.Add(
                        ":NAME_OF_SUBWORK",
                        OracleDbType.Varchar2,
                        500).Value =
                        (object?)work.NameOfSubWork ?? DBNull.Value;

                    command.Parameters.Add(
                        ":USER_NAME",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)work.UserName ?? DBNull.Value;

                    command.Parameters.Add(
                        ":STATUS",
                        OracleDbType.Int32).Value =
                        work.Status;

                    await command.ExecuteNonQueryAsync();
                }


                // ==========================================
                // STATION_DETAILS
                // ==========================================

                foreach (var station in stations)
                {
                    const string sql = @"
                INSERT INTO STATION_DETAILS
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
                        new OracleCommand(sql, connection);

                    command.Transaction = transaction;

                    command.Parameters.Add(
                        ":ZONE",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)station.Zone ?? DBNull.Value;

                    command.Parameters.Add(
                        ":CIRCLE",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)station.Circle ?? DBNull.Value;

                    command.Parameters.Add(
                        ":DIVISION",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)station.Division ?? DBNull.Value;

                    command.Parameters.Add(
                        ":VOLTAGE_CLASS",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)station.VoltageClass ?? DBNull.Value;

                    command.Parameters.Add(
                        ":SUBSTATION_NAME",
                        OracleDbType.Varchar2,
                        500).Value =
                        (object?)station.SubstationName ?? DBNull.Value;

                    command.Parameters.Add(
                        ":STATION_CODE",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)station.StationCode ?? DBNull.Value;

                    command.Parameters.Add(
                        ":TYPE_OF_SUBSTATION",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)station.TypeOfSubstation ?? DBNull.Value;

                    command.Parameters.Add(
                        ":DATE_OF_COMMISSIONING",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)station.DateOfCommissioning ?? DBNull.Value;

                    // Excel doesn't provide Work Code
                    command.Parameters.Add(
                        ":WORK_CODE",
                        OracleDbType.Varchar2,
                        100).Value =
                        DBNull.Value;

                    await command.ExecuteNonQueryAsync();
                }


                // ==========================================
                // LINE_CODE_DETAILS
                // ==========================================

                foreach (var line in lines)
                {
                    const string sql = @"
                INSERT INTO LINE_CODE_DETAILS
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
                        new OracleCommand(sql, connection);

                    command.Transaction = transaction;

                    command.Parameters.Add(
                        ":ZONE",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)line.Zone ?? DBNull.Value;

                    command.Parameters.Add(
                        ":CIRCLE",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)line.Circle ?? DBNull.Value;

                    command.Parameters.Add(
                        ":DIVISION",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)line.Division ?? DBNull.Value;

                    command.Parameters.Add(
                        ":VOLTAGECLASS",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)line.VoltageClass ?? DBNull.Value;

                    command.Parameters.Add(
                        ":LINE_CODE",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)line.LineCode ?? DBNull.Value;

                    command.Parameters.Add(
                        ":LINE_NAME",
                        OracleDbType.Varchar2,
                        500).Value =
                        (object?)line.LineName ?? DBNull.Value;

                    command.Parameters.Add(
                        ":ZONE_CODE",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)line.ZoneCode ?? DBNull.Value;

                    command.Parameters.Add(
                        ":LINE_TYPE",
                        OracleDbType.Varchar2,
                        100).Value =
                        (object?)line.LineType ?? DBNull.Value;

                    // Excel doesn't provide Work Code
                    command.Parameters.Add(
                        ":WORK_CODE",
                        OracleDbType.Varchar2,
                        100).Value =
                        DBNull.Value;

                    await command.ExecuteNonQueryAsync();
                }


                // ==========================================
                // COMMIT
                // ==========================================

                await transaction.CommitAsync();
            }
            catch
            {
                // ==========================================
                // ROLLBACK
                // ==========================================

                await transaction.RollbackAsync();

                throw;
            }
        }
        private async Task ValidateExistingOracleCodesAsync(
    List<StationImportModel> stations,
    List<LineCodeImportModel> lines,
    ExcelImportResult result)
        {
            var connectionString =
                _configuration.GetConnectionString("OracleConnection");

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                result.AddError(
                    "Oracle Database",
                    0,
                    "Oracle connection string was not found.");

                return;
            }

            await using var connection =
                new OracleConnection(connectionString);

            await connection.OpenAsync();


            // ==========================================
            // CHECK STATION CODES
            // ==========================================

            foreach (var station in stations)
            {
                if (string.IsNullOrWhiteSpace(station.StationCode))
                    continue;

                const string sql = @"
            SELECT COUNT(1)
            FROM STATION_DETAILS
            WHERE STATION_CODE = :STATION_CODE";

                await using var command =
                    new OracleCommand(sql, connection);

                command.Parameters.Add(
                    ":STATION_CODE",
                    OracleDbType.Varchar2,
                    100).Value =
                    station.StationCode;

                var count =
                    Convert.ToInt32(await command.ExecuteScalarAsync());

                if (count > 0)
                {
                    result.AddError(
                        "Station_details",
                        0,
                        $"Station Code '{station.StationCode}' already exists in STATION_DETAILS.");
                }
            }


            // ==========================================
            // CHECK LINE CODES
            // ==========================================

            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line.LineCode))
                    continue;

                const string sql = @"
            SELECT COUNT(1)
            FROM LINE_CODE_DETAILS
            WHERE LINE_CODE = :LINE_CODE";

                await using var command =
                    new OracleCommand(sql, connection);

                command.Parameters.Add(
                    ":LINE_CODE",
                    OracleDbType.Varchar2,
                    100).Value =
                    line.LineCode;

                var count =
                    Convert.ToInt32(await command.ExecuteScalarAsync());

                if (count > 0)
                {
                    result.AddError(
                        "Line_names_with_line_codes",
                        0,
                        $"Line Code '{line.LineCode}' already exists in LINE_CODE_DETAILS.");
                }
            }
        }
    }
}