using Kptcl_Excel_Load.Models;
using Oracle.ManagedDataAccess.Client;

namespace Kptcl_Excel_Load.Services
{
    public class OracleDbService
    {
        private readonly IConfiguration _configuration;

        public OracleDbService(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        private OracleConnection GetConnection()
        {
            string connectionString =
                _configuration.GetConnectionString("OracleConnection");

            return new OracleConnection(connectionString);
        }


        // ==========================================
        // TEST ORACLE CONNECTION
        // ==========================================

        public async Task<bool> TestConnectionAsync()
        {
            using var connection = GetConnection();

            await connection.OpenAsync();

            return connection.State ==
                   System.Data.ConnectionState.Open;
        }


        // ==========================================
        // GET WORK ORDERS
        // ==========================================

        public async Task<List<WorkOrder>> GetWorkOrdersAsync()
        {
            var result = new List<WorkOrder>();

            using var connection = GetConnection();

            await connection.OpenAsync();

            using var command = new OracleCommand(
                OracleQueries.GetWorkOrders,
                connection);

            using var reader =
                await command.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                result.Add(new WorkOrder
                {
                    WORK_CODE =
                        reader["WORK_CODE"]?.ToString(),

                    NAME_OF_WORK =
                        reader["NAME_OF_WORK"]?.ToString(),

                    SUB_WORKCODE =
                        reader["SUB_WORKCODE"]?.ToString(),

                    NAME_OF_SUBWORK =
                        reader["NAME_OF_SUBWORK"]?.ToString(),

                    USER_NAME =
                        reader["USER_NAME"]?.ToString(),

                    STATUS =
                        reader["STATUS"] == DBNull.Value
                            ? null
                            : Convert.ToInt32(reader["STATUS"])
                });
            }

            return result;
        }


        // ==========================================
        // GET STATION DETAILS
        // ==========================================

        public async Task<List<StationDetails>> GetStationDetailsAsync()
        {
            var result = new List<StationDetails>();

            using var connection = GetConnection();

            await connection.OpenAsync();

            using var command = new OracleCommand(
                OracleQueries.GetStationDetails,
                connection);

            using var reader =
                await command.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                result.Add(new StationDetails
                {
                    ZONE =
                        reader["ZONE"]?.ToString(),

                    CIRCLE =
                        reader["CIRCLE"]?.ToString(),

                    DIVISION =
                        reader["DIVISION"]?.ToString(),

                    VOLTAGE_CLASS =
                        reader["VOLTAGE_CLASS"]?.ToString(),

                    SUBSTATION_NAME =
                        reader["SUBSTATION_NAME"]?.ToString(),

                    STATION_CODE =
                        reader["STATION_CODE"]?.ToString(),

                    TYPE_OF_SUBSTATION =
                        reader["TYPE_OF_SUBSTATION"]?.ToString(),

                    DATE_OF_COMMISSIONING =
                        reader["DATE_OF_COMMISSIONING"]?.ToString(),

                    WORK_CODE =
                        reader["WORK_CODE"]?.ToString()
                });
            }

            return result;
        }


        // ==========================================
        // GET LINE CODE DETAILS
        // ==========================================

        public async Task<List<LineCodeDetails>> GetLineCodeDetailsAsync()
        {
            var result = new List<LineCodeDetails>();

            using var connection = GetConnection();

            await connection.OpenAsync();

            using var command = new OracleCommand(
                OracleQueries.GetLineCodeDetails,
                connection);

            using var reader =
                await command.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                result.Add(new LineCodeDetails
                {
                    ZONE =
                        reader["ZONE"]?.ToString(),

                    CIRCLE =
                        reader["CIRCLE"]?.ToString(),

                    DIVISION =
                        reader["DIVISION"]?.ToString(),

                    VOLTAGECLASS =
                        reader["VOLTAGECLASS"]?.ToString(),

                    LINE_CODE =
                        reader["LINE_CODE"]?.ToString(),

                    LINE_NAME =
                        reader["LINE_NAME"]?.ToString(),

                    ZONE_CODE =
                        reader["ZONE_CODE"]?.ToString(),

                    LINE_TYPE =
                        reader["LINE_TYPE"]?.ToString(),

                    WORK_CODE =
                        reader["WORK_CODE"]?.ToString()
                });
            }

            return result;
        }
    }
}