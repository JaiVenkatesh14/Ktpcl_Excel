namespace Kptcl_Excel_Load.Services
{
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

            public async Task<bool> TestConnectionAsync()
            {
                var connectionString =
                    _configuration.GetConnectionString("OracleConnection");

                if (string.IsNullOrWhiteSpace(connectionString))
                {
                    throw new Exception(
                        "Oracle connection string is not configured.");
                }

                using var connection =
                    new OracleConnection(connectionString);

                await connection.OpenAsync();

                return connection.State ==
                       System.Data.ConnectionState.Open;
            }
        }
    }
}
