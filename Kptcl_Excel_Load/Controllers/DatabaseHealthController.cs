using Microsoft.AspNetCore.Mvc;
using Oracle.ManagedDataAccess.Client;

namespace Kptcl_Excel_Load.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DatabaseHealthController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public DatabaseHealthController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public async Task<IActionResult> CheckDatabase()
        {
            try
            {
                var connectionString =
                    _configuration.GetConnectionString("OracleConnection");

                if (string.IsNullOrWhiteSpace(connectionString))
                {
                    return StatusCode(500, new
                    {
                        success = false,
                        connected = false,
                        message = "Oracle connection string is not configured."
                    });
                }

                await using var connection =
                    new OracleConnection(connectionString);

                await connection.OpenAsync();

                await connection.CloseAsync();

                return Ok(new
                {
                    success = true,
                    connected = true,
                    message = "Oracle Database Connected"
                });
            }
            catch (Exception ex)
            {
                return StatusCode(503, new
                {
                    success = false,
                    connected = false,
                    message = "Oracle Database Disconnected",
                    error = ex.Message
                });
            }
        }
    }
}