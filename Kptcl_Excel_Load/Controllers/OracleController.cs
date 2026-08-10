using Kptcl_Excel_Load.Services;
using Kptcl_Excel_Load.Services.Kptcl_Excel_Load.Services;
using Microsoft.AspNetCore.Mvc;

namespace Kptcl_Excel_Load.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class OracleController : ControllerBase
    {
        private readonly OracleDbService _oracleDbService;

        public OracleController(OracleDbService oracleDbService)
        {
            _oracleDbService = oracleDbService;
        }

        [HttpGet("test-connection")]
        public async Task<IActionResult> TestConnection()
        {
            try
            {
                var connected =
                    await _oracleDbService.TestConnectionAsync();

                if (connected)
                {
                    return Ok(new
                    {
                        success = true,
                        message = "Oracle connection successful."
                    });
                }

                return StatusCode(500, new
                {
                    success = false,
                    message = "Oracle connection failed."
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new
                {
                    success = false,
                    message = "Oracle connection failed.",
                    error = ex.Message
                });
            }
        }
    }
}
