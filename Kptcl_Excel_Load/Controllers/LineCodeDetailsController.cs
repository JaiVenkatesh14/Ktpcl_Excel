using Microsoft.AspNetCore.Mvc;
using Kptcl_Excel_Load.Services;

namespace Kptcl_Excel_Load.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class LineCodeDetailsController : ControllerBase
    {
        private readonly OracleDbService _oracleDbService;

        public LineCodeDetailsController(OracleDbService oracleDbService)
        {
            _oracleDbService = oracleDbService;
        }

        [HttpGet]
        public async Task<IActionResult> GetLineCodeDetails()
        {
            try
            {
                var data = await _oracleDbService.GetLineCodeDetailsAsync();

                return Ok(new
                {
                    success = true,
                    count = data.Count,
                    data = data
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new
                {
                    success = false,
                    message = ex.Message
                });
            }
        }
    }
}