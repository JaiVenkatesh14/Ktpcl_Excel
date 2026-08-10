using KPTCL_API_STAGG.Services;
using Microsoft.AspNetCore.Mvc;

namespace KPTCL_API_STAGG.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ExcelImportController : ControllerBase
    {
        private readonly ExcelImportService _excelImportService;

        public ExcelImportController(
            ExcelImportService excelImportService)
        {
            _excelImportService = excelImportService;
        }

        [HttpPost("validate")]
        [Consumes("multipart/form-data")]
        public async Task<IActionResult> ValidateExcel(
            IFormFile file)
        {
            if (file == null)
            {
                return BadRequest(new
                {
                    success = false,
                    message = "Please select an Excel file."
                });
            }

            var result =
                await _excelImportService.ImportExcelAsync(file);

            if (!result.Success)
            {
                return BadRequest(result);
            }

            return Ok(result);
        }
    }
}