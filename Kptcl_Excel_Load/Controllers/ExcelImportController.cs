using KPTCL_API_STAGG.Services;
using Microsoft.AspNetCore.Http;
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

        // =====================================================
        // EXCEL UPLOAD + VALIDATION + DATABASE INSERT
        // =====================================================

        [HttpPost("validate")]
        [Consumes("multipart/form-data")]
        public async Task<IActionResult> ValidateExcel(
            IFormFile file)
        {
            // -----------------------------------------------
            // Check file
            // -----------------------------------------------

            if (file == null || file.Length == 0)
            {
                return BadRequest(new
                {
                    success = false,
                    message = "Please select an Excel file.",
                    errors = new object[0],
                    workOrders = new object[0],
                    stations = new object[0],
                    lines = new object[0]
                });
            }

            // -----------------------------------------------
            // Check extension
            // -----------------------------------------------

            var extension =
                Path.GetExtension(file.FileName);

            if (!string.Equals(
                    extension,
                    ".xlsx",
                    StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest(new
                {
                    success = false,
                    message = "Only .xlsx Excel files are supported."
                });
            }

            // -----------------------------------------------
            // Check file size
            // -----------------------------------------------

            const long maxFileSize =
                10 * 1024 * 1024; // 10 MB

            if (file.Length > maxFileSize)
            {
                return BadRequest(new
                {
                    success = false,
                    message = "Excel file size cannot exceed 10 MB."
                });
            }

            // -----------------------------------------------
            // Process Excel
            // -----------------------------------------------

            try
            {
                var result =
                    await _excelImportService
                        .ImportExcelAsync(file);

                // -------------------------------------------
                // Validation / database error
                // -------------------------------------------

                if (!result.Success)
                {
                    return BadRequest(result);
                }

                // -------------------------------------------
                // Successful import
                // -------------------------------------------

                return Ok(result);
            }
            catch (Exception ex)
            {
                // -------------------------------------------
                // Unexpected error
                // -------------------------------------------

                return StatusCode(
                    StatusCodes.Status500InternalServerError,
                    new
                    {
                        success = false,
                        message =
                            "An unexpected error occurred while importing the Excel file.",
                        error = ex.Message
                    });
            }
        }
    }
}