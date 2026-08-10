namespace KPTCL_API_STAGG.Models
{
    public class ExcelImportResult
    {
        public bool Success { get; set; }

        public string Message { get; set; } = string.Empty;

        public List<ImportError> Errors { get; set; } = new();

        public List<WorkOrderImportModel> WorkOrders { get; set; } = new();

        public List<StationImportModel> Stations { get; set; } = new();

        public List<LineCodeImportModel> Lines { get; set; } = new();

        public void AddError(
            string sheet,
            int row,
            string message)
        {
            Errors.Add(new ImportError
            {
                Sheet = sheet,
                Row = row,
                Message = message
            });
        }
    }

    public class ImportError
    {
        public string Sheet { get; set; } = string.Empty;

        public int Row { get; set; }

        public string Message { get; set; } = string.Empty;
    }
}