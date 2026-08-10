namespace KPTCL_API_STAGG.Models
{
    public class WorkOrderImportModel
    {
        public string? WorkCode { get; set; }

        public string? NameOfWork { get; set; }

        public string? SubWorkCode { get; set; }

        public string? NameOfSubWork { get; set; }

        public string? UserName { get; set; }

        public int Status { get; set; }
    }
}