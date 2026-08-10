namespace Kptcl_Excel_Load.Models
{
    public class WorkOrder
    {
        public string WORK_CODE { get; set; }
        public string NAME_OF_WORK { get; set; }
        public string SUB_WORKCODE { get; set; }
        public string NAME_OF_SUBWORK { get; set; }
        public string USER_NAME { get; set; }
        public int? STATUS { get; set; }
    }
}