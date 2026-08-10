namespace Kptcl_Excel_Load.Services
{
    public static class OracleQueries
    {
        public const string GetWorkOrders = @"
            SELECT *
            FROM WORK_ORDER";

        public const string GetStationDetails = @"
            SELECT *
            FROM STATION_DETAILS";

        public const string GetLineCodeDetails = @"
            SELECT *
            FROM LINE_CODE_DETAILS";
    }
}