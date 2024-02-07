using Smartsheet.Api.Models;

namespace IndiaEventsWebApi.Helper
{    
    public static class SheetHelper
    {
        public static long GetColumnIdByName(Sheet sheet,string columnName)
        {
            foreach (var column in sheet.Columns)
            {
                if (column.Title == columnName)
                {
                    return column.Id.Value;
                }
            }
            return 0;
        }
        

    }
}
