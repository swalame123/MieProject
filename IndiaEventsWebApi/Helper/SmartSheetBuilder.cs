using Serilog;
using Smartsheet.Api;

namespace IndiaEventsWebApi.Helper
{
    public class SmartSheetBuilder
    {
        //private static SemaphoreSlim semaphore;
        public static SmartsheetClient AccessClient(string accessToken, SemaphoreSlim semaphore)
        {
            try
            {
                //semaphore = new SemaphoreSlim(1);
                //semaphore.Wait();
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                return smartsheet;
            }
            catch (Exception ex)
            {
                Log.Error($"Error occured on method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return (SmartsheetClient)ex;
            }
            //finally
            //{
            //    semaphore.Release();
            //}
        }
    }
}
