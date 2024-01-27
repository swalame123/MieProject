using IndiaEventsWebApi.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;

namespace IndiaEventsWebApi.Controllers.HCPMaster
{
    [Route("api/[controller]")]
    [ApiController]
    public class HCPController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public HCPController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }

        [HttpPost("PostHcpData")]
        public IActionResult PostHcpData(AllObjModels formDataList)
        {

            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string sheetId1 = configuration.GetSection("SmartsheetSettings:HcpMaster").Value;
            string sheetId2 = configuration.GetSection("SmartsheetSettings:HcpMaster1").Value;
            string sheetId3 = configuration.GetSection("SmartsheetSettings:HcpMaster2").Value;
            string sheetId4 = configuration.GetSection("SmartsheetSettings:HcpMaster3").Value;
            string sheetId5 = configuration.GetSection("SmartsheetSettings:HcpMaster4").Value;

            long.TryParse(sheetId1, out long parsedSheetId1);
            long.TryParse(sheetId2, out long parsedSheetId2);
            long.TryParse(sheetId3, out long parsedSheetId3);
            long.TryParse(sheetId4, out long parsedSheetId4);
            long.TryParse(sheetId5, out long parsedSheetId5);

            Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);
            Sheet sheet2 = smartsheet.SheetResources.GetSheet(parsedSheetId2, null, null, null, null, null, null, null);
            Sheet sheet3 = smartsheet.SheetResources.GetSheet(parsedSheetId3, null, null, null, null, null, null, null);
            Sheet sheet4 = smartsheet.SheetResources.GetSheet(parsedSheetId4, null, null, null, null, null, null, null);
            Sheet sheet5 = smartsheet.SheetResources.GetSheet(parsedSheetId5, null, null, null, null, null, null, null);





            return Ok(sheet1);
        }

    }
}
