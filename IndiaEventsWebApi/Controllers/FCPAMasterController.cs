using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Runtime.InteropServices;

namespace IndiaEventsWebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
   
    public class FCPAMasterController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public FCPAMasterController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }
      





        [HttpGet("GetHCPDataUsingMISCode")]
        public IActionResult GetHCPDataUsingMISCode(string misCode)
        {
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string[] sheetIds = {
        //configuration.GetSection("SmartsheetSettings:HcpMaster").Value,
        //configuration.GetSection("SmartsheetSettings:HcpMaster1").Value,
        //configuration.GetSection("SmartsheetSettings:HcpMaster2").Value,
        //configuration.GetSection("SmartsheetSettings:HcpMaster3").Value,
        //configuration.GetSection("SmartsheetSettings:HcpMaster4").Value,
        configuration.GetSection("SmartsheetSettings:ApprovedSpeakers").Value,
        configuration.GetSection("SmartsheetSettings:ApprovedTrainers").Value
    };

            foreach (string i in sheetIds)
            {
                long.TryParse(i, out long p);
                Sheet sheeti = smartsheet.SheetResources.GetSheet(p, null, null, null, null, null, null, null);

                Column misCodeColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "MisCode");
                Column fcpaSignOffDateColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "FCPA Sign Off Date");
                Column fcpaExpiryDateColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "FCPA Expiry Date");
                Column fcpaValidColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "FCPA Valid?");

                //if (misCodeColumn != null)
                //{
                //    Row existingRow = sheeti.Rows.FirstOrDefault(row =>
                //        row.Cells != null &&
                //        row.Cells.Any(cell =>
                //            cell.ColumnId == misCodeColumn.Id &&
                //            cell.Value != null &&
                //            cell.Value.ToString() == misCode)
                //    );
                if (misCodeColumn != null)
                {
                    Row existingRow = sheeti.Rows.FirstOrDefault(row =>
                        row.Cells != null &&

                        row.Cells.Any(cell =>
                            cell.ColumnId == misCodeColumn.Id && cell.Value != null && cell.Value.ToString() == misCode
                        )
                    );

                    if (existingRow != null)
                    {
                       
                        Cell misCodeCell = existingRow.Cells.FirstOrDefault(cell => cell.ColumnId == misCodeColumn.Id);
                        Cell fcpaSignOffDateCell = existingRow.Cells.FirstOrDefault(cell => cell.ColumnId == fcpaSignOffDateColumn.Id);
                        Cell fcpaExpiryDateCell = existingRow.Cells.FirstOrDefault(cell => cell.ColumnId == fcpaExpiryDateColumn.Id);
                        Cell fcpaValidCell = existingRow.Cells.FirstOrDefault(cell => cell.ColumnId == fcpaValidColumn.Id);

                       
                        var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(p, existingRow.Id.Value,null);
                        var url = "";
                        foreach(var attachment in attachments.Data)
                        {
                            if (attachment != null)
                            {
                                var AID =(long) attachment.Id;
                                var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(p, AID);
                                url = file.Url;

                            }
                        }

                        return Ok(new
                        {
                            MISCode = misCodeCell?.Value,
                            FCPASignOffDate = fcpaSignOffDateCell?.Value,
                            FCPAExpiryDate = fcpaExpiryDateCell?.Value,
                            FCPAValid = fcpaValidCell?.Value,
                            Attachments = attachments,
                            Url = url

                        }); 
                    }
                }
            }

            return Ok("False");
        }
      
    }
}

