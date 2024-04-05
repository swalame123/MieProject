using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models.RequestSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using static Org.BouncyCastle.Bcpg.Attr.ImageAttrib;

namespace IndiaEventsWebApi.Controllers.RequestSheets
{
    [Route("api/[controller]")]
    [ApiController]
    public class AddorDeleteExpenseController : ControllerBase
    {
         private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;

        private readonly string sheetId6;
        private readonly string sheetId7;

        public AddorDeleteExpenseController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

            sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;

        }

        [HttpPost("AddPanelist"), DisableRequestSizeLimit]
        public IActionResult AddPanelist(AddNewExpense formdata)
        {
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

            try
            {
                Row newRow1 = new()
                {
                    Cells = new List<Cell>()
                };

                Row newRow6 = new()
                {
                    Cells = new List<Cell>()
                };
                newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.ExpenseType });
                newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = formdata.EventId });
                //newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "AmountExcludingTax?"), Value = formdata.ExpenseAmountExcludingTax });


                //newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "AmountExcludingTax?"), Value = formdata.AmountExcludingTax });
                newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount Excluding Tax"), Value = formdata.ExcludingTaxAmount });

                newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = formdata.AmountExcludingTax });
                newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.BtcorBte });
                //newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BudgetAmount"), Value = formdata.BudgetAmount });
                //newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = formdata.BtcAmount });
                //newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = formdata.BteAmount });
                newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formdata.EventTopic });
                newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formdata.EventType });
                newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Venue name"), Value = formdata.VenueName });
                newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formdata.EventStartDate });
                newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formdata.EventEndDate });
                smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, new Row[] { newRow6 });



                //if (formData.IsDeviationUpload == "Yes")
                //{
                //    List<string> DeviationNames = new();
                //    foreach (string p in formData.DeviationFiles)
                //    {
                //        string[] words = p.Split(':');
                //        string r = words[0];

                //        DeviationNames.Add(r);
                //    }
                //    foreach (string deviationname in DeviationNames)
                //    {
                //        string file = deviationname.Split(".")[0];
                //        string eventId = formData.EventId;
                //        try
                //        {
                //            Row newRow7 = new()
                //            {
                //                Cells = new List<Cell>()
                //            };

                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formData.EventName });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formData.EventType });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formData.EventStartDate });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formData.EventStartTime });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formData.EventEndTime });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "VenueName"), Value = formData.VenueName });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formData.City });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formData.State });


                //            if (file == "HCPHonorarium6LExceededFile")
                //            {
                //                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
                //                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = "Yes" });//formDataList.HandsOnTraining.EventOpen30days });
                //            }
                //            else if (file == "TrainerHonorarium12LExceededFile")
                //            {
                //                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });
                //                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = "Yes" });//formDataList.HandsOnTraining.EventWithin7days });

                //            }
                //            else if (file == "Travel_Accomodation3LExceededFile")
                //            {
                //                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Food and Beverages expense exceeds 1500" });
                //                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "PRE-F&B Expense Excluding Tax"), Value = "Yes" });
                //            }
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formData.SalesHeadEmail });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formData.FinanceEmail });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formData.InitiatorName });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formData.InitiatorEmail });

                //            IList<Row> addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                //            int j = 1;
                //            foreach (var p in formData.DeviationFiles)
                //            {
                //                string[] words = p.Split(':');
                //                string r = words[0];
                //                string q = words[1];
                //                if (deviationname == r)
                //                {
                //                    string name = r.Split(".")[0];
                //                    string filePath = SheetHelper.testingFile(q, eventId, name);
                //                    Row addedRow = addeddeviationrow[0];
                //                    Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                //                            sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                //                    j++;
                //                    if (System.IO.File.Exists(filePath))
                //                    {
                //                        SheetHelper.DeleteFile(filePath);
                //                    }
                //                }
                //            }
                //        }
                //        catch (Exception ex)
                //        {
                //            return BadRequest(ex.Message);
                //        }
                //    }
                //}

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }


            return Ok(new
            { Message = " Success!" });
        }


    }
}
