using IndiaEvents.Models.Models.RequestSheets;
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
    public class AddorDeletePanelistController : ControllerBase
    {


        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;

        private readonly string sheetId4;
        private readonly string sheetId7;

        public AddorDeletePanelistController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

            sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;

        }

        [HttpPost("AddPanelist"), DisableRequestSizeLimit]
        public IActionResult AddPanelist(EventRequestTrainerData formData)
        {
            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

            try
            {
                Row newRow1 = new()
                {
                    Cells = new List<Cell>()
                };

                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "MISCode"), Value = SheetHelper.MisCodeCheck(formData.MISCode )});
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HcpRole"), Value = formData.HCPRole });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HCPName"), Value = formData.TrainerName });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "TrainerCode"), Value = formData.TrainerCode });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Qualification"), Value = formData.TrainerQualification });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Country"), Value = formData.TrainerCountry });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Speciality"), Value = formData.Speciality });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Tier"), Value = formData.TrainerCategory });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HCP Type"), Value = formData.TrainerType });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Rationale"), Value = formData.Rationale });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "FCPA Date"), Value = formData.FCPAIssueDate });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HonorariumRequired"), Value = formData.IsHonorariumApplicable });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "PresentationDuration"), Value = formData.Presentation_Speaking_WorkshopDuration });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "PanelSessionPreparationDuration"), Value = formData.DevelopmentofPresentationPanelSessionPreparation });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "PanelDiscussionDuration"), Value = formData.PaneldiscussionSessionduration });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "QASessionDuration"), Value = formData.QASession });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "BriefingSession"), Value = formData.Speaker_TrainerBriefing });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "TotalSessionHours"), Value = formData.TotalNoOfHours });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "LC BTC/BTE"), Value = formData.IsLCBTC_BTE });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Accomodation BTC/BTE"), Value = formData.IsAccomodationBTC_BTE });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Honorarium Amount Excluding Tax"), Value = formData.HonorariumAmountexcludingTax });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HonorariumAmount"), Value = formData.HonorariumAmountincludingTax });
                //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(she1t1, "AgreementAmount"), Value = formData.AgreementAmount });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "YTD Spend Including Current Event"), Value = formData.YTDspendIncludingCurrentEvent });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Global FMV"), Value = formData.IsGlobalFMVCheck });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "ExpenseType"), Value = formData.ExpenseType });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Mode of Travel"), Value = formData.TravelSelection });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Travel BTC/BTE"), Value = formData.IsExpenseBTC_BTE });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Travel Excluding Tax"), Value = formData.TravelAmountExcludingTax });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Travel"), Value = formData.TravelAmountIncludingTax });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Accomodation Excluding Tax"), Value = formData.AccomodationAmountExcludingTax });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Accomodation"), Value = formData.AccomodationAmountIncludingTax });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Local Conveyance Excluding Tax"), Value = formData.LocalConveyanceAmountexcludingTax });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "LocalConveyance"), Value = formData.LocalConveyanceAmountincludingTax });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Travel/Accomodation Spend Including Current Event"), Value = formData.TravelandAccomodationspendincludingcurrentevent });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Currency"), Value = formData.BenificiaryDetailsData.Currency });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Other Currency"), Value = formData.BenificiaryDetailsData.EnterCurrencyType });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Beneficiary Name"), Value = formData.BenificiaryDetailsData.BenificiaryName });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Bank Account Number"), Value = formData.BenificiaryDetailsData.BankAccountNumber });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Bank Name"), Value = formData.BenificiaryDetailsData.BankName });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "PAN card name"), Value = formData.BenificiaryDetailsData.NameasPerPAN });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Pan Number"), Value = formData.BenificiaryDetailsData.PANCardNumber });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IFSC Code"), Value = formData.BenificiaryDetailsData.IFSCCode });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Email Id"), Value = formData.BenificiaryDetailsData.EmailID });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IBN Number"), Value = formData.BenificiaryDetailsData.IbnNumber });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Swift Code"), Value = formData.BenificiaryDetailsData.SwiftCode });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Tax Residence Certificate"), Value = formData.BenificiaryDetailsData.TaxResidenceCertificateDate });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "AgreementAmount"), Value = formData.AgreementAmount });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formData.EventName });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Type"), Value = formData.EventType });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Venue name"), Value = formData.VenueName });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Date Start"), Value = formData.EventStartDate });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event End Date"), Value = formData.EventEndDate });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "TotalSpend"), Value = formData.FinalAmount });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId"), Value = formData.EventId });

                smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow1 });



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




        [HttpDelete("DeleteData/{RowInvId}")]
        public IActionResult DeleteData(string RowInvId)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
                long.TryParse(sheetId, out long parsedSheetId);

                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                //Row existingRow = GetRowById(smartsheet, parsedSheetId, RowInvId);
                Row existingRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == RowInvId));


                if (existingRow == null)
                {
                    return NotFound($"Row with id {RowInvId} not found.");
                }
                var Id = (long)existingRow.Id;
                smartsheet.SheetResources.RowResources.DeleteRows(parsedSheetId, new long[] { Id }, true);

                return Ok(new { Message = "Data Deleted successfully." });
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }












    }
}
