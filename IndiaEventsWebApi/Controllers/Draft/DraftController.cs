using Aspose.Pdf.Plugins;
using IndiaEvents.Models.Models.Draft;
using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Collections.Concurrent;
using System.Drawing.Printing;
using System.Globalization;
using System.Text;

namespace IndiaEventsWebApi.Controllers.Draft
{
    [Route("api/[controller]")]
    [ApiController]
    public class DraftController : ControllerBase
    {
        private readonly string accessToken;
        private readonly string sheetId;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;
        private readonly Sheet sheet;
        public DraftController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            sheetId = configuration.GetSection("SmartsheetSettings:DraftSheet").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
        }
        [HttpPost("PostDraftData")]
        public IActionResult PostDraftData(PostDraftData formDataList)
        {
            try
            {
                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventDate"), Value = formDataList.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "InitiatorName"), Value = formDataList.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Email"), Value = formDataList.InitiatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Role"), Value = formDataList.Role });
                var addedRows = smartsheet.SheetResources.RowResources.AddRows((long)sheet.Id, new Row[] { newRow });
                var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet, "Draft ID");
                var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                var val = eventIdCell.DisplayValue;
                if (formDataList.Isfile == "Yes")
                {
                    foreach (var p in formDataList.UploadFiles)
                    {
                        string[] words = p.Split(':');
                        var r = words[0];
                        var q = words[1];
                        var name = r.Split(".")[0];
                        var filePath = SheetHelper.testingFile(q, val, name);
                        var addedRow = addedRows[0];
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile((long)sheet.Id, addedRow.Id.Value, filePath, "application/msword");
                        if (System.IO.File.Exists(filePath))
                        {
                            SheetHelper.DeleteFile(filePath);
                        }
                    }
                }
                return Ok(new
                { Message = $"Row with id {val} created Successfully!" });
            }
            catch (Exception ex)
            {
                return BadRequest($"Could not find {ex.Message}");
            }
        }


        [HttpGet("GetDraftDatausingId")]
        public IActionResult GetDraftDatausingId(string id)
        {
            Dictionary<string, object> DraftData = new Dictionary<string, object>();
            Row existingRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == id));
            if (existingRow != null)
            {
                List<string> columnNames = new List<string>();
                foreach (Column column in sheet.Columns)
                {
                    columnNames.Add(column.Title);
                }
                for (int i = 0; i < columnNames.Count; i++)
                {
                    var x = existingRow.Cells[i].Value;
                    if (columnNames[i] == "Brands")
                    {
                        if (x != null || x == "")
                        {
                            List<object> brandsList = ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
                            DraftData[columnNames[i]] = brandsList;
                        }

                    }
                    else if (columnNames[i] == "Panelists")
                    {
                        if (x != null || x == "")
                        {
                            List<object> Panelists = ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
                            DraftData[columnNames[i]] = Panelists;
                        }
                    }
                    else if (columnNames[i] == "SlideKits")
                    {
                        if (x != null || x == "")
                        {
                            List<object> SlideKits = ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
                            DraftData[columnNames[i]] = SlideKits;
                        }
                    }
                    else if (columnNames[i] == "Invitees")
                    {
                        if (x != null || x == "")
                        {
                            List<object> Invitees = ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
                            DraftData[columnNames[i]] = Invitees;
                        }
                    }


                    else if (columnNames[i] == "Expenses")
                    {
                        if (x != null || x == "")
                        {
                            List<object> Expenses = ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
                            DraftData[columnNames[i]] = Expenses;

                        }
                    }
                    else
                    {
                        DraftData[columnNames[i]] = existingRow.Cells[i].Value;
                    }

                    
                }

                var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet.Id.Value, existingRow.Id.Value, null);

                List<Dictionary<string, object>> attachmentsList = new List<Dictionary<string, object>>();
                
                foreach (var attachment in attachments.Data)
                {
                    var AID = (long)attachment.Id;
                    var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet.Id.Value, AID);
                    Dictionary<string, object> attachmentInfo = new Dictionary<string, object>
                    {
                        { "Name", file.Name.Split("-")[1] },
                        { "Url", file.Url }
                    };
                    attachmentsList.Add(attachmentInfo);
                }
                DraftData["Attachments"] = attachmentsList;
            }
            return Ok(DraftData);
        }

        [HttpDelete("DeleteRowUsingEventId")]
        public IActionResult DeleteRowUsingEventId(string id)
        {
            try
            {
                Row existingRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == id));
                if (existingRow == null)
                {
                    return NotFound($"Row with id {id} not found.");
                }
                var Id = existingRow.Id.Value;
                smartsheet.SheetResources.RowResources.DeleteRows(sheet.Id.Value, new long[] { Id }, true);

                return Ok(new { Message = "Data Deleted successfully." });
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpPut("UpdateDraftData")]
        public IActionResult UpdateDraftData(DraftData formDataList)
        {
            StringBuilder addedBrandsData = new StringBuilder();
            StringBuilder addedInviteesData = new StringBuilder();
            StringBuilder addedHcpData = new StringBuilder();
            StringBuilder addedSlideKitData = new StringBuilder();
            StringBuilder addedExpences = new StringBuilder();

            int addedSlideKitDataNo = 1;
            int addedHcpDataNo = 1;
            int addedInviteesDataNo = 1;
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;

            foreach (var formdata in formDataList.EventRequestExpenseSheet)
            {
                string rowData = $"Expense: {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} |BTC/BTE: {formdata.BtcorBte} | BTC Amount: {formdata.BtcAmount} |BTE Amount: {formdata.BteAmount} | Budget Amount: {formdata.BudgetAmount}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;

            }
            string Expense = addedExpences.ToString();

            foreach (var formdata in formDataList.EventRequestHCPSlideKits)
            {

                string rowData = $" MIS: {formdata.MIS} | SlideKitType: {formdata.SlideKitType} ";
                addedSlideKitData.AppendLine(rowData);
                addedSlideKitDataNo++;
            }
            string slideKit = addedSlideKitData.ToString();

            foreach (var formdata in formDataList.RequestBrandsList)
            {

                string rowData = $"BrandName: {formdata.BrandName} |ProjectId: {formdata.ProjectId} |PercentAllocation: {formdata.PercentAllocation}";
                addedBrandsData.AppendLine(rowData);
                addedBrandsDataNo++;
            }
            string brand = addedBrandsData.ToString();

            foreach (var formdata in formDataList.EventRequestInvitees)
            {

                string rowData = $"InviteeName: {formdata.InviteeName} | MISCode: {formdata.InviteeName} | LocalConveyance: {formdata.LocalConveyance} | BtcorBte: {formdata.BtcorBte} | LcAmount: {formdata.LcAmount} | InviteedFrom: {formdata.InviteedFrom} | Speciality: {formdata.Speciality} | HCPType: {formdata.HCPType} ";
                //string rowData = $"{addedInviteesDataNo}. {formdata.InviteeName}";
                addedInviteesData.AppendLine(rowData);
                addedInviteesDataNo++;

            }
            string Invitees = addedInviteesData.ToString();


            foreach (var formdata in formDataList.EventRequestHcpRole)
            {


                string rowData = $"HcpName: {formdata.HcpName} |HcpRole: {formdata.HcpRole} | MisCode: {formdata.MisCode} |SpeakerCode: {formdata.SpeakerCode} |TrainerCode: {formdata.TrainerCode} " +
                    $"|Speciality: {formdata.Speciality} |Tier: {formdata.Tier} | GOorNGO: {formdata.GOorNGO} |HonorariumRequired: {formdata.HonorariumRequired} |HonarariumAmount: {formdata.HonarariumAmount}" +
                    $" | Travel: {formdata.Travel} |Accomdation: {formdata.Accomdation} | LocalConveyance: {formdata.LocalConveyance} |FinalAmount: {formdata.FinalAmount} |Rationale: {formdata.Rationale}" +
                    $" |PresentationDuration: {formdata.PresentationDuration} |PanelSessionPreperationDuration: {formdata.PanelSessionPreperationDuration} " +
                    $"| PanelDisscussionDuration: {formdata.PanelDisscussionDuration} |QASessionDuration: {formdata.QASessionDuration} |BriefingSession: {formdata.BriefingSession} " +
                    $"| TotalSessionHours: {formdata.TotalSessionHours} |IsInclidingGst: {formdata.IsInclidingGst} |AgreementAmount: {formdata.AgreementAmount} ";
                // string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |{formdata.HcpName} | Honr.Amt: {} |Trav.&Acc.Amt: {} ";

                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;

            }
            string HCP = addedHcpData.ToString();
            var targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formDataList.Draft.DraftId));
            if (targetRow != null)
            {

                Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                //updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Accounts Given Details"), Value = FinanceAccounts });
                //updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "HON-Finance Accounts Approval"), Value = updatedFormData.Status });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Event Topic"), Value = formDataList.Draft.EventTopic });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventType"), Value = formDataList.Draft.EventType });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventDate"), Value = formDataList.Draft.EventDate });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "StartTime"), Value = formDataList.Draft.StartTime });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EndTime"), Value = formDataList.Draft.EndTime });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Event End Date"), Value = formDataList.Draft.EventEndDate });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "VenueName"), Value = formDataList.Draft.VenueName });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "City"), Value = formDataList.Draft.City });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "State"), Value = formDataList.Draft.State });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "HCP Role"), Value = formDataList.Draft.HCPRole });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "IsAdvanceRequired"), Value = formDataList.Draft.IsAdvanceRequired });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Advance Amount"), Value = formDataList.Draft.AdvanceAmount });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Class III Event Code"), Value = formDataList.Draft.ClassIIIEventCode });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Meeting Type"), Value = formDataList.Draft.MeetingType });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sponsorship Society Name"), Value = formDataList.Draft.SponcershipSocietyname });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Venue Country"), Value = formDataList.Draft.VenueCountry });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Medical Utility Type"), Value = formDataList.Draft.MedicalUtilityType });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Medical Utility Description"), Value = formDataList.Draft.MedicalUtilityDescription });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Valid To"), Value = formDataList.Draft.ValidFrom });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Valid From"), Value = formDataList.Draft.ValidTo });
                if (formDataList.Draft.IsPanelists == "Yes")
                {
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Panelists"), Value = HCP });
                }
                if (formDataList.Draft.IsInvitees == "Yes")
                {
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Invitees"), Value = Invitees });
                }
                if (formDataList.Draft.IsBrands == "Yes")
                {
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Brands"), Value = brand });
                }
                if (formDataList.Draft.IsSlideKits == "Yes")
                {
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "SlideKits"), Value = slideKit });
                }
                if (formDataList.Draft.IsExpense == "Yes")
                {
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Expenses"), Value = Expense });
                }

                smartsheet.SheetResources.RowResources.UpdateRows((long)sheet.Id, new Row[] { updateRow });

            }

            return Ok(new { Message = "Draft Updated" });

        }


        private List<object> ConvertToJsonObject(string data)
        {
            string[] lines = data.Split('\n');
            List<object> resultList = new List<object>();

            foreach (var line in lines)
            {
                string[] values = line.Split('|');

                if (values.Length > 0)
                {
                    Dictionary<string, string> item = new Dictionary<string, string>();

                    foreach (var value in values)
                    {
                        string[] keyValue = value.Trim().Split(':');
                        if (keyValue.Length == 2)
                        {
                            item.Add(keyValue[0].Trim(), keyValue[1].Trim());
                        }
                    }

                    resultList.Add(item);
                }
            }

            return resultList;
        }
    }
}

