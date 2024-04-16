using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Text;

namespace IndiaEventsWebApi.Controllers.EventsController
{
    [Route("api/[controller]")]
    [ApiController]
    public class GetAndUpdateForAllEventsController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;
        private readonly string sheetId1;
        private readonly string sheetId2;
        private readonly string sheetId3;
        private readonly string sheetId4;
        private readonly string sheetId5;
        private readonly string sheetId6;
        private readonly string sheetId7;
        private readonly string sheetId8;
        private readonly string sheetId9;
        public GetAndUpdateForAllEventsController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
            sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
            sheetId3 = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
            sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            sheetId5 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
            sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;
            sheetId8 = configuration.GetSection("SmartsheetSettings:EventRequestBeneficiary").Value;
            sheetId9 = configuration.GetSection("SmartsheetSettings:EventRequestProductBrandsList").Value;
        }
        #region
        //[HttpGet("GetDataFromAllSheetsUsingEventId")]
        //public IActionResult GetDataFromAllSheetsUsingEventId(string eventId)
        //{
        //    Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
        //    //Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
        //    //Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
        //    //Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
        //    //Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
        //    //Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
        //    //Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);
        //    //Sheet sheet8 = SheetHelper.GetSheetById(smartsheet, sheetId8);
        //    //Sheet sheet9 = SheetHelper.GetSheetById(smartsheet, sheetId9);
        //    Dictionary<string, object> DraftData = new();
        //    Row existingRow = sheet1.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));

        //    if (existingRow != null)
        //    {
        //        List<string> columnNames = new List<string>();
        //        foreach (Column column in sheet1.Columns)
        //        {
        //            columnNames.Add(column.Title);
        //        }
        //        for (int i = 0; i < columnNames.Count; i++)
        //        {
        //            var x = existingRow.Cells[i].Value;
        //            if (columnNames[i] == "Brands")
        //            {
        //                if (x != null || x == "")
        //                {
        //                    List<object> brandsList = SheetHelper.ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
        //                    DraftData[columnNames[i]] = brandsList;
        //                }

        //            }
        //            else if (columnNames[i] == "Panelists")
        //            {
        //                if (x != null || x == "")
        //                {
        //                    List<object> Panelists = SheetHelper.ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
        //                    DraftData[columnNames[i]] = Panelists;
        //                }
        //            }
        //            else if (columnNames[i] == "SlideKits")
        //            {
        //                if (x != null || x == "")
        //                {
        //                    List<object> SlideKits = SheetHelper.ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
        //                    DraftData[columnNames[i]] = SlideKits;
        //                }
        //            }
        //            else if (columnNames[i] == "Invitees")
        //            {
        //                if (x != null || x == "")
        //                {
        //                    List<object> Invitees = SheetHelper.ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
        //                    DraftData[columnNames[i]] = Invitees;
        //                }
        //            }


        //            else if (columnNames[i] == "Expenses")
        //            {
        //                if (x != null || x == "")
        //                {
        //                    List<object> Expenses = SheetHelper.ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
        //                    DraftData[columnNames[i]] = Expenses;

        //                }
        //            }
        //            else
        //            {
        //                DraftData[columnNames[i]] = existingRow.Cells[i].Value;
        //            }


        //        }
        //        var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet1.Id.Value, existingRow.Id.Value, null);

        //        List<Dictionary<string, object>> attachmentsList = new List<Dictionary<string, object>>();

        //        foreach (var attachment in attachments.Data)
        //        {
        //            var AID = (long)attachment.Id;
        //            var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet1.Id.Value, AID);
        //            Dictionary<string, object> attachmentInfo = new Dictionary<string, object>
        //            {
        //                { "Name", file.Name },
        //                { "Url", file.Url }
        //            };
        //            attachmentsList.Add(attachmentInfo);
        //        }
        //        DraftData["Attachments"] = attachmentsList;
        //    }
        //    return Ok(DraftData);





        //}
        #endregion

        [HttpGet("GetDataFromAllSheetsUsingEventIdInPreEvent")]
        public IActionResult GetDataFromAllSheetsUsingEventId(string eventId)
        {

            Dictionary<string, object> resultData = new Dictionary<string, object>();

            List<Sheet> sheets = new List<Sheet>
            {
                SheetHelper.GetSheetById(smartsheet, sheetId1),
                SheetHelper.GetSheetById(smartsheet, sheetId2),
                SheetHelper.GetSheetById(smartsheet, sheetId3),
                SheetHelper.GetSheetById(smartsheet, sheetId4),
                SheetHelper.GetSheetById(smartsheet, sheetId5),
                SheetHelper.GetSheetById(smartsheet, sheetId6),
                //SheetHelper.GetSheetById(smartsheet, sheetId7),
                SheetHelper.GetSheetById(smartsheet, sheetId8),
                SheetHelper.GetSheetById(smartsheet, sheetId9)
            };
            foreach (var sheet in sheets)
            {
                List<Dictionary<string, object>> rowsData = new List<Dictionary<string, object>>();

                foreach (var row in sheet.Rows)
                {
                    if (row.Cells.Any(c => c.DisplayValue == eventId))
                    {
                        Dictionary<string, object> rowData = new Dictionary<string, object>();

                        List<string> columnNames = new List<string>();
                        foreach (Column column in sheet.Columns)
                        {
                            columnNames.Add(column.Title);
                        }

                        for (int i = 0; i < columnNames.Count; i++)
                        {
                            rowData[columnNames[i]] = row.Cells[i].Value;
                        }

                        var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet.Id.Value, row.Id.Value, null);
                        List<Dictionary<string, object>> attachmentsList = new List<Dictionary<string, object>>();
                        foreach (var attachment in attachments.Data)
                        {
                            var AID = (long)attachment.Id;
                            var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet.Id.Value, AID);
                            Dictionary<string, object> attachmentInfo = new Dictionary<string, object>
                    {
                        { "Name", file.Name },
                        { "Url", file.Url }
                    };
                            attachmentsList.Add(attachmentInfo);
                        }

                        rowData["Attachments"] = attachmentsList;
                        rowsData.Add(rowData);
                    }
                }

                if (rowsData.Count > 0)
                {
                    resultData[sheet.Name] = rowsData;
                }
            }




            string jsonData = JsonConvert.SerializeObject(resultData);

            return Ok(jsonData);





           // return Ok(resultData);




        }

        [HttpPut("UpdateClassIPreEvent")]
        public IActionResult UpdateClassIPreEvent(UpdateAllObjModels formDataList)
        {
            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
            Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

            StringBuilder addedBrandsData = new();
            StringBuilder addedInviteesData = new();
            StringBuilder addedMEnariniInviteesData = new();
            StringBuilder addedHcpData = new();
            StringBuilder addedSlideKitData = new();
            StringBuilder addedExpences = new();

            int addedSlideKitDataNo = 1;
            int addedHcpDataNo = 1;
            int addedInviteesDataNo = 1;
            int addedInviteesDataNoforMenarini = 1;
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;

            int TotalHonorariumAmount = 0;
            int TotalTravelAmount = 0;
            int TotalAccomodateAmount = 0;
            int TotalHCPLcAmount = 0;
            int TotalInviteesLcAmount = 0;
            int TotalExpenseAmount = 0;

            foreach (var formdata in formDataList.EventRequestExpenseSheet)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;
                var amount = SheetHelper.NumCheck(formdata.Amount);
                TotalExpenseAmount += amount;
            }

            string Expense = addedExpences.ToString();
            foreach (var formdata in formDataList.EventRequestHCPSlideKits)
            {
                string rowData = $"{addedSlideKitDataNo}. {formdata.HcpName} | {formdata.SlideKitType}";
                addedSlideKitData.AppendLine(rowData);
                addedSlideKitDataNo++;
            }
            string slideKit = addedSlideKitData.ToString();
            foreach (var formdata in formDataList.RequestBrandsList)
            {
                string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                addedBrandsData.AppendLine(rowData);
                addedBrandsDataNo++;
            }
            string brand = addedBrandsData.ToString();
            foreach (var formdata in formDataList.EventRequestInvitees)
            {
                if (formdata.InviteedFrom == "Menarini Employees")
                {
                    string row = $"{addedInviteesDataNoforMenarini}. {formdata.InviteeName}";
                    addedMEnariniInviteesData.AppendLine(row);
                    addedInviteesDataNoforMenarini++;
                }
                else
                {
                    string rowData = $"{addedInviteesDataNo}. {formdata.InviteeName}";
                    addedInviteesData.AppendLine(rowData);
                    addedInviteesDataNo++;
                }
                TotalInviteesLcAmount = TotalInviteesLcAmount + SheetHelper.NumCheck(formdata.LcAmount);
            }
            string Invitees = addedInviteesData.ToString();
            string MenariniInvitees = addedMEnariniInviteesData.ToString();
            foreach (var formdata in formDataList.EventRequestHcpRole)
            {
                int HM = SheetHelper.NumCheck(formdata.HonarariumAmount);
                int t = SheetHelper.NumCheck(formdata.Travel) + SheetHelper.NumCheck(formdata.Accomdation);
                string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |{formdata.HcpName} | Honr.Amt: {HM} |Trav.&Acc.Amt: {t} ";
                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;
                TotalTravelAmount += SheetHelper.NumCheck(formdata.Travel);
                TotalTravelAmount += SheetHelper.NumCheck(formdata.Travel);
                TotalAccomodateAmount += SheetHelper.NumCheck(formdata.Accomdation);
                TotalHCPLcAmount += SheetHelper.NumCheck(formdata.LocalConveyance);
            }
            string HCP = addedHcpData.ToString();
            int c = TotalHCPLcAmount + TotalInviteesLcAmount;
            int total = TotalHonorariumAmount + TotalTravelAmount + TotalAccomodateAmount + TotalHCPLcAmount + TotalInviteesLcAmount + TotalExpenseAmount;
            int s = TotalTravelAmount + TotalAccomodateAmount;


            return Ok();




        }

    }




}
