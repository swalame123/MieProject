using IndiaEventsWebApi.Models.EventTypeSheets;
using Microsoft.AspNetCore.Http.HttpResults;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Text;

namespace IndiaEventsWebApi.Helper
{
    public class ApiCalls
    {
        #region
        public static IList<Row> AddWebinarData(SmartsheetClient smartsheet, Sheet sheet1, WebinarPayload formDataList)
        {
            try
            {



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
                double TotalHonorariumAmount = 0;
                double TotalTravelAmount = 0;
                double TotalAccomodateAmount = 0;
                double TotalHCPLcAmount = 0;
                double TotalInviteesLcAmount = 0;
                double TotalExpenseAmount = 0;
                foreach (var formdata in formDataList.EventRequestExpenseSheet)
                {
                    string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                    addedExpences.AppendLine(rowData);
                    addedExpencesNo++;
                    var amount = SheetHelper.NumCheck(formdata.Amount);
                    TotalExpenseAmount = TotalExpenseAmount + amount;
                }
                string Expense = addedExpences.ToString();

                StringBuilder addedExpencesBTE = new();
                int addedExpencesNoBTE = 1;
                foreach (var formdata in formDataList.EventRequestExpenseSheet)
                {
                    if (formdata.BtcorBte.ToLower() == "bte")
                    {
                        string rowData = $"{addedExpencesNoBTE}. {formdata.Expense} | Amount: {formdata.Amount}";
                        addedExpencesBTE.AppendLine(rowData);
                        addedExpencesNoBTE++;
                    }
                }
                string BTEExpense = addedExpencesBTE.ToString();

                foreach (var formdata in formDataList.EventRequestHCPSlideKits)
                {
                    string rowData = $"{addedSlideKitDataNo}. {formdata.HcpName} | {formdata.SlideKitType} | Id :{formdata.SlideKitDocument}";
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
                    double HM = SheetHelper.NumCheck(formdata.HonarariumAmount);
                    double t = SheetHelper.NumCheck(formdata.Travel) + SheetHelper.NumCheck(formdata.Accomdation);

                    double roundedValue = Math.Round(t, 2);

                    string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |{formdata.HcpName} | Honr.Amt: {HM} |Trav.&Acc.Amt: {roundedValue} |Rationale : {formdata.Rationale}";
                    addedHcpData.AppendLine(rowData);
                    addedHcpDataNo++;
                    TotalHonorariumAmount = TotalHonorariumAmount + SheetHelper.NumCheck(formdata.HonarariumAmount);
                    TotalTravelAmount = TotalTravelAmount + SheetHelper.NumCheck(formdata.Travel);
                    TotalAccomodateAmount = TotalAccomodateAmount + SheetHelper.NumCheck(formdata.Accomdation);
                    TotalHCPLcAmount = TotalHCPLcAmount + SheetHelper.NumCheck(formdata.LocalConveyance);
                }
                string HCP = addedHcpData.ToString();


                double cc = TotalHCPLcAmount + TotalInviteesLcAmount;

                double totalAmount = TotalHonorariumAmount + TotalTravelAmount + TotalAccomodateAmount + TotalHCPLcAmount + TotalInviteesLcAmount + TotalExpenseAmount;

                double ss = TotalTravelAmount + TotalAccomodateAmount;

                double c = Math.Round(cc, 2);
                double total = Math.Round(totalAmount, 2);
                double s = Math.Round(ss, 2);

                Row newRow = new()
                {
                    Cells = new List<Cell>()
                };
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.Webinar.EventTopic });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Type"), Value = formDataList.Webinar.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Date"), Value = formDataList.Webinar.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Start Time"), Value = formDataList.Webinar.StartTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "End Time"), Value = formDataList.Webinar.EndTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Meeting Type"), Value = formDataList.Webinar.Meeting_Type });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Invitees"), Value = Invitees });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "MIPL Invitees"), Value = MenariniInvitees });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "SlideKits"), Value = slideKit });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.Webinar.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventOpen30days"), Value = formDataList.Webinar.EventOpen30days });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventWithin7days"), Value = formDataList.Webinar.EventWithin7days });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Name"), Value = formDataList.Webinar.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = SheetHelper.NumCheck(formDataList.Webinar.AdvanceAmount) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = SheetHelper.NumCheck(formDataList.Webinar.TotalExpenseBTC) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = SheetHelper.NumCheck(formDataList.Webinar.TotalExpenseBTE) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Honorarium Amount"), Value = Math.Round(TotalHonorariumAmount, 2) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel Amount"), Value = Math.Round(TotalTravelAmount, 2) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel & Accommodation Amount"), Value = s });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Accommodation Amount"), Value = Math.Round(TotalAccomodateAmount, 2) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Budget Amount"), Value = total });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = c });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = Math.Round(TotalExpenseAmount, 2) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.Webinar.Initiator_Email });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.Webinar.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.Webinar.Sales_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.Webinar.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Coordinator"), Value = formDataList.Webinar.MarketingCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.Webinar.Marketing_Head });
                //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Treasury"), Value = formData.RequestHonorariumList.MarketingHeadEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.Webinar.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.Webinar.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.Webinar.Finance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.Webinar.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.Webinar.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.Webinar.MedicalAffairsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "BTE Expense Details"), Value = formDataList.Webinar.BTEExpenseDetails });


                IList<Row> addedRow = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });
                return addedRow;
            }
            catch (Exception ex)
            {

                return (IList<Row>)AddWebinarData(smartsheet, sheet1, formDataList);
            }
        }
        #endregion
        public static async Task<Attachment> AddAttachmentsToSheet(SmartsheetClient smartsheet, Sheet sheet1, Row addedRow, string filePath, int count = 0)
        {

            try
            {

                Attachment attachment = await Task.Run(() => smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                          sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword"));
                return attachment;

            }
            catch (Exception ex)
            {
                if (count >= 8)
                {
                    throw ex;
                }

                return await AddAttachmentsToSheet(smartsheet, sheet1, addedRow, filePath, count + 1);
            }

        }

        public static async Task<Attachment> UpdateAttachmentsToSheet(SmartsheetClient smartsheet, Sheet sheet1, long Id, string filePath, int count = 0)
        {

            try
            {

                Attachment attachment = smartsheet.SheetResources.AttachmentResources.VersioningResources.AttachNewVersion(
                                    sheet1.Id.Value, Id, filePath, "application/msword");
                return attachment;

            }
            catch (Exception ex)
            {
                if (count >= 8)
                {
                    throw ex;
                }

                return await UpdateAttachmentsToSheet(smartsheet, sheet1, Id, filePath, count + 1);
            }

        }

        public static async Task<Attachment> AddAttachmentsToSheetSync(SmartsheetClient smartsheet, Sheet sheet1, Row addedRow, string filename, int count = 0)
        {

            try
            {
                var folderName = Path.Combine("Resources", "Images");
                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                string filePath = Path.Combine(pathToSave, filename);
                Attachment attachment = await Task.Run(() => smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                          sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword"));
                return attachment;

            }
            catch (Exception ex)
            {
                if (count >= 8)
                {
                    throw ex;
                }

                return await AddAttachmentsToSheetSync(smartsheet, sheet1, addedRow, filename, count + 1);
            }

        }

        public static async Task<PaginatedResult<Attachment>> GetAttachmantsFromSheet(SmartsheetClient smartsheet, Sheet sheet1, Row row, int count = 0)
        {

            try
            {
                PaginatedResult<Attachment> attachments = await Task.Run(() => smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet1.Id.Value, row.Id.Value, null));
                return attachments;
            }
            catch (Exception ex)
            {


                if (count >= 8)
                {
                    throw ex;
                }

                return await GetAttachmantsFromSheet(smartsheet, sheet1, row, count + 1);
            }

        }

        public static async Task<Attachment> GetAttachment(SmartsheetClient smartsheet, Sheet sheet1, long AID, int count = 0)
        {

            try
            {
                Attachment attachments = await Task.Run(() => smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet1.Id.Value, AID));
                return attachments;
            }
            catch (Exception ex)
            {


                if (count >= 8)
                {
                    throw ex;
                }

                return await GetAttachment(smartsheet, sheet1, AID, count + 1);
            }

        }


        //public static async Task<PaginatedResult<Attachment>> DeleteAttachment(SmartsheetClient smartsheet, Sheet sheet4, long Id, int count = 0)
        //{
        //    try
        //    {
        //        PaginatedResult<Attachment> attachments = await Task.Run(() => smartsheet.SheetResources.AttachmentResources.DeleteAttachment(sheet4.Id.Value, Id));
        //        return attachments;
        //    }
        //    catch (Exception ex)
        //    {
        //        // Handle the exception as needed
        //        if (count >= 8)
        //        {
        //            throw ex;
        //        }

        //        return await DeleteAttachment(smartsheet, sheet4, Id, count + 1);
        //    }
        //}
        //public static PaginatedResult<Attachment> DeleteAttachment(SmartsheetClient smartsheet, Sheet sheet4, long Id, int count = 0)
        //{

        //    try
        //    {
        //        var attachments = smartsheet.SheetResources.AttachmentResources.DeleteAttachment(sheet4.Id.Value, Id));

        //        return attachments;
        //    }
        //    catch (Exception ex)
        //    {


        //        if (count >= 8)
        //        {
        //            throw ex;
        //        }

        //        return await DeleteAttachment(smartsheet, sheet4, Id, count + 1);
        //    }
        //}

        public static IList<Row> WebDetails(SmartsheetClient smartsheet, Sheet sheet4, Row newRow1, int count = 0)
        {
            try
            {
                IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });
                return row;
            }
            catch (Exception ex)
            {
                if (count >= 5)
                {
                    throw ex;
                }
                return WebDetails(smartsheet, sheet4, newRow1, count + 1);

            }

        }
        public static IList<Row> PanelDetails(SmartsheetClient smartsheet, Sheet sheet4, Row newRow1, int count = 0)
        {
            try
            {
                IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });
                return row;
            }
            catch (Exception ex)
            {
                if (count >= 5)
                {
                    throw ex;
                }
                return PanelDetails(smartsheet, sheet4, newRow1, count + 1);

            }

        }
        public static IList<Row> SlideKitDetails(SmartsheetClient smartsheet, Sheet sheet4, Row newRow1, int count = 0)
        {
            try
            {
                IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });
                return row;
            }
            catch (Exception ex)
            {
                if (count >= 5)
                {
                    throw ex;
                }
                return SlideKitDetails(smartsheet, sheet4, newRow1, count + 1);

            }

        }
        public static IList<Row> BrandsDetails(SmartsheetClient smartsheet, Sheet sheet2, List<Row> newRows2, int count = 0)
        {
            try
            {
                IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, newRows2.ToArray());
                return row;
            }
            catch (Exception ex)
            {
                if (count >= 5)
                {
                    throw ex;
                }
                return BrandsDetails(smartsheet, sheet2, newRows2, count + 1);

            }

        }
        public static IList<Row> InviteesDetails(SmartsheetClient smartsheet, Sheet sheet, List<Row> newRows2, int count = 0)
        {
            try
            {
                IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet.Id.Value, newRows2.ToArray());
                return row;
            }
            catch (Exception ex)
            {
                if (count >= 5)
                {
                    throw ex;
                }
                return InviteesDetails(smartsheet, sheet, newRows2, count + 1);

            }

        }
        public static IList<Row> ExpenseDetails(SmartsheetClient smartsheet, Sheet sheet, List<Row> newRows2, int count = 0)
        {
            try
            {
                IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet.Id.Value, newRows2.ToArray());
                return row;
            }
            catch (Exception ex)
            {
                if (count >= 5)
                {
                    throw ex;
                }
                return ExpenseDetails(smartsheet, sheet, newRows2, count + 1);

            }

        }
        public static IList<Row> HonorariumDetails(SmartsheetClient smartsheet, Sheet sheet, Row newRows2, int count = 0)
        {
            try
            {
                IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet.Id.Value, new Row[] { newRows2 });
                return row;
            }
            catch (Exception ex)
            {
                if (count >= 5)
                {
                    throw ex;
                }
                return HonorariumDetails(smartsheet, sheet, newRows2, count + 1);

            }

        }
        public static IList<Row> EventSettlementDetails(SmartsheetClient smartsheet, Sheet sheet, Row newRows2, int count = 0)
        {
            try
            {
                IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet.Id.Value, new Row[] { newRows2 });
                return row;
            }
            catch (Exception ex)
            {
                if (count >= 5)
                {
                    throw ex;
                }
                return HonorariumDetails(smartsheet, sheet, newRows2, count + 1);

            }

        }
        public static IList<Row> UpdateRole(SmartsheetClient smartsheet, Sheet sheet4, Row updateRows, int count = 0)
        {
            try
            {
                IList<Row> row = smartsheet.SheetResources.RowResources.UpdateRows(sheet4.Id.Value, new Row[] { updateRows });
                return row;
            }
            catch (Exception ex)
            {
                if (count >= 5)
                {
                    throw ex;
                }
                return UpdateRole(smartsheet, sheet4, updateRows, count + 1);

            }

        }
        public static IList<Row> BulkUpdateRows(SmartsheetClient smartsheet, Sheet sheet, List<Row> liRowsToUpdate, int count = 0)
        {
            try
            {
                IList<Row> rows = smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, liRowsToUpdate);
                return rows;
            }
            catch (Exception ex)
            {
                if (count >= 5)
                {
                    throw ex;
                }
                return BulkUpdateRows(smartsheet, sheet, liRowsToUpdate, count + 1);

            }

        }


        public static IList<Row> DeviationData(SmartsheetClient smartsheet, Sheet sheet4, Row newRow1, int count = 0)
        {
            try
            {
                IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });
                return row;
            }
            catch (Exception ex)
            {
                if (count >= 5)
                {
                    throw ex;
                }
                return DeviationData(smartsheet, sheet4, newRow1, count + 1);

            }

        }
        public static List<Dictionary<string, object>> HcpData(string[] sheetIds, SmartsheetClient smartsheet, int count = 1)
        {
            List<Dictionary<string, object>> sheetData = new List<Dictionary<string, object>>();
            try
            {

                foreach (string sheetId in sheetIds)
                {
                    Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                    List<string> columnNames = sheet.Columns.Select(column => column.Title).ToList();
                    foreach (Row row in sheet.Rows)
                    {
                        Dictionary<string, object> rowData = new Dictionary<string, object>();
                        for (int i = 0; i < row.Cells.Count && i < columnNames.Count; i++)
                        {
                            rowData[columnNames[i]] = row.Cells[i].Value;
                        }
                        sheetData.Add(rowData);
                    }
                }

                return sheetData;
            }
            catch (Exception ex)
            {
                if (count >= 5)
                {
                    throw ex;
                }
                return HcpData(sheetIds, smartsheet, count + 1);
            }
        }
        public static async Task<Attachment> UpdateAttachments(SmartsheetClient smartsheet, long sheetId, long Id, string filePath, int count = 0)
        {

            try
            {

                Attachment attachment = smartsheet.SheetResources.AttachmentResources.VersioningResources.AttachNewVersion(
                                    sheetId, Id, filePath, "application/msword");
                return attachment;

            }
            catch (Exception ex)
            {
                if (count >= 8)
                {
                    throw ex;
                }

                return await UpdateAttachments(smartsheet, sheetId, Id, filePath, count + 1);
            }

        }

    }
}
