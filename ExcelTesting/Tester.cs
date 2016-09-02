using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;

namespace ExcelTesting
{
    public class Tester
    {
        private DateTime startDate;
        private DateTime endDate;

        #region Column mapping
        // Hard-coded values are used to improve reading speed; the report is relatively stable
        private const int colMandate = 1;
        private const int colWorkflow = 2;
        private const int colEmployeeNumber = 3;
        private const int colEmployeeName = 4;
        private const int colDate = 5;
        private const int colClientBranch = 6;
        private const int colClientArea = 7;
        private const int colProjectName = 8;
        private const int colTaskType = 9;
        private const int colFundingType = 10;
        private const int colBillable = 11;
        private const int colCategory = 12;
        private const int colHours = 13;
        private const int colServiceGroup = 14;
        private const int colRequestType = 15;
        private const int colRequestRef = 16;
        private const int colTeam = 17;
        private const int colTicketSummary = 18;
        private const int colTaskDescription = 19;
        private const int colTimesheetComment = 20;
        private const int colRate = 21;
        private const int colAmountBilled = 22;
        private const int colResourceRole = 23;
        private const int colResourceLevel = 24;
        private const int colErrorMessage = 25;
        private const int colEstimateDays = 26;
        private const int colActualDays = 27;
        private const int colAnnualOperatingPlan = 28;
        private const int colStatementOfWork = 29;
        private const int colTicketBudget = 30;
        private const int colTotalInvoiced = 31;
        private const int colTotalRemaining = 32;
        #endregion


        public Tester(DateTime startDate, DateTime endDate)
        {
            this.startDate = startDate;
            this.endDate = endDate;
        }

        public TestResult CheckForDateValidityFirst(ExcelWorksheet sheet)
        {
            var start = DateTime.Now;
            var entries = new List<XlsmRateReportEntry>();

            for (int i = sheet.Dimension.Start.Row; i <= sheet.Dimension.End.Row; i++)
            {
                try
                {
                    var entry = new XlsmRateReportEntry();

                    entry.Date = DateTime.Parse(sheet.Cells[i, colDate].Value.ToString());

                    if (entry.Date <= startDate || entry.Date >= endDate)
                    {
                        continue;
                    }
                    entry.Mandate = ReadCellAsString(sheet.Cells[i, colMandate]);

                    entry.Workflow = ReadCellAsString(sheet.Cells[i, colWorkflow]);

                    entry.EmployeeNumber = (int)sheet.Cells[i, colEmployeeNumber].Value;

                    entry.EmployeeName = ReadCellAsString(sheet.Cells[i, colEmployeeName]);

                    entry.ClientBranch = sheet.Cells[i, colClientBranch].Value?.ToString();

                    entry.ClientArea = ReadCellAsString(sheet.Cells[i, colClientArea]);

                    entry.ProjectName = ReadCellAsString(sheet.Cells[i, colProjectName]);

                    entry.TaskType = ReadCellAsString(sheet.Cells[i, colTaskType]);

                    entry.FundingType = ReadCellAsString(sheet.Cells[i, colFundingType]);

                    entry.IsBillable = sheet.Cells[i, colBillable].Value?.ToString().Equals("Billable") ?? false;

                    entry.Category = ReadCellAsString(sheet.Cells[i, colCategory]);

                    entry.Hours = (double)sheet.Cells[i, colHours].Value;

                    entry.ServiceGroup = ReadCellAsString(sheet.Cells[i, colServiceGroup]);

                    entry.RequestType = ReadCellAsString(sheet.Cells[i, colRequestType]);

                    entry.RequestReference = ReadCellAsString(sheet.Cells[i, colRequestRef]);

                    entry.TeamColumn = ReadCellAsString(sheet.Cells[i, colTeam]);

                    entry.TicketSummary = ReadCellAsString(sheet.Cells[i, colTicketSummary]);

                    entry.TicketSummary = ReadCellAsString(sheet.Cells[i, colTicketSummary]);

                    entry.TaskDescription = ReadCellAsString(sheet.Cells[i, colTaskDescription]);

                    entry.TimesheetComment = ReadCellAsString(sheet.Cells[i, colTimesheetComment]);

                    entries.Add(entry);
                }
                catch (Exception)
                {
                    // If a row causes problems, record it and move on

                    continue;
                }

            }

            var lapsed = (DateTime.Now - start).Seconds;

            return new TestResult { Seconds = lapsed, ResultsFound = entries.Count, Description = "Loop through entire set, parse date column to DateTime and skip if it falls out of range" };
        }



        private string ReadCellAsString(ExcelRange cell)
        {
            return cell.Value?.ToString() ?? string.Empty;
        }
    }
}
