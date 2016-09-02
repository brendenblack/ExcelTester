using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTesting
{
    public class XlsmRateReportEntry
    {
        public string Mandate { get; set; }

        public string Workflow { get; set; }

        public int EmployeeNumber { get; set; }

        public string EmployeeName { get; set; }

        public DateTime Date { get; set; }

        public string ClientBranch { get; set; }

        public string ClientArea { get; set; }

        public string ProjectName { get; set; }

        public string TaskType { get; set; }

        public string FundingType { get; set; }

        public bool IsBillable { get; set; }

        public string Category { get; set; }

        public double Hours { get; set; }

        public string ServiceGroup { get; set; }

        public string RequestType { get; set; }

        public string RequestReference { get; set; }

        /// <summary>
        /// The value of the spreadsheet column, that more often than not will not reflect the actual team of the resource
        /// </summary>
        public string TeamColumn { get; set; }

        public string TicketSummary { get; set; }

        public string TaskDescription { get; set; }

        public string TimesheetComment { get; set; }



    }
}
