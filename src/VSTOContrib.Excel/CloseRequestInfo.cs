using Microsoft.Office.Interop.Excel;

namespace VSTOContrib.Excel
{
    internal class CloseRequestInfo
    {
        public CloseRequestInfo(Workbook workbook, int count)
        {
            Workbook = workbook;
            WorkbookCount = count;
        }

        public Workbook Workbook { get; set; }

        public int WorkbookCount { get; set; }
    }
}