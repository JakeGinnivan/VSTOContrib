using Microsoft.Office.Interop.Excel;

namespace Office.Excel.Contrib
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