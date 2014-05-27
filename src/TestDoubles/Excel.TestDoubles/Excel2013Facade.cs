using System;
using System.Linq;

namespace Excel.TestDoubles
{
    public class Excel2013Facade
    {
        public Excel2013Facade()
        {
            Application = new ApplicationTestDouble(new Version(14, 0));
        }

        public ApplicationTestDouble Application { get; private set; }

        public Tuple<WorkbookTestDouble, WindowTestDouble> NewWorksheet()
        {
            if (Application.Workbooks.Count == 0)
            {
                var windowTestDouble = Application.Windows.OfType<WindowTestDouble>().Single();
                var workbookTestDouble = new WorkbookTestDouble(Application, windowTestDouble);
                ((WorkbooksTestDouble)Application.Workbooks).Add(workbookTestDouble);
                Application.RaiseNewWorkbook(workbookTestDouble);
                Application.RaiseWorkbookOpen(workbookTestDouble);

                return Tuple.Create(workbookTestDouble, windowTestDouble);
            }

            throw new NotImplementedException();
        }
    }
}