using Microsoft.Office.Tools;

namespace Office.Contrib.Testing
{
    public class OfficeObjectMother
    {
        public static CustomTaskPaneCollection CreateCustomTaskPaneCollection()
        {
            return new CustomTaskPaneCollection();
        }
    }
}
