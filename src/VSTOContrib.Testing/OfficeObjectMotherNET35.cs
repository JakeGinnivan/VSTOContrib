using Microsoft.Office.Tools;

namespace VSTOContrib.Testing
{
    public class OfficeObjectMother
    {
        public static CustomTaskPaneCollection CreateCustomTaskPaneCollection()
        {
            return new CustomTaskPaneCollection();
        }
    }
}
