using Microsoft.Office.Tools;
using NSubstitute;

namespace Office.Contrib.Testing
{
    public static class OfficeObjectMother
    {
        public static CustomTaskPaneCollection CreateCustomTaskPaneCollection()
        {
            return Substitute.For<CustomTaskPaneCollection>();
        }
    }
}
