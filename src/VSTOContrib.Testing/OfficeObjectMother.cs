using Microsoft.Office.Tools;
using NSubstitute;

namespace VSTOContrib.Testing
{
    public static class OfficeObjectMother
    {
        public static CustomTaskPaneCollection CreateCustomTaskPaneCollection()
        {
            return Substitute.For<CustomTaskPaneCollection>();
        }
    }
}
