using Microsoft.Office.Tools;
using NSubstitute;

namespace Office.Contrib.Testing
{
    public static class OfficeObjectMother
    {
        public static CustomTaskPaneCollection CreateCustomTaskPaneCollection()
        {
#if NET35
            return new CustomTaskPaneCollection();
#else
            return Substitute.For<CustomTaskPaneCollection>();
#endif
        }
    }
}
