using System;
using VSTOContrib.PowerPoint.RibbonFactory;
using Office = Microsoft.Office.Core;

namespace AddTextAddin
{
    public partial class ThisAddIn
    {
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            var assemblyContainingViewModels = typeof(PresentationViewModel).Assembly; // This should be the assembly containing all your VSTOContrib viewmodels
            return new PowerPointRibbonFactory(this, assemblyContainingViewModels);
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}
