using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            VSTOContrib.Core.RibbonFactory.RibbonFactory.Current.SetApplication(Application, this);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            //Required for WPF support
            if (System.Windows.Application.Current == null)
                new System.Windows.Application { ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown };

            var assemblyContainingViewModels = typeof(ThisAddIn).Assembly; // This should be the assembly containing all your VSTOContrib viewmodels
            return new VSTOContrib.PowerPoint.RibbonFactory.PowerPointRibbonFactory(new VSTOContrib.Core.DefaultViewModelFactory(), () => CustomTaskPanes, Globals.Factory, assemblyContainingViewModels);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
