using System;
using System.Collections.Generic;
using System.Reflection;
using FacebookToOutlook;
using Microsoft.Office.Core;
using Office.Utility;
using Outlook.Utility;
using Outlook.Utility.RibbonFactory;

namespace FacebookToOutlookAddin
{
    public partial class ThisAddIn
    {
        private AddinCore _core;

        private static void ThisAddIn_Startup(object sender, EventArgs e)
        {
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return RibbonFactory.Instance;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _core.Dispose();
        }

        private void InternalStartup()
        {
            _core = new AddinCore(Application.Session);
            RibbonFactory.Instance.InitialiseFactory(_core.Resolve<IEnumerable<IRibbonViewModel>>());
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
    }
}
