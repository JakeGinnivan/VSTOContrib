using System;
using System.Collections.Generic;
using System.Windows;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Outlook;
using VSTOContrib.Core;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Outlook.RibbonFactory;
using Application = System.Windows.Application;

namespace OutlookQuickStart
{
    public partial class ThisAddIn
    {
        AddinBootstrapper core;

        void ThisAddInStartup(object sender, EventArgs e)
        {
            
            // Required for WPF Integration in Outlook
            if (System.Windows.Application.Current == null)
                new Application {ShutdownMode = ShutdownMode.OnExplicitShutdown};

            //To enable background checking of updates uncomment this code
            //new VstoClickOnceUpdater()
            //    .CheckForUpdateAsync(
            //        r =>
            //        {
            //            if (r.Updated)
            //            {
            //                MessageBox.Show(@"Add-in updated");
            //            }
            //        });
        }

        void ThisAddInShutdown(object sender, EventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new OutlookRibbonFactory(typeof(AddinBootstrapper).Assembly);
        }

        protected override object RequestService(Guid serviceGuid)
        {
            var service = base.RequestService(serviceGuid);
            if (serviceGuid == typeof(_FormRegionStartup).GUID)
            {
                var manager = (_FormRegionStartup)base.RequestService(serviceGuid);
                return new Wrapper(manager);
                //if (this.formRegionManager == null)

                //    this.formRegionManager = new RSSMessages.Regions.FormRegionManager();

                //return this.formRegionManager;

            }

            return base.RequestService(serviceGuid);
        }

        private void InternalStartup()
        {
            core = new AddinBootstrapper();
            OutlookRibbonFactory.SetApplication(Application);
            RibbonFactory.Current.InitialiseFactory(
                t => (IRibbonViewModel)core.Resolve(t),
                CustomTaskPanes);

            Startup += ThisAddInStartup;
            Shutdown += ThisAddInShutdown;
        }
    }

    public class Wrapper : FormRegionStartup
    {
        readonly _FormRegionStartup wrapped;

        public Wrapper(_FormRegionStartup wrapped)
        {
            this.wrapped = wrapped;
        }

        public object GetFormRegionStorage(string FormRegionName, object Item, int LCID, OlFormRegionMode FormRegionMode, OlFormRegionSize FormRegionSize)
        {
            return wrapped.GetFormRegionStorage(FormRegionName, Item, LCID, FormRegionMode, FormRegionSize);
        }

        public void BeforeFormRegionShow(FormRegion FormRegion)
        {
            wrapped.BeforeFormRegionShow(FormRegion);
        }

        public object GetFormRegionManifest(string FormRegionName, int LCID)
        {
            return wrapped.GetFormRegionManifest(FormRegionName, LCID);
        }

        public object GetFormRegionIcon(string FormRegionName, int LCID, OlFormRegionIcon Icon)
        {
            return wrapped.GetFormRegionIcon(FormRegionName, LCID, Icon);
        }
    }
}
