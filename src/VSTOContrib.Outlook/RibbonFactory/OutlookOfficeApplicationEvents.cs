using System;
using Microsoft.Office.Interop.Outlook;
using VSTOContrib.Core;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Outlook.RibbonFactory
{
    internal class OutlookOfficeApplicationEvents : IOfficeApplicationEvents
    {
        readonly string captionSuffix = string.Empty;
        const string OutlookLpClassName = "rctrl_renwnd32\0";
        Explorers explorers;
        Inspectors inspectors;

        private void RegisterExplorers()
        {
            explorers.NewExplorer += NewExplorer;

            foreach (Explorer explorer in explorers)
                NewExplorer(explorer);
        }

        private void RegisterInspectors()
        {
            inspectors.NewInspector += NewInspector;

            foreach (Inspector inspector in inspectors)
                NewInspector(inspector);
        }

        void NewInspector(Inspector inspector)
        {
            var wrapper = new InspectorWrapper(inspector);
            wrapper.Closed += InspectorClosed;

            var ribbonType = InspectorToRibbonTypeConverter.Convert(inspector);
            var newViewEventArgs = new NewViewEventArgs(ToOfficeWindow(inspector), wrapper.CurrentContext, ribbonType.GetEnumDescription());
            NewView(newViewEventArgs);
        }

        void NewExplorer(Explorer explorer)
        {
            var wrapper = new ExplorerWrapper(explorer);
            wrapper.Closed += ExplorerClosed;

            var newViewEventArgs = new NewViewEventArgs(ToOfficeWindow(explorer), explorer, OutlookRibbonType.OutlookExplorer.GetEnumDescription());
            NewView(newViewEventArgs);
        }

        private void ExplorerClosed(object sender, ExplorerEventArgs e)
        {
            var wrapper = (ExplorerWrapper)sender;
            wrapper.Closed -= ExplorerClosed;

            ViewClosed(ToOfficeWindow(e.Explorer));
        }

        void InspectorClosed(object sender, InspectorClosedEventArgs e)
        {
            var wrapper = (InspectorWrapper) sender;
            wrapper.Closed -= InspectorClosed;

            ViewClosed(ToOfficeWindow(e.Inspector));
        }

        public void Initialise(object application)
        {
            var outlookApplication = (_Application) application;
            explorers = outlookApplication.Explorers;
            inspectors = outlookApplication.Inspectors;
            RegisterExplorers();
            RegisterInspectors();
        }

        public event Action<NewViewEventArgs> NewView;
        public event Action<OfficeWin32Window> ViewClosed;
        public event Action<object> ContextClosed;

        public OfficeWin32Window ToOfficeWindow(object view)
        {
            return new OfficeWin32Window(view, OutlookLpClassName, captionSuffix);
        }

        public void Dispose()
        {
            explorers.NewExplorer -= NewExplorer;
            inspectors.NewInspector -= NewInspector;
            explorers.ReleaseComObject();
            inspectors.ReleaseComObject();
            explorers = null;
            inspectors = null;
        }
    }
}