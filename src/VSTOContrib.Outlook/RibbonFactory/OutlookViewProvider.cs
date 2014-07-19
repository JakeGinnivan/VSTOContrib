using System;
using Microsoft.Office.Interop.Outlook;
using VSTOContrib.Core;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Outlook.RibbonFactory
{
    internal class OutlookViewProvider : IViewProvider
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
            var newViewEventArgs = new NewViewEventArgs(new OfficeWin32Window(inspector, OutlookLpClassName, captionSuffix), wrapper.CurrentContext, ribbonType.GetEnumDescription());
            NewView(this, newViewEventArgs);

            if (!newViewEventArgs.Handled)
                inspector.ReleaseComObject();
        }

        void NewExplorer(Explorer explorer)
        {
            var wrapper = new ExplorerWrapper(explorer);
            wrapper.Closed += ExplorerClosed;

            var newViewEventArgs = new NewViewEventArgs(new OfficeWin32Window(explorer, OutlookLpClassName, captionSuffix), explorer, OutlookRibbonType.OutlookExplorer.GetEnumDescription());
            NewView(this, newViewEventArgs);

            if (!newViewEventArgs.Handled)
                explorer.ReleaseComObject();
        }

        private void ExplorerClosed(object sender, ExplorerEventArgs e)
        {
            var wrapper = (ExplorerWrapper)sender;
            wrapper.Closed -= ExplorerClosed;

            ViewClosed(this, new ViewClosedEventArgs(new OfficeWin32Window(e.Explorer, OutlookLpClassName, captionSuffix), e.Explorer));
        }

        void InspectorClosed(object sender, InspectorClosedEventArgs e)
        {
            var wrapper = (InspectorWrapper) sender;
            wrapper.Closed -= InspectorClosed;

            ViewClosed(this, new ViewClosedEventArgs(new OfficeWin32Window(e.Inspector, OutlookLpClassName, captionSuffix), e.CurrentContext));
        }

        public void Initialise(object application)
        {
            var outlookApplication = (_Application) application;
            explorers = outlookApplication.Explorers;
            inspectors = outlookApplication.Inspectors;
            RegisterExplorers();
            RegisterInspectors();
        }

        public event EventHandler<NewViewEventArgs> NewView = (sender, args) => { };
        public event EventHandler<ViewClosedEventArgs> ViewClosed = (sender, args) => { };

        public void CleanupReferencesTo(OfficeWin32Window view, object context)
        {
        }

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