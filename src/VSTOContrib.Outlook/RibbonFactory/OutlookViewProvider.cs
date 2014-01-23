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
        Explorers explorers;
        Inspectors inspectors;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="outlookApplication"></param>
        public OutlookViewProvider(_Application outlookApplication) 
        {
            explorers = outlookApplication.Explorers;
            inspectors = outlookApplication.Inspectors;
        }

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
            var handler = NewView;
            if (handler == null) return;

            var wrapper = new InspectorWrapper(inspector);
            wrapper.Closed += InspectorClosed;

            var ribbonType = InspectorToRibbonTypeConverter.Convert(inspector);
            var newViewEventArgs = new NewViewEventArgs(inspector, wrapper.CurrentContext, ribbonType.GetEnumDescription());
            handler(this, newViewEventArgs);

            if (!newViewEventArgs.Handled)
                inspector.ReleaseComObject();
        }

        void NewExplorer(Explorer explorer)
        {
            var handler = NewView;
            if (handler == null) return;

            var wrapper = new ExplorerWrapper(explorer);
            wrapper.Closed += ExplorerClosed;

            var newViewEventArgs = new NewViewEventArgs(explorer, explorer, OutlookRibbonType.OutlookExplorer.GetEnumDescription());
            handler(this, newViewEventArgs);

            if (!newViewEventArgs.Handled)
                explorer.ReleaseComObject();
        }

        private void ExplorerClosed(object sender, ExplorerEventArgs e)
        {
            var wrapper = (ExplorerWrapper)sender;
            wrapper.Closed -= ExplorerClosed;

            var handler = ViewClosed;

            if (handler != null)
                handler(this, new ViewClosedEventArgs(e.Explorer, e.Explorer));
        }

        void InspectorClosed(object sender, InspectorClosedEventArgs e)
        {
            var wrapper = (InspectorWrapper) sender;
            wrapper.Closed -= InspectorClosed;

            var handler = ViewClosed;

            if (handler != null)
                handler(this, new ViewClosedEventArgs(e.Inspector, e.CurrentContext));
        }

        public void Initialise()
        {
            RegisterExplorers();
            RegisterInspectors();
        }

        public event EventHandler<NewViewEventArgs> NewView;
        public event EventHandler<ViewClosedEventArgs> ViewClosed;
        public event EventHandler<HideCustomTaskPanesForContextEventArgs> UpdateCustomTaskPanesVisibilityForContext;

        public void CleanupReferencesTo(object view, object context)
        {
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