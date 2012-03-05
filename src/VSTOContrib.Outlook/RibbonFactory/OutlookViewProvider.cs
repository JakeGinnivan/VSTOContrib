using System;
using Microsoft.Office.Interop.Outlook;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Outlook.RibbonFactory
{
    internal class OutlookViewProvider : IViewProvider<OutlookRibbonType>
    {
        private Explorers _explorers;
        private Inspectors _inspectors;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="outlookApplication"></param>
        public OutlookViewProvider(_Application outlookApplication) 
        {
            _explorers = outlookApplication.Explorers;
            _inspectors = outlookApplication.Inspectors;
        }

        private void RegisterExplorers()
        {
            _explorers.NewExplorer += NewExplorer;

            foreach (Explorer explorer in _explorers)
                NewExplorer(explorer);
        }

        private void RegisterInspectors()
        {
            _inspectors.NewInspector += NewInspector;

            foreach (Inspector inspector in _inspectors)
                NewInspector(inspector);
        }

        void NewInspector(Inspector inspector)
        {
            var handler = NewView;
            if (handler == null) return;

            var wrapper = new InspectorWrapper(inspector);
            wrapper.Closed += InspectorClosed;

            var ribbonType = InspectorToRibbonTypeConverter.Convert(inspector);
            var newViewEventArgs = new NewViewEventArgs<OutlookRibbonType>(inspector, wrapper.CurrentContext, ribbonType);
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

            var newViewEventArgs = new NewViewEventArgs<OutlookRibbonType>(explorer, wrapper.CurrentContext, OutlookRibbonType.OutlookExplorer);
            handler(this, newViewEventArgs);

            if (!newViewEventArgs.Handled)
                explorer.ReleaseComObject();
        }

        private void ExplorerClosed(object sender, ExplorerClosedEventArgs e)
        {
            var wrapper = (ExplorerWrapper)sender;
            wrapper.Closed -= ExplorerClosed;

            var handler = ViewClosed;

            if (handler != null)
                handler(this, new ViewClosedEventArgs(e.Explorer, e.CurrentContext));
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

        public event EventHandler<NewViewEventArgs<OutlookRibbonType>> NewView;
        public event EventHandler<ViewClosedEventArgs> ViewClosed;

        public void CleanupReferencesTo(object view, object context)
        {
        }

        public void Dispose()
        {
            _explorers.NewExplorer -= NewExplorer;
            _inspectors.NewInspector -= NewInspector;
            _explorers.ReleaseComObject();
            _inspectors.ReleaseComObject();
            _explorers = null;
            _inspectors = null;
        }
    }
}