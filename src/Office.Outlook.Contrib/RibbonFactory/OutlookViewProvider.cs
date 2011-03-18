using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Outlook;
using Office.Contrib.Extensions;
using Office.Contrib.RibbonFactory;

namespace Office.Outlook.Contrib.RibbonFactory
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

            ((InspectorEvents_10_Event)inspector).Close += ViewClose;

            var ribbonType = InspectorToRibbonTypeConverter.Convert(inspector);
            var newViewEventArgs = new NewViewEventArgs<OutlookRibbonType>(inspector, ribbonType);
            handler(this, newViewEventArgs);

            if (!newViewEventArgs.Handled)
                inspector.ReleaseComObject();
        }

        void NewExplorer(Explorer explorer)
        {
            var handler = NewView;
            if (handler == null) return;

            ((ExplorerEvents_10_Event)explorer).Close += ViewClose;

            var newViewEventArgs = new NewViewEventArgs<OutlookRibbonType>(explorer, OutlookRibbonType.OutlookExplorer);
            handler(this, newViewEventArgs);

            if (!newViewEventArgs.Handled)
                explorer.ReleaseComObject();
        }

        void ViewClose()
        {

            var handler = ViewClosed;

            if (handler != null)
                handler(this, new ViewClosedEventArgs(_inspectors.Cast<object>()));
        }

        public void Initialise()
        {
            RegisterExplorers();
            RegisterInspectors();
        }

        public event EventHandler<NewViewEventArgs<OutlookRibbonType>> NewView;
        public event EventHandler<ViewClosedEventArgs> ViewClosed;

        public void CleanupReferencesTo(object view)
        {
            var inspectorEvent = view as InspectorEvents_10_Event;
            if (inspectorEvent != null)
                inspectorEvent.Close -= ViewClose;

            var explorerEvent = view as ExplorerEvents_10_Event;
            if (explorerEvent != null)
                explorerEvent.Close -= ViewClose;
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