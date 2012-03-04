using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;
using VSTOContrib.Core.Wpf;
using VSTOContrib.Word.RibbonFactory;

namespace RazorDocs.Core
{
    [WordRibbonViewModel]
    public class RazorDocs : OfficeViewModelBase, IRibbonViewModel, IRegisterCustomTaskPane
    {
        private bool _panelShown;
        private Document _document;
        private ICustomTaskPaneWrapper _razorDocsTaskPane;

        public void Initialised(object context)
        {
            _document = (Document)context;
        }

        public void CurrentViewChanged(object currentView)
        {
        }

        public IRibbonUI RibbonUi { get; set; }

        public bool PanelShown
        {
            get { return _panelShown; }
            set
            {
                if (_panelShown == value) return;
                _panelShown = value;
                _razorDocsTaskPane.Visible = value;
                OnPropertyChanged(() => PanelShown);
            }
        }

        public void RegisterTaskPanes(Register register)
        {
            _razorDocsTaskPane = register(
                () => new WpfPanelHost
                {
                    Child = new RazorDocsPanel
                    {
                        //DataContext = new RazorDocsPanelViewModel(this)
                    }
                }, "RazorDocs");
            _razorDocsTaskPane.Visible = true;
            PanelShown = true;
            _razorDocsTaskPane.VisibleChanged += RazorDocsTaskPaneVisibleChanged;
            RazorDocsTaskPaneVisibleChanged(this, EventArgs.Empty);
        }

        public void Cleanup()
        {
            _razorDocsTaskPane.VisibleChanged -= RazorDocsTaskPaneVisibleChanged;
        }

        private void RazorDocsTaskPaneVisibleChanged(object sender, EventArgs e)
        {
            _panelShown = _razorDocsTaskPane.Visible;
            OnPropertyChanged(() => PanelShown);
        }
    }
}
