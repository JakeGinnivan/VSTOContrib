using System;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using VSTOContrib.Core;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;
using VSTOContrib.Core.Wpf;
using Factory = Microsoft.Office.Tools.Factory;

namespace QuoteGeneratorAddin.Core.OfficeContexts
{
    public class DocumentViewModel : OfficeViewModelBase, IRibbonViewModel, IRegisterCustomTaskPane
    {
        bool panelShown, ribbonVisible;
        ICustomTaskPaneWrapper myAddinTaskPane;
        Workbook workbook;
        readonly IQuotesService quotes;

        public DocumentViewModel(IQuotesService quotes)
        {
            this.quotes = quotes;
        }

        public IRibbonUI RibbonUi { get; set; }

        public Factory VstoFactory { get; set; }

        public void Initialised(object context)
        {
            workbook = context as Workbook;
        }

        public bool RibbonVisible
        {
            get { return ribbonVisible; }
            set
            {
                ribbonVisible = value;
                OnPropertyChanged(()=>RibbonVisible);
            }
        }

        public void CurrentViewChanged(object currentView)
        {
            RibbonVisible = workbook != null;
            panelShown = workbook != null;
        }
        
        public bool PanelShown
        {
            get { return panelShown; }
            set
            {
                if (panelShown == value) return;
                panelShown = value;
                myAddinTaskPane.Visible = value;
                OnPropertyChanged(() => PanelShown);
            }
        }

        public void RegisterTaskPanes(Register register)
        {
            myAddinTaskPane = register(() =>
                {
                    var button = new System.Windows.Controls.Button
                    {
                        Content = "Insert Quote"
                    };
                    button.Click += InsertQuote;
                    var host = new WpfPanelHost
                    {
                        Child = new UserControl
                        {
                            Content = new StackPanel
                            {
                                Children =
                                {
                                    button
                                }
                            }
                        }
                    };
                    return host;
                }, "Quotes!");
            myAddinTaskPane.Visible = true;
            myAddinTaskPane.VisibleChanged += TaskPaneVisibleChanged;
            TaskPaneVisibleChanged(this, EventArgs.Empty);
        }

        void InsertQuote(object sender, RoutedEventArgs e)
        {
            using (var application = workbook.Application.WithComCleanup())
            using (var selection = ((Range)application.Resource.Selection).WithComCleanup())
            {
                selection.Resource.Value = new[] {quotes.GetQuote()};
            }
        }

        public void Cleanup()
        {
            myAddinTaskPane.VisibleChanged -= TaskPaneVisibleChanged;
        }

        private void TaskPaneVisibleChanged(object sender, EventArgs e)
        {
            panelShown = myAddinTaskPane.Visible;
            OnPropertyChanged(() => PanelShown);
        }
    }
}
