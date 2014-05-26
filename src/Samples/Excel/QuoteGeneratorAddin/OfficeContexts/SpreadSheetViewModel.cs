using System;
using System.Drawing;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using QuoteGeneratorAddin.Core.OfficeContexts;
using QuoteGeneratorAddin.Properties;
using VSTOContrib.Core;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.Wpf;

namespace QuoteGeneratorAddin.OfficeContexts
{
    public class SpreadSheetViewModel : OfficeViewModelBase, IRibbonViewModel, IRegisterCustomTaskPane
    {
        bool panelShown, ribbonVisible;
        ICustomTaskPaneWrapper myAddinTaskPane;
        Workbook workbook;
        readonly IQuotesService quotes;

        public SpreadSheetViewModel(IQuotesService quotes)
        {
            this.quotes = quotes;
        }

        public IRibbonUI RibbonUi { get; set; }

        public Factory VstoFactory { get; set; }
        public object CurrentView { get; set; }

        public void Initialised(object context)
        {
            workbook = context as Workbook;
            RibbonVisible = workbook != null;
            panelShown = workbook != null;
        }

        public Bitmap ShowPanelImage { get { return Resources.icon; } }

        public bool RibbonVisible
        {
            get { return ribbonVisible; }
            set
            {
                ribbonVisible = value;
                OnPropertyChanged(() => RibbonVisible);
            }
        }

        public void ThrowError(IRibbonControl control)
        {
            throw null;
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
                selection.Resource.Value = new[] { quotes.GetQuote() };
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
