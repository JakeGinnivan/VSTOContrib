using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VSTOContrib.Core;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;
using VSTOContrib.Core.Extensions;
using ExcelAddIn1.Properties;
using System.Drawing;
using System.Windows;
using System.Windows.Controls;
using VSTOContrib.Core.Wpf;
using Microsoft.Office.Tools;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public class DocumentViewModel : OfficeViewModelBase, IRibbonViewModel, IRegisterCustomTaskPane
    {
        bool panelShown, ribbonVisible;
        ICustomTaskPaneWrapper myAddinTaskPane;
        Microsoft.Office.Interop.Excel.Workbook workbook;

        public IRibbonUI RibbonUi { get; set; }

        public Factory VstoFactory { get; set; }

        public void Initialised(object context)
        {
            workbook = context as Microsoft.Office.Interop.Excel.Workbook;
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
                selection.Resource.Value = new[] { DateTime.Now.ToString() + " " + GetHashCode() };
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
