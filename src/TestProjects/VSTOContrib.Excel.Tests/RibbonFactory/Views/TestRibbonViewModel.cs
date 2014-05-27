using System.ComponentModel;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Excel.RibbonFactory;

namespace VSTOContrib.Excel.Tests.RibbonFactory.Views
{
    [ExcelRibbonViewModel]
    public class TestRibbonViewModel : IRibbonViewModel, INotifyPropertyChanged, IRegisterCustomTaskPane
    {
        public IRibbonUI RibbonUi { get; set; }
        public Factory VstoFactory { get; set; }
        public object CurrentView { get; set; }

        public void Initialised(object context)
        {
        }

        public bool PanelShown { get; set; }

        public bool MyActionEnabled(IRibbonControl control)
        {
            return true;
        }

        public void Cleanup()
        {
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, e);
        }

        public void RegisterTaskPanes(Register register)
        {
            register(() => null, "Title");
        }
    }
}