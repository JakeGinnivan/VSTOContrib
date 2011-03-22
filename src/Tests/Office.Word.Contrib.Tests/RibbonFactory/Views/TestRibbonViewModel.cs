using System.ComponentModel;
using Microsoft.Office.Core;
using Office.Contrib.RibbonFactory;
using Office.Contrib.RibbonFactory.Interfaces;
using Office.Word.Contrib.RibbonFactory;

namespace Office.Word.Contrib.Tests.RibbonFactory.Views
{
    [WordRibbonViewModel]
    public class TestRibbonViewModel : IRibbonViewModel, INotifyPropertyChanged
    {
        public IRibbonUI RibbonUi { get; set; }

        public void Displayed(object context)
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
    }
}