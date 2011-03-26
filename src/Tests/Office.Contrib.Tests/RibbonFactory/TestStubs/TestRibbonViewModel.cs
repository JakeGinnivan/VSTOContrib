using System;
using System.ComponentModel;
using Microsoft.Office.Core;
using Office.Contrib.RibbonFactory;
using Office.Contrib.RibbonFactory.Interfaces;

namespace Office.Contrib.Tests.RibbonFactory.TestStubs
{
    [RibbonViewModel(TestRibbonTypes.RibbonType1)]
    public class TestRibbonViewModel : IRibbonViewModel, INotifyPropertyChanged
    {
        public IRibbonUI RibbonUi { get; set; }

        public void Initialised(object context)
        {
        }

        public void CurrentViewChanged(object currentView)
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