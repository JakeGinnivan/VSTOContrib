using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace WordAddIn1
{
    public class MyAddinPanelViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public MyAddinPanelViewModel()
        {
            Text = "Test";
        }

        public string Text { get; set; }
    }
}
