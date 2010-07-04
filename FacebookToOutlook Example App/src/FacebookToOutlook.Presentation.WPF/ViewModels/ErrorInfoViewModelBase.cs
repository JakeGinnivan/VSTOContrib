using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using GalaSoft.MvvmLight;

namespace FacebookToOutlook.Presentation.ViewModels
{
    public class ErrorInfoViewModelBase : ViewModelBase, IDataErrorInfo
    {
        protected bool ThrowOnInvalidPropertyName { get; set; }
        private readonly Dictionary<string, string> _errors = new Dictionary<string, string>();

        protected ErrorInfoViewModelBase()
        {
            PropertyChanged += ViewModelBase_PropertyChanged;
        }

        void ViewModelBase_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "IsValid" || e.PropertyName == "Error" || e.PropertyName == string.Empty) return;

            RaisePropertyChanged("IsValid");
            RaisePropertyChanged("Error");
        }

        protected void SetError(string propertyName, string error)
        {
            if (_errors.ContainsKey(propertyName))
                _errors[propertyName] = error;
            else
                _errors.Add(propertyName, error);
        }
        protected void ClearError(string propertyName)
        {
            if (_errors.ContainsKey(propertyName))
                _errors.Remove(propertyName);
        }

        public bool IsValid
        {
            get
            {
                return _errors.Count == 0;
            }
        }

        public string this[string columnName]
        {
            get
            {
                return _errors.ContainsKey(columnName) ? _errors[columnName] : string.Empty;
            }
        }

        public string Error
        {
            get { return string.Join("\r\n", _errors.Select(e => string.Format(" - {0}: {1}", e.Key, e.Value))); }
        }
    }
}
