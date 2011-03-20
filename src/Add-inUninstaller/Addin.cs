using System;
using System.ComponentModel;
using Microsoft.Win32;

namespace Add_inUninstaller
{
    public class Addin : INotifyPropertyChanged
    {
        private string _name;
        private string _description;
        private string _manifest;
        private bool _manifestExists;
        private string _product;
        private RegistryKey _registryKey;
        public event PropertyChangedEventHandler PropertyChanged;

        public string AddinName
        {
            get { return _name; }
            set
            {
                _name = value;
                OnPropertyChanged("Name");
            }
        }

        public string Description
        {
            get { return _description; }
            set
            {
                _description = value;
                OnPropertyChanged("Description");
            }
        }

        public string Manifest
        {
            get { return _manifest; }
            set
            {
                _manifest = value;
                OnPropertyChanged("Manifest");
            }
        }

        public bool ManifestExists
        {
            get { return _manifestExists; }
            set
            {
                _manifestExists = value;
                OnPropertyChanged("ManifestExists");
            }
        }

        public string Product
        {
            get { return _product; }
            set
            {
                _product = value;
                OnPropertyChanged("Product");
            }
        }

        public RegistryKey RegistryKey
        {
            get { return _registryKey; }
            set
            {
                _registryKey = value;
                OnPropertyChanged("RegistryKey");
            }
        }

        private void OnPropertyChanged(string propertyName)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}