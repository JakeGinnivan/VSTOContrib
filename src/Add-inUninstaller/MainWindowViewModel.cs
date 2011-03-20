using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Collections.ObjectModel;
using Microsoft.Win32;
using Office.Contrib;
using Office.Contrib.Extensions;
using System.IO;

namespace Add_inUninstaller
{
    public class MainWindowViewModel : INotifyPropertyChanged
    {
        private Addin _selectedAddin;
        private readonly ObservableCollection<Addin> _addins = new ObservableCollection<Addin>();

        public MainWindowViewModel()
        {
            UninstallCommand = new DelegateCommand(Uninstall, CanUninstall);
            LoadAddins();
        }

        private void LoadAddins()
        {
            //Outlook
            AddAddinsForPath("Outlook", Registry.CurrentUser, @"Software\Microsoft\Office\Outlook\Addins");
            AddAddinsForPath("Outlook", Registry.CurrentUser, @"Software\Wow6432Node\Microsoft\Office\Outlook\Addins");
            AddAddinsForPath("Outlook", Registry.LocalMachine, @"SOFTWARE\Microsoft\Office\Outlook\Addins");
            AddAddinsForPath("Outlook", Registry.LocalMachine, @"SOFTWARE\Wow6432Node\Microsoft\Office\Outlook\Addins");
            //Word
            AddAddinsForPath("Word", Registry.CurrentUser, @"Software\Microsoft\Office\Word\Addins");
            AddAddinsForPath("Word", Registry.CurrentUser, @"Software\Wow6432Node\Microsoft\Office\Word\Addins");
            AddAddinsForPath("Word", Registry.LocalMachine, @"SOFTWARE\Microsoft\Office\Word\Addins");
            AddAddinsForPath("Word", Registry.LocalMachine, @"SOFTWARE\Wow6432Node\Microsoft\Office\Word\Addins");
            //Excel
            AddAddinsForPath("Excel", Registry.CurrentUser, @"Software\Microsoft\Office\Excel\Addins");
            AddAddinsForPath("Excel", Registry.CurrentUser, @"Software\Wow6432Node\Microsoft\Office\Excel\Addins");
            AddAddinsForPath("Excel", Registry.LocalMachine, @"SOFTWARE\Microsoft\Office\Excel\Addins");
            AddAddinsForPath("Excel", Registry.LocalMachine, @"SOFTWARE\Wow6432Node\Microsoft\Office\Excel\Addins");
        }

        private void AddAddinsForPath(string product, RegistryKey startignKey, string path)
        {
            if (!startignKey.Exists(path)) return;
            var addinsKey = startignKey.OpenSubKey(path);

            foreach (var subKeyName in addinsKey.GetSubKeyNames())
            {
                var subKey = addinsKey.OpenSubKey(subKeyName);
                var manifest = subKey.GetValue("Manifest");

                if (manifest == null) continue;

                var manifestFile = manifest
                    .ToString()
                    .Replace("|vstolocal", string.Empty)
                    .Replace("file:///", string.Empty);

                var addin = new Addin
                {
                    AddinName = subKey.GetValue("FriendlyName").ToString(),
                    Description = subKey.GetValue("Description").ToString(),
                    RegistryKey = subKey,
                    Product = product,
                    Manifest = manifestFile,
                    ManifestExists = File.Exists(manifestFile)
                };

                _addins.Add(addin);
            }
        }

        public Addin SelectedAddin
        {
            get { return _selectedAddin; }
            set 
            {
                _selectedAddin = value;
                OnPropertyChanged("SelectedAddin");
                ((DelegateCommand) UninstallCommand).RaiseCanExecuteChanged();
            }
        }

        public ObservableCollection<Addin> Addins
        {
            get { return _addins; }
        }

        private bool CanUninstall()
        {
            return SelectedAddin != null;
        }

        private void Uninstall()
        {
            try
            {
                if (SelectedAddin.ManifestExists)
                {
                    var installerPath = VstoClickOnceUpdater.GetInstallerPath();

                    if (installerPath == null)
                    {
                        throw new InvalidOperationException("Cannot find VSTO Installer");
                    }
                    var installerArgs = string.Format(" /U {0}", SelectedAddin.Manifest);

                    var vstoInstallerOutput = new StringBuilder();

                    var vstoStartInfo = new ProcessStartInfo(installerPath, installerArgs);
                    var returnCode = vstoStartInfo.StartProcess((sender, e) => vstoInstallerOutput.Append(e.Data));

                    var message = vstoInstallerOutput.ToString();

                    Debug.WriteLine("VSTO Installer Returned: {0}", returnCode);
                    Debug.Write(message);
                }
                else
                {
                    //Manifest doesn't exist, delete registry keys manually
                    SelectedAddin.RegistryKey.DeleteKey();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Uninstalling");
                return;
            }

            MessageBox.Show(string.Format("{0} uninstalled", SelectedAddin.AddinName), "Success");
            _addins.Remove(SelectedAddin);
            SelectedAddin = null;
        }

        public ICommand UninstallCommand { get; private set; }
        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
