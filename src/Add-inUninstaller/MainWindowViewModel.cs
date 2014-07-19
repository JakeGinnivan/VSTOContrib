using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Collections.ObjectModel;
using Microsoft.Win32;
using System.IO;
using VSTOContrib.Core;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.Wpf;

namespace Add_inUninstaller
{
    public class MainWindowViewModel : INotifyPropertyChanged
    {
        private Addin selectedAddin;
        private readonly ObservableCollection<Addin> addins = new ObservableCollection<Addin>();

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
            //PowerPoint
            AddAddinsForPath("PowerPoint", Registry.CurrentUser, @"Software\Microsoft\Office\PowerPoint\Addins");
            AddAddinsForPath("PowerPoint", Registry.CurrentUser, @"Software\Wow6432Node\Microsoft\Office\PowerPoint\Addins");
            AddAddinsForPath("PowerPoint", Registry.LocalMachine, @"SOFTWARE\Microsoft\Office\PowerPoint\Addins");
            AddAddinsForPath("PowerPoint", Registry.LocalMachine, @"SOFTWARE\Wow6432Node\Microsoft\Office\PowerPoint\Addins");
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

                addins.Add(addin);
            }
        }

        public Addin SelectedAddin
        {
            get { return selectedAddin; }
            set 
            {
                selectedAddin = value;
                OnPropertyChanged("SelectedAddin");
                ((DelegateCommand) UninstallCommand).RaiseCanExecuteChanged();
            }
        }

        public ObservableCollection<Addin> Addins
        {
            get { return addins; }
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
                    var installerArgs = string.Format(" /U \"{0}\"", SelectedAddin.Manifest);

                    var vstoInstallerOutput = new StringBuilder();

                    var vstoStartInfo = new ProcessStartInfo(installerPath, installerArgs);
                    var returnCode = vstoStartInfo.StartProcess((sender, e) => vstoInstallerOutput.Append(e.Data));

                    var message = vstoInstallerOutput.ToString();

                    if (returnCode != 0)
                    {
                        SelectedAddin.RegistryKey.DeleteKey();
                        MessageBox.Show(string.Format(
                            "Add-in was not installed through VSTOInstaller (probably visual studio instead), {0} was manually removed", 
                            SelectedAddin.AddinName), "Success");
                    }
                    else
                    {
                        MessageBox.Show(string.Format("{0} uninstalled", SelectedAddin.AddinName), "Success");                        
                    }
                }
                else
                {
                    //Manifest doesn't exist, delete registry keys manually
                    SelectedAddin.RegistryKey.DeleteKey();
                    MessageBox.Show(string.Format("{0} uninstalled", SelectedAddin.AddinName), "Success");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Uninstalling");
                return;
            }

            addins.Remove(SelectedAddin);
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
