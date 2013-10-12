using System;
using System.Net;
using System.Runtime.InteropServices;
using System.Security;
using System.Windows.Controls;
using System.Windows.Input;
using IronGitHub;
using VSTOContrib.Core;
using VSTOContrib.Core.RibbonFactory.Internal;
using VSTOContrib.Core.Wpf;

namespace GitHubForOutlook.Core.Features.Settings
{
    public class SettingsViewModel : NotifyPropertyChanged, ISettingsViewModel
    {
        ICustomTaskPaneWrapper taskPane;
        readonly GitHubApi githubApi;
        Action loginCallback;

        public SettingsViewModel(GitHubApi githubApi)
        {
            this.githubApi = githubApi;
            LoginCommand = new DelegateCommand<PasswordBox>(Login, p=>!LoggingIn);
            ClearLoginDetailsCommand = new DelegateCommand(ClearLoginDetails);
        }

        void ClearLoginDetails()
        {
            Properties.Settings.Default.UserName = null;
            Properties.Settings.Default.AuthorisationId = null;
            Properties.Settings.Default.Save();
            HasLoginDetailsAlready = false;
        }

        async void Login(PasswordBox obj)
        {
            LoggingIn = true;
            var secureString = ConvertToUnsecureString(obj.SecurePassword);
            var result = await githubApi.Authorize(new NetworkCredential(Username, secureString), new[]
            {
                Scopes.Repo
            }, "GitHub for Outlook");

            Properties.Settings.Default.AuthorisationId = result.Id;
            Properties.Settings.Default.AuthToken = result.Token;
            Properties.Settings.Default.UserName = Username;
            Properties.Settings.Default.Save();
            taskPane.Visible = false;
            loginCallback();
        }

        public string Username { get; set; }
        public bool LoggingIn { get; set; }
        public bool HasLoginDetailsAlready { get; set; }

        public ICommand LoginCommand { get; private set; }
        public ICommand ClearLoginDetailsCommand { get; private set; }

        public string CurrentUsername
        {
            get { return Properties.Settings.Default.UserName; }
        }

        public void Init(ICustomTaskPaneWrapper settingsTaskPane)
        {
            taskPane = settingsTaskPane;
        }

        public void LoginCallback(Action action)
        {
            loginCallback = action;
        }

        public static string ConvertToUnsecureString(SecureString securePassword)
        {
            if (securePassword == null)
                throw new ArgumentNullException("securePassword");

            IntPtr unmanagedString = IntPtr.Zero;
            try
            {
                unmanagedString = Marshal.SecureStringToGlobalAllocUnicode(securePassword);
                return Marshal.PtrToStringUni(unmanagedString);
            }
            finally
            {
                Marshal.ZeroFreeGlobalAllocUnicode(unmanagedString);
            }
        }
    }


}
