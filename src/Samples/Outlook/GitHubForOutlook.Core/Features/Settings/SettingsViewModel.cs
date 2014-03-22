using System;
using System.Net;
using System.Runtime.InteropServices;
using System.Security;
using System.Windows.Controls;
using System.Windows.Input;
using IronGitHub;
using VSTOContrib.Core;
using VSTOContrib.Core.Wpf;

namespace GitHubForOutlook.Core.Features.Settings
{
    public class SettingsViewModel : NotifyPropertyChanged, ISettingsViewModel
    {
        readonly GitHubApi githubApi;
        readonly IGitHubSettings settings;

        public SettingsViewModel(GitHubApi githubApi, IGitHubSettings settings)
        {
            this.githubApi = githubApi;
            this.settings = settings;
            LoginCommand = new DelegateCommand<PasswordBox>(Login, p => !LoggingIn);
            CancelCommand = new DelegateCommand(() => OnClose());
            ClearLoginDetailsCommand = new DelegateCommand(ClearLoginDetails);
        }

        void ClearLoginDetails()
        {
            settings.ClearAuthInfo();
        }

        async void Login(PasswordBox obj)
        {
            LoggingIn = true;
            var secureString = ConvertToUnsecureString(obj.SecurePassword);
            var result = await githubApi.Authorize(new NetworkCredential(Username, secureString), new[]
            {
                Scopes.Repo
            }, "GitHub for Outlook");

            settings.UpdateAuthInfo(result.Id, result.Token, Username);
            OnClose();
        }

        public string Username { get; set; }
        public bool LoggingIn { get; set; }

        public bool HasLoginDetailsAlready
        {
            get { return settings.LoginDetailsSet; }
        }

        public ICommand LoginCommand { get; private set; }
        public ICommand ClearLoginDetailsCommand { get; private set; }
        public ICommand CancelCommand { get; private set; }

        public string CurrentUsername
        {
            get { return settings.Username; }
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

        public event Action OnClose = () => { };
    }
}
