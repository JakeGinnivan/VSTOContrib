using System;

namespace GitHubForOutlook.Core.Features.Settings
{
    class GitHubSettings : IGitHubSettings
    {
        public string AuthorisationId { get { return Properties.Settings.Default.AuthorisationId; } }
        public string AuthToken { get { return Properties.Settings.Default.AuthToken; } }
        public string Username { get { return Properties.Settings.Default.UserName; } }
        public bool LoginDetailsSet { get { return !string.IsNullOrEmpty(AuthToken); }}

        public void UpdateAuthInfo(string authorisationId, string authToken, string username)
        {
            Properties.Settings.Default.AuthorisationId = authorisationId;
            Properties.Settings.Default.AuthToken = authToken;
            Properties.Settings.Default.UserName = username;
            Properties.Settings.Default.Save();
            SettingsUpdated();
        }

        public void ClearAuthInfo()
        {
            Properties.Settings.Default.AuthorisationId = null;
            Properties.Settings.Default.AuthToken = null;
            Properties.Settings.Default.UserName = null;
            Properties.Settings.Default.Save();
            SettingsUpdated();
        }

        public event Action SettingsUpdated = () => { };
    }
}