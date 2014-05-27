using System;

namespace GitHubForOutlookAddin.Features.Settings
{
    public interface IGitHubSettings
    {
        string AuthorisationId { get; }
        string AuthToken { get; }
        string Username { get; }
        bool LoginDetailsSet { get; }
        void UpdateAuthInfo(string authorisationId, string authToken, string username);
        void ClearAuthInfo();
        event Action SettingsUpdated;
    }
}