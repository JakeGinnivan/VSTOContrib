using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using Facebook;

namespace FacebookToOutlookCore.Views
{
    public partial class FacebookLoginView 
    {
        private readonly Uri navigateUrl;

        public FacebookLoginView(string appId, string[] extendedPermissions)
            : this(appId, extendedPermissions, false)
        {
        }

        public FacebookLoginView(string appId, string[] extendedPermissions, bool logout)
        {
            var oauth = new FacebookOAuthClient { AppId = appId };

            var loginParameters = new Dictionary<string, object>
                    {
                        { "response_type", "token" },
                        { "display", "popup" }
                    };

            if (extendedPermissions != null && extendedPermissions.Length > 0)
            {
                var scope = new StringBuilder();
                scope.Append(string.Join(",", extendedPermissions));
                loginParameters["scope"] = scope.ToString();
            }

            var loginUrl = oauth.GetLoginUrl(loginParameters);

            if (logout)
            {
                var logoutParameters = new Dictionary<string, object>
                                           {
                                               { "next", loginUrl }
                                           };

                navigateUrl = oauth.GetLogoutUrl(logoutParameters);
            }
            else
            {
                navigateUrl = loginUrl;
            }

            InitializeComponent();
        }

        private void WindowLoaded(object sender, RoutedEventArgs e)
        {
            browser.Navigate(navigateUrl.AbsoluteUri);
        }

        private void BrowserNavigated(object sender, System.Windows.Navigation.NavigationEventArgs e)
        {
            FacebookOAuthResult result;
            if (FacebookOAuthResult.TryParse(e.Uri, out result))
            {
                FacebookOAuthResult = result;
                DialogResult = result.IsSuccess;
            }
            else
            {
                FacebookOAuthResult = null;
            }
            Close();
        }

        public FacebookOAuthResult FacebookOAuthResult { get; private set; }
    }
}
