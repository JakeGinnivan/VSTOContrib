using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using FacebookToOutlook.Core;
using Microsoft.Office.Interop.Outlook;
using Office.Utility.Extensions;
using Outlook.Utility.Extensions;

namespace FacebookToOutlook.Data.Adapters
{
    public class FacebookUserAdapter : IOutlookFacebookUser
    {
        private readonly _ContactItem _contactItem;
        private Uri _picturePath;
        private static readonly string PictureCachePath;
        public const string IsLinkedToFacebookUserProperty = "IsLinkedToFacebookUser";
        public const string FacebookUserIdProperty = "FacebookUserId";

        static FacebookUserAdapter()
        {
            try
            {
                PictureCachePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "FacebookToOutlook");
                if (!Directory.Exists(PictureCachePath)) Directory.CreateDirectory(PictureCachePath);
            }
            catch (IOException)
            {}
        }

        public FacebookUserAdapter(_ContactItem contactItem)
        {
            _contactItem = contactItem;
        }

        public string EntryId
        {
            get { return _contactItem.EntryID; }
        }

        public bool IsLinkedToFacebookUser
        {
            get
            {
                return _contactItem.GetPropertyValue(IsLinkedToFacebookUserProperty, OlUserPropertyType.olYesNo, false, Convert.ToBoolean, false);
            }
            private set
            {
                _contactItem.SetPropertyValue(IsLinkedToFacebookUserProperty, OlUserPropertyType.olYesNo, value, true);
                OnPropertyChanged("IsLinkedToFacebookUser");
            }
        }

        public long UserId
        {
            get
            {
                return _contactItem.GetPropertyValue(FacebookUserIdProperty, OlUserPropertyType.olText, false, Convert.ToInt64, -1);
            }
            set
            {
                _contactItem.SetPropertyValue(FacebookUserIdProperty, OlUserPropertyType.olText, value, true);
                IsLinkedToFacebookUser = (value != -1);
                OnPropertyChanged("UserId");
            }
        }

        public string Name
        {
            get { return _contactItem.FullName; }
            set
            {
                _contactItem.FullName = value;
                OnPropertyChanged("Name");
            }
        }

        public string Company
        {
            get { return _contactItem.CompanyName; }
            set
            {
                _contactItem.CompanyName = value;
                OnPropertyChanged("Company");
            }
        }

        public DateTime? Birthday
        {
            get { return _contactItem.Birthday; }
            set
            {
                _contactItem.Birthday = value ?? DateTime.MinValue;
                OnPropertyChanged("Birthday");
            }
        }

        public Uri PictureUri
        {
            get
            {
                if (!_contactItem.HasPicture)
                    return null;
                if (_picturePath != null)
                    return _picturePath;

                using (var attachments = _contactItem.Attachments.WithComCleanup())
                {
                    foreach (var attachment in attachments.ComLinq<Attachment>()
                        .Where(att => att.DisplayName == "ContactPicture.jpg"))
                    {
                        try
                        {
                            _picturePath = new Uri(Path.Combine(
                                PictureCachePath,
                                string.Format("Contact_{0}.jpg", _contactItem.EntryID)));

                            if (!File.Exists(_picturePath.LocalPath))
                                attachment.SaveAsFile(_picturePath.LocalPath);
                        }
                        catch
                        {
                            _picturePath = null;
                        }
                    }
                }

                return _picturePath;
            }
            set
            {
                using (var client = new WebClient())
                {
                    switch (value.Scheme)
                    {
                        case "http":
                            _picturePath = new Uri(Path.Combine(
                                PictureCachePath,
                                string.Format("Contact_{0}.jpg", EntryId)));
                            client.DownloadFile(value, _picturePath.AbsolutePath);
                            break;
                        case "pack":
                            return;
                        default:
                            _picturePath = value;
                            break;
                    }

                    _contactItem.AddPicture(_picturePath.AbsolutePath);
                    OnPropertyChanged("PicturePath");
                }
            }
        }

        private void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
