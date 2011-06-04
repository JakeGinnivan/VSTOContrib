using System;
using System.ComponentModel;
using FacebookToOutlookCore.Model.Interfaces;

namespace FacebookToOutlookCore.Model
{
    public class FacebookUser : IFacebookUser
    {
        private long _userId;
        private string _name;
        private string _company;
        private DateTime? _birthday;
        private Uri _picturePath;

        public long UserId
        {
            get { return _userId; }
            set
            {
                _userId = value;
                OnPropertyChanged("UserId");
            }
        }

        public string Name
        {
            get { return _name; }
            set
            {
                _name = value;
                OnPropertyChanged("Name");
            }
        }

        public string Company
        {
            get { return _company; }
            set
            {
                _company = value;
                OnPropertyChanged("Company");
            }
        }

        public DateTime? Birthday
        {
            get { return _birthday; }
            set
            {
                _birthday = value;
                OnPropertyChanged("Birthday");
            }
        }

        public Uri PictureUri
        {
            get { return _picturePath ?? new Uri("pack://application:,,,/FacebookToOutlook;component/Resources/photo_portrait.png", UriKind.Absolute); }
            set
            {
                _picturePath = value;
                OnPropertyChanged("PicturePath");
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