using System;
using System.ComponentModel;

namespace FacebookToOutlookCore.Model.Interfaces
{
    public interface IFacebookUser : INotifyPropertyChanged
    {
        long UserId { get; set; }
        string Name { get; set; }
        string Company {get;set;}
        DateTime? Birthday { get; set; }
        Uri PictureUri { get; set; }
    }
}
