using System;
using Microsoft.Office.Interop.Outlook;
using VSTOContrib.Outlook;

namespace TwitterFeedCore
{
    public class ContactAdapter
    {
        private readonly ContactItem _currentItem;
        private const string TwitterUsernameProperty = "TwitterUsername";

        public ContactAdapter(ContactItem currentItem)
        {
            _currentItem = currentItem;
        }

        public string TwitterUsername
        {
            get
            {
                return _currentItem.GetPropertyValue(TwitterUsernameProperty, OlUserPropertyType.olText, false, Convert.ToString, "");                
            }
            set
            {
                _currentItem.SetPropertyValue(TwitterUsernameProperty, OlUserPropertyType.olText, value, true);
            }
        }

        public ContactItem Contact
        {
            get { return _currentItem; }
        }
    }
}