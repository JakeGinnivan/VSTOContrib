using System;
using Microsoft.Office.Interop.Outlook;
using VSTOContrib.Outlook;

namespace TwitterFeedOutlookAddin.Core
{
    public class ContactAdapter
    {
        private const string TwitterUsernameProperty = "TwitterUsername";
        private readonly ContactItem currentItem;

        public ContactAdapter(ContactItem currentItem)
        {
            this.currentItem = currentItem;
        }

        public string TwitterUsername
        {
            get
            {
                return currentItem.GetPropertyValue(TwitterUsernameProperty, OlUserPropertyType.olText, false, Convert.ToString, "");
            }
            set
            {
                currentItem.SetPropertyValue(TwitterUsernameProperty, OlUserPropertyType.olText, value, true);
            }
        }

        public ContactItem Contact
        {
            get { return currentItem; }
        }
    }
}