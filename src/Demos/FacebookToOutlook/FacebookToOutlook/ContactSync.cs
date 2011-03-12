using System.Collections.Generic;
using FacebookToOutlook.Core;
using FacebookToOutlook.Data;
using FacebookToOutlook.Presentation.ViewModels.ContactSync;

namespace FacebookToOutlook.Presentation
{
    public class ContactSync
    {
        private readonly IOutlookRepository _outlookRepository;

        public ContactSync(IOutlookRepository outlookRepository)
        {
            _outlookRepository = outlookRepository;
        }

        public void Sync(IEnumerable<MatchedUserViewModel> matchedFacebookUsers)
        {
            var usersToSave = new List<IOutlookFacebookUser>();
            foreach (var matchedUserViewModel in matchedFacebookUsers)
            {
                matchedUserViewModel.OutlookContact.Birthday = matchedUserViewModel.MatchingFacebookContact.Birthday;
                matchedUserViewModel.OutlookContact.PictureUri = matchedUserViewModel.MatchingFacebookContact.PictureUri;
                usersToSave.Add(matchedUserViewModel.OutlookContact);
            }

            _outlookRepository.SaveOutlookContacts(usersToSave);
        }
    }
}