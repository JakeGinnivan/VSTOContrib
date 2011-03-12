namespace FacebookToOutlook.Core
{
    public interface IOutlookFacebookUser : IFacebookUser
    {
        string EntryId { get; }
    }
}