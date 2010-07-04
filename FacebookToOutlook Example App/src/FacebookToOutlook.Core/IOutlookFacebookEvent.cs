namespace FacebookToOutlook.Core
{
    public interface IOutlookFacebookEvent : IFacebookEvent
    {
        string EntryId { get; }
    }
}
