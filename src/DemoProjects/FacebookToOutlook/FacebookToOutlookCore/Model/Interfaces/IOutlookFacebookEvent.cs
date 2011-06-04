namespace FacebookToOutlookCore.Model.Interfaces
{
    public interface IOutlookFacebookEvent : IFacebookEvent
    {
        string EntryId { get; }
    }
}
