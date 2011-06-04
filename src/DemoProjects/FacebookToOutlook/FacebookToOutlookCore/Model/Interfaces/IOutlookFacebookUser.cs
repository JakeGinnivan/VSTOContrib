namespace FacebookToOutlookCore.Model.Interfaces
{
    public interface IOutlookFacebookUser : IFacebookUser
    {
        string EntryId { get; }
    }
}