namespace FacebookToOutlook.Core
{
    public interface IEventConfigurationSettings
    {
        bool MarkAsPrivate { get; set; }
        bool EventReminder { get; set; }
        int RemindMinutesBefore { get; set; }
        BusyStatus ShowTimeAs { get; set; }
        RsvpStatus DownloadTypes { get; set; }
        string Category { get; set; }
    }
}