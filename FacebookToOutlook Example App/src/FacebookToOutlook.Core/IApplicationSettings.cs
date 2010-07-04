namespace FacebookToOutlook.Properties
{
    public interface IApplicationSettings
    {
        int AppointmentTaskPaneWidth { get; set; }
        void Save();
    }
}