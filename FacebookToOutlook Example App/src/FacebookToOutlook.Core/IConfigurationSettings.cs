namespace FacebookToOutlook.Core
{
    public interface IConfigurationSettings
    {
        IEventConfigurationSettings EventConfigurationSettings { get; }
        IContactConfigurationSettings ContactConfigurationSettings { get; }
        void Save();
    }
}