using System;
using System.Configuration;
using FacebookToOutlook.Core;
using FacebookToOutlook.Data;
using FacebookToOutlook.Services;
using Autofac;
using Office.Outlook.Contrib.Services;

namespace FacebookToOutlook.Properties
{
    internal sealed partial class Settings : ISyncSettings, IConfigurationSettings, IApplicationSettings,
        IEventConfigurationSettings, IContactConfigurationSettings, ISynchronisedEventInfo
    {
        [UserScopedSettingAttribute]
        public DateTime? LastSync
        {
            get
            {
                return ((DateTime?)(this["LastSync"]));
            }
            set 
            {
                this["LastSync"] = value;
            }
        }

        [UserScopedSettingAttribute]
        public bool MarkAsPrivate
        {
            get { return (bool?)this["MarkAsPrivate"] ?? false; }
            set { this["MarkAsPrivate"] = value; }
        }

        [UserScopedSettingAttribute]
        public bool EventReminder
        {
            get { return (bool?)this["EventReminder"] ?? true; }
            set { this["EventReminder"] = value; }
        }

        [UserScopedSettingAttribute]
        public int RemindMinutesBefore
        {
            get { return (int?)this["RemindMinutesBefore"] ?? 30; }
            set { this["RemindMinutesBefore"] = value; }
        }

        [UserScopedSettingAttribute]
        public int ShowTimeAsValue
        {
            get { return (int?)this["ShowTimeAsValue"] ?? (int)BusyStatus.Free; }
            set { this["ShowTimeAsValue"] = value; }
        }

        public BusyStatus ShowTimeAs
        {
            get { return (BusyStatus)ShowTimeAsValue; }
            set { ShowTimeAsValue = (int)value; }
        }

        [UserScopedSettingAttribute]
        public int DownloadTypesValue
        {
            get
            {
                return (int?)this["DownloadTypesValue"] ?? (int)(RsvpStatus.Attending | RsvpStatus.NotReplied | RsvpStatus.Unsure);
            }
            set { this["DownloadTypesValue"] = value; }
        }

        public RsvpStatus DownloadTypes
        {
            get { return (RsvpStatus)DownloadTypesValue; }
            set { DownloadTypesValue = (int)value; }
        }

        [UserScopedSettingAttribute]
        public string Category
        {
            get { return (string)this["Category"] ?? "Facebook Event"; }
            set { this["Category"] = value; }
        }

        [UserScopedSettingAttribute]
        public int AppointmentTaskPaneWidth
        {
            get { return (int?)this["AppointmentTaskPaneWidth"] ?? 150; }
            set { this["AppointmentTaskPaneWidth"] = value; }
        }

        public IEventConfigurationSettings EventConfigurationSettings
        {
            get { return this; }
        }

        public IContactConfigurationSettings ContactConfigurationSettings
        {
            get { return this; }
        }
    }
}
