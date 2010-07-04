using System.ComponentModel;

namespace Office.Utility
{
    ///<summary>
    /// Enum representing the different Ribbon Types
    ///</summary>
    public enum RibbonType
    {
        ///<summary>
        /// Appointment Item Inspector
        ///</summary>
        [Description("Microsoft.Outlook.Appointment")]
        OutlookAppointment,
        /// <summary>
        /// Contact Item Inspector
        /// </summary>
        [Description("Microsoft.Outlook.Contact")]
        OutlookContact,
        /// <summary>
        /// Distribution List Inspector
        /// </summary>
        [Description("Microsoft.Outlook.DistributionList")]
        OutlookDistributionList,
        /// <summary>
        /// Outlook Explorer (Main Outlook Window) Explorer
        /// </summary>
        [Description("Microsoft.Outlook.Explorer")]
        OutlookExplorer,
        /// <summary>
        /// Journal Item Inspector
        /// </summary>
        [Description("Microsoft.Outlook.Journal")]
        OutlookJournal,
        /// <summary>
        /// Compose Mail Item Inspector
        /// </summary>
        [Description("Microsoft.Outlook.Mail.Compose")]
        OutlookMailCompose,
        /// <summary>
        /// Read Mail Item Inspector
        /// </summary>
        [Description("Microsoft.Outlook.Mail.Read")]
        OutlookMailRead,
        /// <summary>
        /// Read Meeting Request Inspector
        /// </summary>
        [Description("Microsoft.Outlook.MeetingRequest.Read")]
        OutlookMeetingRequestRead,
        /// <summary>
        /// Send Meeting Request Inspector
        /// </summary>
        [Description("Microsoft.Outlook.MeetingRequest.Send")]
        OutlookMeetingRequestSend,
        /// <summary>
        /// Compose Post Inspector
        /// </summary>
        [Description("Microsoft.Outlook.Post.Compose")]
        OutlookPostCompose,
        /// <summary>
        /// Read Post Inspector
        /// </summary>
        [Description("Microsoft.Outlook.Post.Read")]
        OutlookPostRead,
        /// <summary>
        /// 
        /// </summary>
        [Description("Microsoft.Outlook.Report")]
        OutlookReport,
        /// <summary>
        /// 
        /// </summary>
        [Description("Microsoft.Outlook.Resend")]
        OutlookResend,
        /// <summary>
        /// 
        /// </summary>
        [Description("Microsoft.Outlook.Response.Compose")]
        OutlookResponseCompose,
        /// <summary>
        /// 
        /// </summary>
        [Description("Microsoft.Outlook.Response.CounterPropose")]
        OutlookResponseCounterPropose,
        /// <summary>
        /// 
        /// </summary>
        [Description("Microsoft.Outlook.Response.Read")]
        OutlookResponseRead,
        /// <summary>
        /// Rss item Inspector
        /// </summary>
        [Description("Microsoft.Outlook.RSS")]
        OutlookRSS,
        /// <summary>
        /// 
        /// </summary>
        [Description("Microsoft.Outlook.Sharing.Compose")]
        OutlookSharingCompose,
        /// <summary>
        /// 
        /// </summary>
        [Description("Microsoft.Outlook.Sharing.Read")]
        OutlookSharingRead,
        /// <summary>
        /// Task Item Inspector
        /// </summary>
        [Description("Microsoft.Outlook.Task")]
        OutlookTask
    }
}