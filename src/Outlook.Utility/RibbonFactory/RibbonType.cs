using System;
using System.ComponentModel;

namespace Outlook.Utility.RibbonFactory
{
    ///<summary>
    /// Enum representing the different Ribbon Types
    ///</summary>
    [Flags]
    public enum RibbonType
    {
        ///<summary>
        /// Appointment Item Inspector
        ///</summary>
        [Description("Microsoft.Outlook.Appointment")]
        OutlookAppointment = 1,
        /// <summary>
        /// Contact Item Inspector
        /// </summary>
        [Description("Microsoft.Outlook.Contact")]
        OutlookContact = 1 << 1,
        /// <summary>
        /// Distribution List Inspector
        /// </summary>
        [Description("Microsoft.Outlook.DistributionList")]
        OutlookDistributionList = 1 << 2,
        /// <summary>
        /// Outlook Explorer (Main Outlook Window) Explorer
        /// </summary>
        [Description("Microsoft.Outlook.Explorer")]
        OutlookExplorer = 1 << 3,
        /// <summary>
        /// Journal Item Inspector
        /// </summary>
        [Description("Microsoft.Outlook.Journal")]
        OutlookJournal = 1 << 4,
        /// <summary>
        /// Compose Mail Item Inspector
        /// </summary>
        [Description("Microsoft.Outlook.Mail.Compose")]
        OutlookMailCompose = 1 << 5,
        /// <summary>
        /// Read Mail Item Inspector
        /// </summary>
        [Description("Microsoft.Outlook.Mail.Read")]
        OutlookMailRead = 1 << 6,
        /// <summary>
        /// Read Meeting Request Inspector
        /// </summary>
        [Description("Microsoft.Outlook.MeetingRequest.Read")]
        OutlookMeetingRequestRead = 1 << 7,
        /// <summary>
        /// Send Meeting Request Inspector
        /// </summary>
        [Description("Microsoft.Outlook.MeetingRequest.Send")]
        OutlookMeetingRequestSend = 1 << 8,
        /// <summary>
        /// Compose Post Inspector
        /// </summary>
        [Description("Microsoft.Outlook.Post.Compose")]
        OutlookPostCompose = 1 << 9,
        /// <summary>
        /// Read Post Inspector
        /// </summary>
        [Description("Microsoft.Outlook.Post.Read")]
        OutlookPostRead = 1 << 10,
        /// <summary>
        /// 
        /// </summary>
        [Description("Microsoft.Outlook.Report")]
        OutlookReport = 1 << 11,
        /// <summary>
        /// 
        /// </summary>
        [Description("Microsoft.Outlook.Resend")]
        OutlookResend = 1 << 12,
        /// <summary>
        /// 
        /// </summary>
        [Description("Microsoft.Outlook.Response.Compose")]
        OutlookResponseCompose = 1 << 13,
        /// <summary>
        /// 
        /// </summary>
        [Description("Microsoft.Outlook.Response.CounterPropose")]
        OutlookResponseCounterPropose = 1 << 14,
        /// <summary>
        /// 
        /// </summary>
        [Description("Microsoft.Outlook.Response.Read")]
        OutlookResponseRead = 1 << 15,
        /// <summary>
        /// Rss item Inspector
        /// </summary>
        [Description("Microsoft.Outlook.RSS")]
        OutlookRSS = 1 << 16,
        /// <summary>
        /// 
        /// </summary>
        [Description("Microsoft.Outlook.Sharing.Compose")]
        OutlookSharingCompose = 1 << 17,
        /// <summary>
        /// 
        /// </summary>
        [Description("Microsoft.Outlook.Sharing.Read")]
        OutlookSharingRead = 1 << 18,
        /// <summary>
        /// Task Item Inspector
        /// </summary>
        [Description("Microsoft.Outlook.Task")]
        OutlookTask = 1 << 19,
    }
}