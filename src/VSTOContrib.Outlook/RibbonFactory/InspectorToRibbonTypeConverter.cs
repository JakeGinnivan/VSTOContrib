using System;
using System.Reflection;
using Microsoft.Office.Interop.Outlook;

namespace VSTOContrib.Outlook.RibbonFactory
{
    /// <summary>
    /// Makes a best attempt to go from an inspector to the appropriate ribbon type
    /// http://msdn.microsoft.com/en-us/library/bb176446(office.12).aspx is reference for Outlook Message classes
    /// Using logic from http://www.add-in-express.com/creating-addins-blog/2009/04/22/outlook-regions-read-compose/
    /// </summary>
    public static class InspectorToRibbonTypeConverter
    {
        /// <summary>
        /// Converts the specified inspector.
        /// </summary>
        /// <param name="inspector">The inspector.</param>
        /// <returns></returns>
        public static OutlookRibbonType Convert(Inspector inspector)
        {
            object item = inspector.CurrentItem;
            var type = item.GetType();

            var messageClass = (string)type.InvokeMember(
                "MessageClass",
                BindingFlags.GetProperty,
                null, item, null);

            if (messageClass.StartsWith("IPM.Appointment"))
                return OutlookRibbonType.OutlookAppointment;
            if (messageClass.StartsWith("IPM.Contact"))
                return OutlookRibbonType.OutlookContact;
            if (messageClass.StartsWith("IPM.Activity"))
                return OutlookRibbonType.OutlookJournal;
            if (messageClass.StartsWith("IPM.Note"))
                return ConvertMail(item);
            if (messageClass.StartsWith("IPM.Schedule.Meeting"))
                return ConvertMeeting(item);
            if (messageClass.StartsWith("IPM.Post"))
                return ConvertPost(item);
            if (messageClass.StartsWith("IPM.Task"))
                return OutlookRibbonType.OutlookTask;
            if (messageClass.StartsWith("IPM.OLE.CLASS"))
                return OutlookRibbonType.OutlookAppointment;

            throw new ArgumentOutOfRangeException(string.Format("MessageClass {0} is unrecognised", messageClass));
        }

        private static OutlookRibbonType ConvertPost(object item)
        {
            var post = (PostItem) item;

            return !post.Saved || post.Size == 0 ? OutlookRibbonType.OutlookPostCompose : OutlookRibbonType.OutlookPostRead;
        }

        private static OutlookRibbonType ConvertMeeting(object item)
        {
            var meeting = (MeetingItem) item;

            return meeting.Sent ? OutlookRibbonType.OutlookMeetingRequestRead : OutlookRibbonType.OutlookMeetingRequestSend;
        }

        private static OutlookRibbonType ConvertMail(object item)
        {
            var mail = (MailItem) item;

            return mail.Sent ? OutlookRibbonType.OutlookMailRead : OutlookRibbonType.OutlookMailCompose;
        }
    }
}
