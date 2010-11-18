using Office.Utility.Extensions;
using Microsoft.Office.Interop.Outlook;
using Outlook.Utility.Interfaces;

namespace Outlook.Utility.Extensions
{
	/// <summary>
	/// Provides cleanup extension methods for interfaces exposed by Microsoft.Office.Interop.Outlook.dll
	/// </summary>
	public static class OutlookCleanupExtensions
	{
		/// <summary>
		/// Wrapper interface for _IRecipientControl which adds IDispose to the interface
		/// </summary>
		public static I_IRecipientControl WithComCleanup(this _IRecipientControl resource)
		{
			return resource.WithComCleanup<_IRecipientControl, I_IRecipientControl>();
		}

	/// <summary>
		/// Wrapper interface for _DRecipientControl which adds IDispose to the interface
		/// </summary>
		public static I_DRecipientControl WithComCleanup(this _DRecipientControl resource)
		{
			return resource.WithComCleanup<_DRecipientControl, I_DRecipientControl>();
		}

	/// <summary>
		/// Wrapper interface for _DRecipientControlEvents which adds IDispose to the interface
		/// </summary>
		public static I_DRecipientControlEvents WithComCleanup(this _DRecipientControlEvents resource)
		{
			return resource.WithComCleanup<_DRecipientControlEvents, I_DRecipientControlEvents>();
		}

	/// <summary>
		/// Wrapper interface for _DRecipientControlEvents_Event which adds IDispose to the interface
		/// </summary>
		public static I_DRecipientControlEvents_Event WithComCleanup(this _DRecipientControlEvents_Event resource)
		{
			return resource.WithComCleanup<_DRecipientControlEvents_Event, I_DRecipientControlEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for _RecipientControl which adds IDispose to the interface
		/// </summary>
		public static I_RecipientControl WithComCleanup(this _RecipientControl resource)
		{
			return resource.WithComCleanup<_RecipientControl, I_RecipientControl>();
		}

	/// <summary>
		/// Wrapper interface for _IDocSiteControl which adds IDispose to the interface
		/// </summary>
		public static I_IDocSiteControl WithComCleanup(this _IDocSiteControl resource)
		{
			return resource.WithComCleanup<_IDocSiteControl, I_IDocSiteControl>();
		}

	/// <summary>
		/// Wrapper interface for _DDocSiteControl which adds IDispose to the interface
		/// </summary>
		public static I_DDocSiteControl WithComCleanup(this _DDocSiteControl resource)
		{
			return resource.WithComCleanup<_DDocSiteControl, I_DDocSiteControl>();
		}

	/// <summary>
		/// Wrapper interface for _DDocSiteControlEvents which adds IDispose to the interface
		/// </summary>
		public static I_DDocSiteControlEvents WithComCleanup(this _DDocSiteControlEvents resource)
		{
			return resource.WithComCleanup<_DDocSiteControlEvents, I_DDocSiteControlEvents>();
		}

	/// <summary>
		/// Wrapper interface for _DDocSiteControlEvents_Event which adds IDispose to the interface
		/// </summary>
		public static I_DDocSiteControlEvents_Event WithComCleanup(this _DDocSiteControlEvents_Event resource)
		{
			return resource.WithComCleanup<_DDocSiteControlEvents_Event, I_DDocSiteControlEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for _DocSiteControl which adds IDispose to the interface
		/// </summary>
		public static I_DocSiteControl WithComCleanup(this _DocSiteControl resource)
		{
			return resource.WithComCleanup<_DocSiteControl, I_DocSiteControl>();
		}

	/// <summary>
		/// Wrapper interface for OlkControl which adds IDispose to the interface
		/// </summary>
		public static IOlkControl WithComCleanup(this OlkControl resource)
		{
			return resource.WithComCleanup<OlkControl, IOlkControl>();
		}

	/// <summary>
		/// Wrapper interface for _OlkTextBox which adds IDispose to the interface
		/// </summary>
		public static I_OlkTextBox WithComCleanup(this _OlkTextBox resource)
		{
			return resource.WithComCleanup<_OlkTextBox, I_OlkTextBox>();
		}

	/// <summary>
		/// Wrapper interface for OlkTextBoxEvents which adds IDispose to the interface
		/// </summary>
		public static IOlkTextBoxEvents WithComCleanup(this OlkTextBoxEvents resource)
		{
			return resource.WithComCleanup<OlkTextBoxEvents, IOlkTextBoxEvents>();
		}

	/// <summary>
		/// Wrapper interface for OlkTextBoxEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IOlkTextBoxEvents_Event WithComCleanup(this OlkTextBoxEvents_Event resource)
		{
			return resource.WithComCleanup<OlkTextBoxEvents_Event, IOlkTextBoxEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for OlkTextBox which adds IDispose to the interface
		/// </summary>
		public static IOlkTextBox WithComCleanup(this OlkTextBox resource)
		{
			return resource.WithComCleanup<OlkTextBox, IOlkTextBox>();
		}

	/// <summary>
		/// Wrapper interface for _OlkLabel which adds IDispose to the interface
		/// </summary>
		public static I_OlkLabel WithComCleanup(this _OlkLabel resource)
		{
			return resource.WithComCleanup<_OlkLabel, I_OlkLabel>();
		}

	/// <summary>
		/// Wrapper interface for OlkLabelEvents which adds IDispose to the interface
		/// </summary>
		public static IOlkLabelEvents WithComCleanup(this OlkLabelEvents resource)
		{
			return resource.WithComCleanup<OlkLabelEvents, IOlkLabelEvents>();
		}

	/// <summary>
		/// Wrapper interface for OlkLabelEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IOlkLabelEvents_Event WithComCleanup(this OlkLabelEvents_Event resource)
		{
			return resource.WithComCleanup<OlkLabelEvents_Event, IOlkLabelEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for OlkLabel which adds IDispose to the interface
		/// </summary>
		public static IOlkLabel WithComCleanup(this OlkLabel resource)
		{
			return resource.WithComCleanup<OlkLabel, IOlkLabel>();
		}

	/// <summary>
		/// Wrapper interface for _OlkCommandButton which adds IDispose to the interface
		/// </summary>
		public static I_OlkCommandButton WithComCleanup(this _OlkCommandButton resource)
		{
			return resource.WithComCleanup<_OlkCommandButton, I_OlkCommandButton>();
		}

	/// <summary>
		/// Wrapper interface for OlkCommandButtonEvents which adds IDispose to the interface
		/// </summary>
		public static IOlkCommandButtonEvents WithComCleanup(this OlkCommandButtonEvents resource)
		{
			return resource.WithComCleanup<OlkCommandButtonEvents, IOlkCommandButtonEvents>();
		}

	/// <summary>
		/// Wrapper interface for OlkCommandButtonEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IOlkCommandButtonEvents_Event WithComCleanup(this OlkCommandButtonEvents_Event resource)
		{
			return resource.WithComCleanup<OlkCommandButtonEvents_Event, IOlkCommandButtonEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for OlkCommandButton which adds IDispose to the interface
		/// </summary>
		public static IOlkCommandButton WithComCleanup(this OlkCommandButton resource)
		{
			return resource.WithComCleanup<OlkCommandButton, IOlkCommandButton>();
		}

	/// <summary>
		/// Wrapper interface for _OlkCheckBox which adds IDispose to the interface
		/// </summary>
		public static I_OlkCheckBox WithComCleanup(this _OlkCheckBox resource)
		{
			return resource.WithComCleanup<_OlkCheckBox, I_OlkCheckBox>();
		}

	/// <summary>
		/// Wrapper interface for OlkCheckBoxEvents which adds IDispose to the interface
		/// </summary>
		public static IOlkCheckBoxEvents WithComCleanup(this OlkCheckBoxEvents resource)
		{
			return resource.WithComCleanup<OlkCheckBoxEvents, IOlkCheckBoxEvents>();
		}

	/// <summary>
		/// Wrapper interface for OlkCheckBoxEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IOlkCheckBoxEvents_Event WithComCleanup(this OlkCheckBoxEvents_Event resource)
		{
			return resource.WithComCleanup<OlkCheckBoxEvents_Event, IOlkCheckBoxEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for OlkCheckBox which adds IDispose to the interface
		/// </summary>
		public static IOlkCheckBox WithComCleanup(this OlkCheckBox resource)
		{
			return resource.WithComCleanup<OlkCheckBox, IOlkCheckBox>();
		}

	/// <summary>
		/// Wrapper interface for _OlkOptionButton which adds IDispose to the interface
		/// </summary>
		public static I_OlkOptionButton WithComCleanup(this _OlkOptionButton resource)
		{
			return resource.WithComCleanup<_OlkOptionButton, I_OlkOptionButton>();
		}

	/// <summary>
		/// Wrapper interface for OlkOptionButtonEvents which adds IDispose to the interface
		/// </summary>
		public static IOlkOptionButtonEvents WithComCleanup(this OlkOptionButtonEvents resource)
		{
			return resource.WithComCleanup<OlkOptionButtonEvents, IOlkOptionButtonEvents>();
		}

	/// <summary>
		/// Wrapper interface for OlkOptionButtonEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IOlkOptionButtonEvents_Event WithComCleanup(this OlkOptionButtonEvents_Event resource)
		{
			return resource.WithComCleanup<OlkOptionButtonEvents_Event, IOlkOptionButtonEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for OlkOptionButton which adds IDispose to the interface
		/// </summary>
		public static IOlkOptionButton WithComCleanup(this OlkOptionButton resource)
		{
			return resource.WithComCleanup<OlkOptionButton, IOlkOptionButton>();
		}

	/// <summary>
		/// Wrapper interface for _OlkComboBox which adds IDispose to the interface
		/// </summary>
		public static I_OlkComboBox WithComCleanup(this _OlkComboBox resource)
		{
			return resource.WithComCleanup<_OlkComboBox, I_OlkComboBox>();
		}

	/// <summary>
		/// Wrapper interface for OlkComboBoxEvents which adds IDispose to the interface
		/// </summary>
		public static IOlkComboBoxEvents WithComCleanup(this OlkComboBoxEvents resource)
		{
			return resource.WithComCleanup<OlkComboBoxEvents, IOlkComboBoxEvents>();
		}

	/// <summary>
		/// Wrapper interface for OlkComboBoxEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IOlkComboBoxEvents_Event WithComCleanup(this OlkComboBoxEvents_Event resource)
		{
			return resource.WithComCleanup<OlkComboBoxEvents_Event, IOlkComboBoxEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for OlkComboBox which adds IDispose to the interface
		/// </summary>
		public static IOlkComboBox WithComCleanup(this OlkComboBox resource)
		{
			return resource.WithComCleanup<OlkComboBox, IOlkComboBox>();
		}

	/// <summary>
		/// Wrapper interface for _OlkListBox which adds IDispose to the interface
		/// </summary>
		public static I_OlkListBox WithComCleanup(this _OlkListBox resource)
		{
			return resource.WithComCleanup<_OlkListBox, I_OlkListBox>();
		}

	/// <summary>
		/// Wrapper interface for OlkListBoxEvents which adds IDispose to the interface
		/// </summary>
		public static IOlkListBoxEvents WithComCleanup(this OlkListBoxEvents resource)
		{
			return resource.WithComCleanup<OlkListBoxEvents, IOlkListBoxEvents>();
		}

	/// <summary>
		/// Wrapper interface for OlkListBoxEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IOlkListBoxEvents_Event WithComCleanup(this OlkListBoxEvents_Event resource)
		{
			return resource.WithComCleanup<OlkListBoxEvents_Event, IOlkListBoxEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for OlkListBox which adds IDispose to the interface
		/// </summary>
		public static IOlkListBox WithComCleanup(this OlkListBox resource)
		{
			return resource.WithComCleanup<OlkListBox, IOlkListBox>();
		}

	/// <summary>
		/// Wrapper interface for _OlkInfoBar which adds IDispose to the interface
		/// </summary>
		public static I_OlkInfoBar WithComCleanup(this _OlkInfoBar resource)
		{
			return resource.WithComCleanup<_OlkInfoBar, I_OlkInfoBar>();
		}

	/// <summary>
		/// Wrapper interface for OlkInfoBarEvents which adds IDispose to the interface
		/// </summary>
		public static IOlkInfoBarEvents WithComCleanup(this OlkInfoBarEvents resource)
		{
			return resource.WithComCleanup<OlkInfoBarEvents, IOlkInfoBarEvents>();
		}

	/// <summary>
		/// Wrapper interface for OlkInfoBarEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IOlkInfoBarEvents_Event WithComCleanup(this OlkInfoBarEvents_Event resource)
		{
			return resource.WithComCleanup<OlkInfoBarEvents_Event, IOlkInfoBarEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for OlkInfoBar which adds IDispose to the interface
		/// </summary>
		public static IOlkInfoBar WithComCleanup(this OlkInfoBar resource)
		{
			return resource.WithComCleanup<OlkInfoBar, IOlkInfoBar>();
		}

	/// <summary>
		/// Wrapper interface for _OlkContactPhoto which adds IDispose to the interface
		/// </summary>
		public static I_OlkContactPhoto WithComCleanup(this _OlkContactPhoto resource)
		{
			return resource.WithComCleanup<_OlkContactPhoto, I_OlkContactPhoto>();
		}

	/// <summary>
		/// Wrapper interface for OlkContactPhotoEvents which adds IDispose to the interface
		/// </summary>
		public static IOlkContactPhotoEvents WithComCleanup(this OlkContactPhotoEvents resource)
		{
			return resource.WithComCleanup<OlkContactPhotoEvents, IOlkContactPhotoEvents>();
		}

	/// <summary>
		/// Wrapper interface for OlkContactPhotoEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IOlkContactPhotoEvents_Event WithComCleanup(this OlkContactPhotoEvents_Event resource)
		{
			return resource.WithComCleanup<OlkContactPhotoEvents_Event, IOlkContactPhotoEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for OlkContactPhoto which adds IDispose to the interface
		/// </summary>
		public static IOlkContactPhoto WithComCleanup(this OlkContactPhoto resource)
		{
			return resource.WithComCleanup<OlkContactPhoto, IOlkContactPhoto>();
		}

	/// <summary>
		/// Wrapper interface for _OlkBusinessCardControl which adds IDispose to the interface
		/// </summary>
		public static I_OlkBusinessCardControl WithComCleanup(this _OlkBusinessCardControl resource)
		{
			return resource.WithComCleanup<_OlkBusinessCardControl, I_OlkBusinessCardControl>();
		}

	/// <summary>
		/// Wrapper interface for OlkBusinessCardControlEvents which adds IDispose to the interface
		/// </summary>
		public static IOlkBusinessCardControlEvents WithComCleanup(this OlkBusinessCardControlEvents resource)
		{
			return resource.WithComCleanup<OlkBusinessCardControlEvents, IOlkBusinessCardControlEvents>();
		}

	/// <summary>
		/// Wrapper interface for OlkBusinessCardControlEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IOlkBusinessCardControlEvents_Event WithComCleanup(this OlkBusinessCardControlEvents_Event resource)
		{
			return resource.WithComCleanup<OlkBusinessCardControlEvents_Event, IOlkBusinessCardControlEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for OlkBusinessCardControl which adds IDispose to the interface
		/// </summary>
		public static IOlkBusinessCardControl WithComCleanup(this OlkBusinessCardControl resource)
		{
			return resource.WithComCleanup<OlkBusinessCardControl, IOlkBusinessCardControl>();
		}

	/// <summary>
		/// Wrapper interface for _OlkPageControl which adds IDispose to the interface
		/// </summary>
		public static I_OlkPageControl WithComCleanup(this _OlkPageControl resource)
		{
			return resource.WithComCleanup<_OlkPageControl, I_OlkPageControl>();
		}

	/// <summary>
		/// Wrapper interface for OlkPageControlEvents which adds IDispose to the interface
		/// </summary>
		public static IOlkPageControlEvents WithComCleanup(this OlkPageControlEvents resource)
		{
			return resource.WithComCleanup<OlkPageControlEvents, IOlkPageControlEvents>();
		}

	/// <summary>
		/// Wrapper interface for OlkPageControlEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IOlkPageControlEvents_Event WithComCleanup(this OlkPageControlEvents_Event resource)
		{
			return resource.WithComCleanup<OlkPageControlEvents_Event, IOlkPageControlEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for OlkPageControl which adds IDispose to the interface
		/// </summary>
		public static IOlkPageControl WithComCleanup(this OlkPageControl resource)
		{
			return resource.WithComCleanup<OlkPageControl, IOlkPageControl>();
		}

	/// <summary>
		/// Wrapper interface for _OlkDateControl which adds IDispose to the interface
		/// </summary>
		public static I_OlkDateControl WithComCleanup(this _OlkDateControl resource)
		{
			return resource.WithComCleanup<_OlkDateControl, I_OlkDateControl>();
		}

	/// <summary>
		/// Wrapper interface for OlkDateControlEvents which adds IDispose to the interface
		/// </summary>
		public static IOlkDateControlEvents WithComCleanup(this OlkDateControlEvents resource)
		{
			return resource.WithComCleanup<OlkDateControlEvents, IOlkDateControlEvents>();
		}

	/// <summary>
		/// Wrapper interface for OlkDateControlEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IOlkDateControlEvents_Event WithComCleanup(this OlkDateControlEvents_Event resource)
		{
			return resource.WithComCleanup<OlkDateControlEvents_Event, IOlkDateControlEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for OlkDateControl which adds IDispose to the interface
		/// </summary>
		public static IOlkDateControl WithComCleanup(this OlkDateControl resource)
		{
			return resource.WithComCleanup<OlkDateControl, IOlkDateControl>();
		}

	/// <summary>
		/// Wrapper interface for _OlkTimeControl which adds IDispose to the interface
		/// </summary>
		public static I_OlkTimeControl WithComCleanup(this _OlkTimeControl resource)
		{
			return resource.WithComCleanup<_OlkTimeControl, I_OlkTimeControl>();
		}

	/// <summary>
		/// Wrapper interface for OlkTimeControlEvents which adds IDispose to the interface
		/// </summary>
		public static IOlkTimeControlEvents WithComCleanup(this OlkTimeControlEvents resource)
		{
			return resource.WithComCleanup<OlkTimeControlEvents, IOlkTimeControlEvents>();
		}

	/// <summary>
		/// Wrapper interface for OlkTimeControlEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IOlkTimeControlEvents_Event WithComCleanup(this OlkTimeControlEvents_Event resource)
		{
			return resource.WithComCleanup<OlkTimeControlEvents_Event, IOlkTimeControlEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for OlkTimeControl which adds IDispose to the interface
		/// </summary>
		public static IOlkTimeControl WithComCleanup(this OlkTimeControl resource)
		{
			return resource.WithComCleanup<OlkTimeControl, IOlkTimeControl>();
		}

	/// <summary>
		/// Wrapper interface for _OlkCategory which adds IDispose to the interface
		/// </summary>
		public static I_OlkCategory WithComCleanup(this _OlkCategory resource)
		{
			return resource.WithComCleanup<_OlkCategory, I_OlkCategory>();
		}

	/// <summary>
		/// Wrapper interface for OlkCategoryEvents which adds IDispose to the interface
		/// </summary>
		public static IOlkCategoryEvents WithComCleanup(this OlkCategoryEvents resource)
		{
			return resource.WithComCleanup<OlkCategoryEvents, IOlkCategoryEvents>();
		}

	/// <summary>
		/// Wrapper interface for OlkCategoryEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IOlkCategoryEvents_Event WithComCleanup(this OlkCategoryEvents_Event resource)
		{
			return resource.WithComCleanup<OlkCategoryEvents_Event, IOlkCategoryEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for OlkCategory which adds IDispose to the interface
		/// </summary>
		public static IOlkCategory WithComCleanup(this OlkCategory resource)
		{
			return resource.WithComCleanup<OlkCategory, IOlkCategory>();
		}

	/// <summary>
		/// Wrapper interface for _OlkFrameHeader which adds IDispose to the interface
		/// </summary>
		public static I_OlkFrameHeader WithComCleanup(this _OlkFrameHeader resource)
		{
			return resource.WithComCleanup<_OlkFrameHeader, I_OlkFrameHeader>();
		}

	/// <summary>
		/// Wrapper interface for OlkFrameHeaderEvents which adds IDispose to the interface
		/// </summary>
		public static IOlkFrameHeaderEvents WithComCleanup(this OlkFrameHeaderEvents resource)
		{
			return resource.WithComCleanup<OlkFrameHeaderEvents, IOlkFrameHeaderEvents>();
		}

	/// <summary>
		/// Wrapper interface for OlkFrameHeaderEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IOlkFrameHeaderEvents_Event WithComCleanup(this OlkFrameHeaderEvents_Event resource)
		{
			return resource.WithComCleanup<OlkFrameHeaderEvents_Event, IOlkFrameHeaderEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for OlkFrameHeader which adds IDispose to the interface
		/// </summary>
		public static IOlkFrameHeader WithComCleanup(this OlkFrameHeader resource)
		{
			return resource.WithComCleanup<OlkFrameHeader, IOlkFrameHeader>();
		}

	/// <summary>
		/// Wrapper interface for _OlkSenderPhoto which adds IDispose to the interface
		/// </summary>
		public static I_OlkSenderPhoto WithComCleanup(this _OlkSenderPhoto resource)
		{
			return resource.WithComCleanup<_OlkSenderPhoto, I_OlkSenderPhoto>();
		}

	/// <summary>
		/// Wrapper interface for OlkSenderPhotoEvents which adds IDispose to the interface
		/// </summary>
		public static IOlkSenderPhotoEvents WithComCleanup(this OlkSenderPhotoEvents resource)
		{
			return resource.WithComCleanup<OlkSenderPhotoEvents, IOlkSenderPhotoEvents>();
		}

	/// <summary>
		/// Wrapper interface for OlkSenderPhotoEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IOlkSenderPhotoEvents_Event WithComCleanup(this OlkSenderPhotoEvents_Event resource)
		{
			return resource.WithComCleanup<OlkSenderPhotoEvents_Event, IOlkSenderPhotoEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for OlkSenderPhoto which adds IDispose to the interface
		/// </summary>
		public static IOlkSenderPhoto WithComCleanup(this OlkSenderPhoto resource)
		{
			return resource.WithComCleanup<OlkSenderPhoto, IOlkSenderPhoto>();
		}

	/// <summary>
		/// Wrapper interface for _TimeZone which adds IDispose to the interface
		/// </summary>
		public static I_TimeZone WithComCleanup(this _TimeZone resource)
		{
			return resource.WithComCleanup<_TimeZone, I_TimeZone>();
		}

	/// <summary>
		/// Wrapper interface for _Application which adds IDispose to the interface
		/// </summary>
		public static I_Application WithComCleanup(this _Application resource)
		{
			return resource.WithComCleanup<_Application, I_Application>();
		}

	/// <summary>
		/// Wrapper interface for _NameSpace which adds IDispose to the interface
		/// </summary>
		public static I_NameSpace WithComCleanup(this _NameSpace resource)
		{
			return resource.WithComCleanup<_NameSpace, I_NameSpace>();
		}

	/// <summary>
		/// Wrapper interface for Recipient which adds IDispose to the interface
		/// </summary>
		public static IRecipient WithComCleanup(this Recipient resource)
		{
			return resource.WithComCleanup<Recipient, IRecipient>();
		}

	/// <summary>
		/// Wrapper interface for AddressEntry which adds IDispose to the interface
		/// </summary>
		public static IAddressEntry WithComCleanup(this AddressEntry resource)
		{
			return resource.WithComCleanup<AddressEntry, IAddressEntry>();
		}

	/// <summary>
		/// Wrapper interface for AddressEntries which adds IDispose to the interface
		/// </summary>
		public static IAddressEntries WithComCleanup(this AddressEntries resource)
		{
			return resource.WithComCleanup<AddressEntries, IAddressEntries>();
		}

	/// <summary>
		/// Wrapper interface for _ContactItem which adds IDispose to the interface
		/// </summary>
		public static I_ContactItem WithComCleanup(this _ContactItem resource)
		{
			return resource.WithComCleanup<_ContactItem, I_ContactItem>();
		}

	/// <summary>
		/// Wrapper interface for Actions which adds IDispose to the interface
		/// </summary>
		public static IActions WithComCleanup(this Actions resource)
		{
			return resource.WithComCleanup<Actions, IActions>();
		}

	/// <summary>
		/// Wrapper interface for Action which adds IDispose to the interface
		/// </summary>
		public static IAction WithComCleanup(this Action resource)
		{
			return resource.WithComCleanup<Action, IAction>();
		}

	/// <summary>
		/// Wrapper interface for Attachments which adds IDispose to the interface
		/// </summary>
		public static IAttachments WithComCleanup(this Attachments resource)
		{
			return resource.WithComCleanup<Attachments, IAttachments>();
		}

	/// <summary>
		/// Wrapper interface for Attachment which adds IDispose to the interface
		/// </summary>
		public static IAttachment WithComCleanup(this Attachment resource)
		{
			return resource.WithComCleanup<Attachment, IAttachment>();
		}

	/// <summary>
		/// Wrapper interface for PropertyAccessor which adds IDispose to the interface
		/// </summary>
		public static IPropertyAccessor WithComCleanup(this PropertyAccessor resource)
		{
			return resource.WithComCleanup<PropertyAccessor, IPropertyAccessor>();
		}

	/// <summary>
		/// Wrapper interface for _PropertyAccessor which adds IDispose to the interface
		/// </summary>
		public static I_PropertyAccessor WithComCleanup(this _PropertyAccessor resource)
		{
			return resource.WithComCleanup<_PropertyAccessor, I_PropertyAccessor>();
		}

	/// <summary>
		/// Wrapper interface for FormDescription which adds IDispose to the interface
		/// </summary>
		public static IFormDescription WithComCleanup(this FormDescription resource)
		{
			return resource.WithComCleanup<FormDescription, IFormDescription>();
		}

	/// <summary>
		/// Wrapper interface for _Inspector which adds IDispose to the interface
		/// </summary>
		public static I_Inspector WithComCleanup(this _Inspector resource)
		{
			return resource.WithComCleanup<_Inspector, I_Inspector>();
		}

	/// <summary>
		/// Wrapper interface for _AttachmentSelection which adds IDispose to the interface
		/// </summary>
		public static I_AttachmentSelection WithComCleanup(this _AttachmentSelection resource)
		{
			return resource.WithComCleanup<_AttachmentSelection, I_AttachmentSelection>();
		}

	/// <summary>
		/// Wrapper interface for Selection which adds IDispose to the interface
		/// </summary>
		public static ISelection WithComCleanup(this Selection resource)
		{
			return resource.WithComCleanup<Selection, ISelection>();
		}

	/// <summary>
		/// Wrapper interface for UserProperties which adds IDispose to the interface
		/// </summary>
		public static IUserProperties WithComCleanup(this UserProperties resource)
		{
			return resource.WithComCleanup<UserProperties, IUserProperties>();
		}

	/// <summary>
		/// Wrapper interface for UserProperty which adds IDispose to the interface
		/// </summary>
		public static IUserProperty WithComCleanup(this UserProperty resource)
		{
			return resource.WithComCleanup<UserProperty, IUserProperty>();
		}

	/// <summary>
		/// Wrapper interface for MAPIFolder which adds IDispose to the interface
		/// </summary>
		public static IMAPIFolder WithComCleanup(this MAPIFolder resource)
		{
			return resource.WithComCleanup<MAPIFolder, IMAPIFolder>();
		}

	/// <summary>
		/// Wrapper interface for _Folders which adds IDispose to the interface
		/// </summary>
		public static I_Folders WithComCleanup(this _Folders resource)
		{
			return resource.WithComCleanup<_Folders, I_Folders>();
		}

	/// <summary>
		/// Wrapper interface for _Items which adds IDispose to the interface
		/// </summary>
		public static I_Items WithComCleanup(this _Items resource)
		{
			return resource.WithComCleanup<_Items, I_Items>();
		}

	/// <summary>
		/// Wrapper interface for _Explorer which adds IDispose to the interface
		/// </summary>
		public static I_Explorer WithComCleanup(this _Explorer resource)
		{
			return resource.WithComCleanup<_Explorer, I_Explorer>();
		}

	/// <summary>
		/// Wrapper interface for Panes which adds IDispose to the interface
		/// </summary>
		public static IPanes WithComCleanup(this Panes resource)
		{
			return resource.WithComCleanup<Panes, IPanes>();
		}

	/// <summary>
		/// Wrapper interface for _NavigationPane which adds IDispose to the interface
		/// </summary>
		public static I_NavigationPane WithComCleanup(this _NavigationPane resource)
		{
			return resource.WithComCleanup<_NavigationPane, I_NavigationPane>();
		}

	/// <summary>
		/// Wrapper interface for NavigationModule which adds IDispose to the interface
		/// </summary>
		public static INavigationModule WithComCleanup(this NavigationModule resource)
		{
			return resource.WithComCleanup<NavigationModule, INavigationModule>();
		}

	/// <summary>
		/// Wrapper interface for _NavigationModule which adds IDispose to the interface
		/// </summary>
		public static I_NavigationModule WithComCleanup(this _NavigationModule resource)
		{
			return resource.WithComCleanup<_NavigationModule, I_NavigationModule>();
		}

	/// <summary>
		/// Wrapper interface for NavigationModules which adds IDispose to the interface
		/// </summary>
		public static INavigationModules WithComCleanup(this NavigationModules resource)
		{
			return resource.WithComCleanup<NavigationModules, INavigationModules>();
		}

	/// <summary>
		/// Wrapper interface for _NavigationModules which adds IDispose to the interface
		/// </summary>
		public static I_NavigationModules WithComCleanup(this _NavigationModules resource)
		{
			return resource.WithComCleanup<_NavigationModules, I_NavigationModules>();
		}

	/// <summary>
		/// Wrapper interface for _AccountSelector which adds IDispose to the interface
		/// </summary>
		public static I_AccountSelector WithComCleanup(this _AccountSelector resource)
		{
			return resource.WithComCleanup<_AccountSelector, I_AccountSelector>();
		}

	/// <summary>
		/// Wrapper interface for _Account which adds IDispose to the interface
		/// </summary>
		public static I_Account WithComCleanup(this _Account resource)
		{
			return resource.WithComCleanup<_Account, I_Account>();
		}

	/// <summary>
		/// Wrapper interface for Store which adds IDispose to the interface
		/// </summary>
		public static IStore WithComCleanup(this Store resource)
		{
			return resource.WithComCleanup<Store, IStore>();
		}

	/// <summary>
		/// Wrapper interface for _Store which adds IDispose to the interface
		/// </summary>
		public static I_Store WithComCleanup(this _Store resource)
		{
			return resource.WithComCleanup<_Store, I_Store>();
		}

	/// <summary>
		/// Wrapper interface for Rules which adds IDispose to the interface
		/// </summary>
		public static IRules WithComCleanup(this Rules resource)
		{
			return resource.WithComCleanup<Rules, IRules>();
		}

	/// <summary>
		/// Wrapper interface for _Rules which adds IDispose to the interface
		/// </summary>
		public static I_Rules WithComCleanup(this _Rules resource)
		{
			return resource.WithComCleanup<_Rules, I_Rules>();
		}

	/// <summary>
		/// Wrapper interface for _Rule which adds IDispose to the interface
		/// </summary>
		public static I_Rule WithComCleanup(this _Rule resource)
		{
			return resource.WithComCleanup<_Rule, I_Rule>();
		}

	/// <summary>
		/// Wrapper interface for RuleActions which adds IDispose to the interface
		/// </summary>
		public static IRuleActions WithComCleanup(this RuleActions resource)
		{
			return resource.WithComCleanup<RuleActions, IRuleActions>();
		}

	/// <summary>
		/// Wrapper interface for _RuleActions which adds IDispose to the interface
		/// </summary>
		public static I_RuleActions WithComCleanup(this _RuleActions resource)
		{
			return resource.WithComCleanup<_RuleActions, I_RuleActions>();
		}

	/// <summary>
		/// Wrapper interface for _RuleAction which adds IDispose to the interface
		/// </summary>
		public static I_RuleAction WithComCleanup(this _RuleAction resource)
		{
			return resource.WithComCleanup<_RuleAction, I_RuleAction>();
		}

	/// <summary>
		/// Wrapper interface for MoveOrCopyRuleAction which adds IDispose to the interface
		/// </summary>
		public static IMoveOrCopyRuleAction WithComCleanup(this MoveOrCopyRuleAction resource)
		{
			return resource.WithComCleanup<MoveOrCopyRuleAction, IMoveOrCopyRuleAction>();
		}

	/// <summary>
		/// Wrapper interface for _MoveOrCopyRuleAction which adds IDispose to the interface
		/// </summary>
		public static I_MoveOrCopyRuleAction WithComCleanup(this _MoveOrCopyRuleAction resource)
		{
			return resource.WithComCleanup<_MoveOrCopyRuleAction, I_MoveOrCopyRuleAction>();
		}

	/// <summary>
		/// Wrapper interface for RuleAction which adds IDispose to the interface
		/// </summary>
		public static IRuleAction WithComCleanup(this RuleAction resource)
		{
			return resource.WithComCleanup<RuleAction, IRuleAction>();
		}

	/// <summary>
		/// Wrapper interface for SendRuleAction which adds IDispose to the interface
		/// </summary>
		public static ISendRuleAction WithComCleanup(this SendRuleAction resource)
		{
			return resource.WithComCleanup<SendRuleAction, ISendRuleAction>();
		}

	/// <summary>
		/// Wrapper interface for _SendRuleAction which adds IDispose to the interface
		/// </summary>
		public static I_SendRuleAction WithComCleanup(this _SendRuleAction resource)
		{
			return resource.WithComCleanup<_SendRuleAction, I_SendRuleAction>();
		}

	/// <summary>
		/// Wrapper interface for Recipients which adds IDispose to the interface
		/// </summary>
		public static IRecipients WithComCleanup(this Recipients resource)
		{
			return resource.WithComCleanup<Recipients, IRecipients>();
		}

	/// <summary>
		/// Wrapper interface for AssignToCategoryRuleAction which adds IDispose to the interface
		/// </summary>
		public static IAssignToCategoryRuleAction WithComCleanup(this AssignToCategoryRuleAction resource)
		{
			return resource.WithComCleanup<AssignToCategoryRuleAction, IAssignToCategoryRuleAction>();
		}

	/// <summary>
		/// Wrapper interface for _AssignToCategoryRuleAction which adds IDispose to the interface
		/// </summary>
		public static I_AssignToCategoryRuleAction WithComCleanup(this _AssignToCategoryRuleAction resource)
		{
			return resource.WithComCleanup<_AssignToCategoryRuleAction, I_AssignToCategoryRuleAction>();
		}

	/// <summary>
		/// Wrapper interface for PlaySoundRuleAction which adds IDispose to the interface
		/// </summary>
		public static IPlaySoundRuleAction WithComCleanup(this PlaySoundRuleAction resource)
		{
			return resource.WithComCleanup<PlaySoundRuleAction, IPlaySoundRuleAction>();
		}

	/// <summary>
		/// Wrapper interface for _PlaySoundRuleAction which adds IDispose to the interface
		/// </summary>
		public static I_PlaySoundRuleAction WithComCleanup(this _PlaySoundRuleAction resource)
		{
			return resource.WithComCleanup<_PlaySoundRuleAction, I_PlaySoundRuleAction>();
		}

	/// <summary>
		/// Wrapper interface for MarkAsTaskRuleAction which adds IDispose to the interface
		/// </summary>
		public static IMarkAsTaskRuleAction WithComCleanup(this MarkAsTaskRuleAction resource)
		{
			return resource.WithComCleanup<MarkAsTaskRuleAction, IMarkAsTaskRuleAction>();
		}

	/// <summary>
		/// Wrapper interface for _MarkAsTaskRuleAction which adds IDispose to the interface
		/// </summary>
		public static I_MarkAsTaskRuleAction WithComCleanup(this _MarkAsTaskRuleAction resource)
		{
			return resource.WithComCleanup<_MarkAsTaskRuleAction, I_MarkAsTaskRuleAction>();
		}

	/// <summary>
		/// Wrapper interface for NewItemAlertRuleAction which adds IDispose to the interface
		/// </summary>
		public static INewItemAlertRuleAction WithComCleanup(this NewItemAlertRuleAction resource)
		{
			return resource.WithComCleanup<NewItemAlertRuleAction, INewItemAlertRuleAction>();
		}

	/// <summary>
		/// Wrapper interface for _NewItemAlertRuleAction which adds IDispose to the interface
		/// </summary>
		public static I_NewItemAlertRuleAction WithComCleanup(this _NewItemAlertRuleAction resource)
		{
			return resource.WithComCleanup<_NewItemAlertRuleAction, I_NewItemAlertRuleAction>();
		}

	/// <summary>
		/// Wrapper interface for RuleConditions which adds IDispose to the interface
		/// </summary>
		public static IRuleConditions WithComCleanup(this RuleConditions resource)
		{
			return resource.WithComCleanup<RuleConditions, IRuleConditions>();
		}

	/// <summary>
		/// Wrapper interface for _RuleConditions which adds IDispose to the interface
		/// </summary>
		public static I_RuleConditions WithComCleanup(this _RuleConditions resource)
		{
			return resource.WithComCleanup<_RuleConditions, I_RuleConditions>();
		}

	/// <summary>
		/// Wrapper interface for _RuleCondition which adds IDispose to the interface
		/// </summary>
		public static I_RuleCondition WithComCleanup(this _RuleCondition resource)
		{
			return resource.WithComCleanup<_RuleCondition, I_RuleCondition>();
		}

	/// <summary>
		/// Wrapper interface for RuleCondition which adds IDispose to the interface
		/// </summary>
		public static IRuleCondition WithComCleanup(this RuleCondition resource)
		{
			return resource.WithComCleanup<RuleCondition, IRuleCondition>();
		}

	/// <summary>
		/// Wrapper interface for ImportanceRuleCondition which adds IDispose to the interface
		/// </summary>
		public static IImportanceRuleCondition WithComCleanup(this ImportanceRuleCondition resource)
		{
			return resource.WithComCleanup<ImportanceRuleCondition, IImportanceRuleCondition>();
		}

	/// <summary>
		/// Wrapper interface for _ImportanceRuleCondition which adds IDispose to the interface
		/// </summary>
		public static I_ImportanceRuleCondition WithComCleanup(this _ImportanceRuleCondition resource)
		{
			return resource.WithComCleanup<_ImportanceRuleCondition, I_ImportanceRuleCondition>();
		}

	/// <summary>
		/// Wrapper interface for AccountRuleCondition which adds IDispose to the interface
		/// </summary>
		public static IAccountRuleCondition WithComCleanup(this AccountRuleCondition resource)
		{
			return resource.WithComCleanup<AccountRuleCondition, IAccountRuleCondition>();
		}

	/// <summary>
		/// Wrapper interface for _AccountRuleCondition which adds IDispose to the interface
		/// </summary>
		public static I_AccountRuleCondition WithComCleanup(this _AccountRuleCondition resource)
		{
			return resource.WithComCleanup<_AccountRuleCondition, I_AccountRuleCondition>();
		}

	/// <summary>
		/// Wrapper interface for Account which adds IDispose to the interface
		/// </summary>
		public static IAccount WithComCleanup(this Account resource)
		{
			return resource.WithComCleanup<Account, IAccount>();
		}

	/// <summary>
		/// Wrapper interface for TextRuleCondition which adds IDispose to the interface
		/// </summary>
		public static ITextRuleCondition WithComCleanup(this TextRuleCondition resource)
		{
			return resource.WithComCleanup<TextRuleCondition, ITextRuleCondition>();
		}

	/// <summary>
		/// Wrapper interface for _TextRuleCondition which adds IDispose to the interface
		/// </summary>
		public static I_TextRuleCondition WithComCleanup(this _TextRuleCondition resource)
		{
			return resource.WithComCleanup<_TextRuleCondition, I_TextRuleCondition>();
		}

	/// <summary>
		/// Wrapper interface for CategoryRuleCondition which adds IDispose to the interface
		/// </summary>
		public static ICategoryRuleCondition WithComCleanup(this CategoryRuleCondition resource)
		{
			return resource.WithComCleanup<CategoryRuleCondition, ICategoryRuleCondition>();
		}

	/// <summary>
		/// Wrapper interface for _CategoryRuleCondition which adds IDispose to the interface
		/// </summary>
		public static I_CategoryRuleCondition WithComCleanup(this _CategoryRuleCondition resource)
		{
			return resource.WithComCleanup<_CategoryRuleCondition, I_CategoryRuleCondition>();
		}

	/// <summary>
		/// Wrapper interface for FormNameRuleCondition which adds IDispose to the interface
		/// </summary>
		public static IFormNameRuleCondition WithComCleanup(this FormNameRuleCondition resource)
		{
			return resource.WithComCleanup<FormNameRuleCondition, IFormNameRuleCondition>();
		}

	/// <summary>
		/// Wrapper interface for _FormNameRuleCondition which adds IDispose to the interface
		/// </summary>
		public static I_FormNameRuleCondition WithComCleanup(this _FormNameRuleCondition resource)
		{
			return resource.WithComCleanup<_FormNameRuleCondition, I_FormNameRuleCondition>();
		}

	/// <summary>
		/// Wrapper interface for ToOrFromRuleCondition which adds IDispose to the interface
		/// </summary>
		public static IToOrFromRuleCondition WithComCleanup(this ToOrFromRuleCondition resource)
		{
			return resource.WithComCleanup<ToOrFromRuleCondition, IToOrFromRuleCondition>();
		}

	/// <summary>
		/// Wrapper interface for _ToOrFromRuleCondition which adds IDispose to the interface
		/// </summary>
		public static I_ToOrFromRuleCondition WithComCleanup(this _ToOrFromRuleCondition resource)
		{
			return resource.WithComCleanup<_ToOrFromRuleCondition, I_ToOrFromRuleCondition>();
		}

	/// <summary>
		/// Wrapper interface for AddressRuleCondition which adds IDispose to the interface
		/// </summary>
		public static IAddressRuleCondition WithComCleanup(this AddressRuleCondition resource)
		{
			return resource.WithComCleanup<AddressRuleCondition, IAddressRuleCondition>();
		}

	/// <summary>
		/// Wrapper interface for _AddressRuleCondition which adds IDispose to the interface
		/// </summary>
		public static I_AddressRuleCondition WithComCleanup(this _AddressRuleCondition resource)
		{
			return resource.WithComCleanup<_AddressRuleCondition, I_AddressRuleCondition>();
		}

	/// <summary>
		/// Wrapper interface for SenderInAddressListRuleCondition which adds IDispose to the interface
		/// </summary>
		public static ISenderInAddressListRuleCondition WithComCleanup(this SenderInAddressListRuleCondition resource)
		{
			return resource.WithComCleanup<SenderInAddressListRuleCondition, ISenderInAddressListRuleCondition>();
		}

	/// <summary>
		/// Wrapper interface for _SenderInAddressListRuleCondition which adds IDispose to the interface
		/// </summary>
		public static I_SenderInAddressListRuleCondition WithComCleanup(this _SenderInAddressListRuleCondition resource)
		{
			return resource.WithComCleanup<_SenderInAddressListRuleCondition, I_SenderInAddressListRuleCondition>();
		}

	/// <summary>
		/// Wrapper interface for AddressList which adds IDispose to the interface
		/// </summary>
		public static IAddressList WithComCleanup(this AddressList resource)
		{
			return resource.WithComCleanup<AddressList, IAddressList>();
		}

	/// <summary>
		/// Wrapper interface for FromRssFeedRuleCondition which adds IDispose to the interface
		/// </summary>
		public static IFromRssFeedRuleCondition WithComCleanup(this FromRssFeedRuleCondition resource)
		{
			return resource.WithComCleanup<FromRssFeedRuleCondition, IFromRssFeedRuleCondition>();
		}

	/// <summary>
		/// Wrapper interface for _FromRssFeedRuleCondition which adds IDispose to the interface
		/// </summary>
		public static I_FromRssFeedRuleCondition WithComCleanup(this _FromRssFeedRuleCondition resource)
		{
			return resource.WithComCleanup<_FromRssFeedRuleCondition, I_FromRssFeedRuleCondition>();
		}

	/// <summary>
		/// Wrapper interface for Rule which adds IDispose to the interface
		/// </summary>
		public static IRule WithComCleanup(this Rule resource)
		{
			return resource.WithComCleanup<Rule, IRule>();
		}

	/// <summary>
		/// Wrapper interface for Categories which adds IDispose to the interface
		/// </summary>
		public static ICategories WithComCleanup(this Categories resource)
		{
			return resource.WithComCleanup<Categories, ICategories>();
		}

	/// <summary>
		/// Wrapper interface for _Categories which adds IDispose to the interface
		/// </summary>
		public static I_Categories WithComCleanup(this _Categories resource)
		{
			return resource.WithComCleanup<_Categories, I_Categories>();
		}

	/// <summary>
		/// Wrapper interface for _Category which adds IDispose to the interface
		/// </summary>
		public static I_Category WithComCleanup(this _Category resource)
		{
			return resource.WithComCleanup<_Category, I_Category>();
		}

	/// <summary>
		/// Wrapper interface for Category which adds IDispose to the interface
		/// </summary>
		public static ICategory WithComCleanup(this Category resource)
		{
			return resource.WithComCleanup<Category, ICategory>();
		}

	/// <summary>
		/// Wrapper interface for View which adds IDispose to the interface
		/// </summary>
		public static IView WithComCleanup(this View resource)
		{
			return resource.WithComCleanup<View, IView>();
		}

	/// <summary>
		/// Wrapper interface for _Views which adds IDispose to the interface
		/// </summary>
		public static I_Views WithComCleanup(this _Views resource)
		{
			return resource.WithComCleanup<_Views, I_Views>();
		}

	/// <summary>
		/// Wrapper interface for _StorageItem which adds IDispose to the interface
		/// </summary>
		public static I_StorageItem WithComCleanup(this _StorageItem resource)
		{
			return resource.WithComCleanup<_StorageItem, I_StorageItem>();
		}

	/// <summary>
		/// Wrapper interface for Table which adds IDispose to the interface
		/// </summary>
		public static ITable WithComCleanup(this Table resource)
		{
			return resource.WithComCleanup<Table, ITable>();
		}

	/// <summary>
		/// Wrapper interface for _Table which adds IDispose to the interface
		/// </summary>
		public static I_Table WithComCleanup(this _Table resource)
		{
			return resource.WithComCleanup<_Table, I_Table>();
		}

	/// <summary>
		/// Wrapper interface for Row which adds IDispose to the interface
		/// </summary>
		public static IRow WithComCleanup(this Row resource)
		{
			return resource.WithComCleanup<Row, IRow>();
		}

	/// <summary>
		/// Wrapper interface for _Row which adds IDispose to the interface
		/// </summary>
		public static I_Row WithComCleanup(this _Row resource)
		{
			return resource.WithComCleanup<_Row, I_Row>();
		}

	/// <summary>
		/// Wrapper interface for Columns which adds IDispose to the interface
		/// </summary>
		public static IColumns WithComCleanup(this Columns resource)
		{
			return resource.WithComCleanup<Columns, IColumns>();
		}

	/// <summary>
		/// Wrapper interface for _Columns which adds IDispose to the interface
		/// </summary>
		public static I_Columns WithComCleanup(this _Columns resource)
		{
			return resource.WithComCleanup<_Columns, I_Columns>();
		}

	/// <summary>
		/// Wrapper interface for _Column which adds IDispose to the interface
		/// </summary>
		public static I_Column WithComCleanup(this _Column resource)
		{
			return resource.WithComCleanup<_Column, I_Column>();
		}

	/// <summary>
		/// Wrapper interface for Column which adds IDispose to the interface
		/// </summary>
		public static IColumn WithComCleanup(this Column resource)
		{
			return resource.WithComCleanup<Column, IColumn>();
		}

	/// <summary>
		/// Wrapper interface for CalendarSharing which adds IDispose to the interface
		/// </summary>
		public static ICalendarSharing WithComCleanup(this CalendarSharing resource)
		{
			return resource.WithComCleanup<CalendarSharing, ICalendarSharing>();
		}

	/// <summary>
		/// Wrapper interface for _CalendarSharing which adds IDispose to the interface
		/// </summary>
		public static I_CalendarSharing WithComCleanup(this _CalendarSharing resource)
		{
			return resource.WithComCleanup<_CalendarSharing, I_CalendarSharing>();
		}

	/// <summary>
		/// Wrapper interface for ItemEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IItemEvents_Event WithComCleanup(this ItemEvents_Event resource)
		{
			return resource.WithComCleanup<ItemEvents_Event, IItemEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for ItemEvents_10_Event which adds IDispose to the interface
		/// </summary>
		public static IItemEvents_10_Event WithComCleanup(this ItemEvents_10_Event resource)
		{
			return resource.WithComCleanup<ItemEvents_10_Event, IItemEvents_10_Event>();
		}

	/// <summary>
		/// Wrapper interface for MailItem which adds IDispose to the interface
		/// </summary>
		public static IMailItem WithComCleanup(this MailItem resource)
		{
			return resource.WithComCleanup<MailItem, IMailItem>();
		}

	/// <summary>
		/// Wrapper interface for _MailItem which adds IDispose to the interface
		/// </summary>
		public static I_MailItem WithComCleanup(this _MailItem resource)
		{
			return resource.WithComCleanup<_MailItem, I_MailItem>();
		}

	/// <summary>
		/// Wrapper interface for Links which adds IDispose to the interface
		/// </summary>
		public static ILinks WithComCleanup(this Links resource)
		{
			return resource.WithComCleanup<Links, ILinks>();
		}

	/// <summary>
		/// Wrapper interface for Link which adds IDispose to the interface
		/// </summary>
		public static ILink WithComCleanup(this Link resource)
		{
			return resource.WithComCleanup<Link, ILink>();
		}

	/// <summary>
		/// Wrapper interface for ItemProperties which adds IDispose to the interface
		/// </summary>
		public static IItemProperties WithComCleanup(this ItemProperties resource)
		{
			return resource.WithComCleanup<ItemProperties, IItemProperties>();
		}

	/// <summary>
		/// Wrapper interface for ItemProperty which adds IDispose to the interface
		/// </summary>
		public static IItemProperty WithComCleanup(this ItemProperty resource)
		{
			return resource.WithComCleanup<ItemProperty, IItemProperty>();
		}

	/// <summary>
		/// Wrapper interface for Conflicts which adds IDispose to the interface
		/// </summary>
		public static IConflicts WithComCleanup(this Conflicts resource)
		{
			return resource.WithComCleanup<Conflicts, IConflicts>();
		}

	/// <summary>
		/// Wrapper interface for Conflict which adds IDispose to the interface
		/// </summary>
		public static IConflict WithComCleanup(this Conflict resource)
		{
			return resource.WithComCleanup<Conflict, IConflict>();
		}

	/// <summary>
		/// Wrapper interface for ContactItem which adds IDispose to the interface
		/// </summary>
		public static IContactItem WithComCleanup(this ContactItem resource)
		{
			return resource.WithComCleanup<ContactItem, IContactItem>();
		}

	/// <summary>
		/// Wrapper interface for ItemEvents which adds IDispose to the interface
		/// </summary>
		public static IItemEvents WithComCleanup(this ItemEvents resource)
		{
			return resource.WithComCleanup<ItemEvents, IItemEvents>();
		}

	/// <summary>
		/// Wrapper interface for ItemEvents_10 which adds IDispose to the interface
		/// </summary>
		public static IItemEvents_10 WithComCleanup(this ItemEvents_10 resource)
		{
			return resource.WithComCleanup<ItemEvents_10, IItemEvents_10>();
		}

	/// <summary>
		/// Wrapper interface for _Conversation which adds IDispose to the interface
		/// </summary>
		public static I_Conversation WithComCleanup(this _Conversation resource)
		{
			return resource.WithComCleanup<_Conversation, I_Conversation>();
		}

	/// <summary>
		/// Wrapper interface for SimpleItems which adds IDispose to the interface
		/// </summary>
		public static ISimpleItems WithComCleanup(this SimpleItems resource)
		{
			return resource.WithComCleanup<SimpleItems, ISimpleItems>();
		}

	/// <summary>
		/// Wrapper interface for _SimpleItems which adds IDispose to the interface
		/// </summary>
		public static I_SimpleItems WithComCleanup(this _SimpleItems resource)
		{
			return resource.WithComCleanup<_SimpleItems, I_SimpleItems>();
		}

	/// <summary>
		/// Wrapper interface for UserDefinedProperties which adds IDispose to the interface
		/// </summary>
		public static IUserDefinedProperties WithComCleanup(this UserDefinedProperties resource)
		{
			return resource.WithComCleanup<UserDefinedProperties, IUserDefinedProperties>();
		}

	/// <summary>
		/// Wrapper interface for _UserDefinedProperties which adds IDispose to the interface
		/// </summary>
		public static I_UserDefinedProperties WithComCleanup(this _UserDefinedProperties resource)
		{
			return resource.WithComCleanup<_UserDefinedProperties, I_UserDefinedProperties>();
		}

	/// <summary>
		/// Wrapper interface for _UserDefinedProperty which adds IDispose to the interface
		/// </summary>
		public static I_UserDefinedProperty WithComCleanup(this _UserDefinedProperty resource)
		{
			return resource.WithComCleanup<_UserDefinedProperty, I_UserDefinedProperty>();
		}

	/// <summary>
		/// Wrapper interface for UserDefinedProperty which adds IDispose to the interface
		/// </summary>
		public static IUserDefinedProperty WithComCleanup(this UserDefinedProperty resource)
		{
			return resource.WithComCleanup<UserDefinedProperty, IUserDefinedProperty>();
		}

	/// <summary>
		/// Wrapper interface for ExchangeUser which adds IDispose to the interface
		/// </summary>
		public static IExchangeUser WithComCleanup(this ExchangeUser resource)
		{
			return resource.WithComCleanup<ExchangeUser, IExchangeUser>();
		}

	/// <summary>
		/// Wrapper interface for _ExchangeUser which adds IDispose to the interface
		/// </summary>
		public static I_ExchangeUser WithComCleanup(this _ExchangeUser resource)
		{
			return resource.WithComCleanup<_ExchangeUser, I_ExchangeUser>();
		}

	/// <summary>
		/// Wrapper interface for ExchangeDistributionList which adds IDispose to the interface
		/// </summary>
		public static IExchangeDistributionList WithComCleanup(this ExchangeDistributionList resource)
		{
			return resource.WithComCleanup<ExchangeDistributionList, IExchangeDistributionList>();
		}

	/// <summary>
		/// Wrapper interface for _ExchangeDistributionList which adds IDispose to the interface
		/// </summary>
		public static I_ExchangeDistributionList WithComCleanup(this _ExchangeDistributionList resource)
		{
			return resource.WithComCleanup<_ExchangeDistributionList, I_ExchangeDistributionList>();
		}

	/// <summary>
		/// Wrapper interface for AddressLists which adds IDispose to the interface
		/// </summary>
		public static IAddressLists WithComCleanup(this AddressLists resource)
		{
			return resource.WithComCleanup<AddressLists, IAddressLists>();
		}

	/// <summary>
		/// Wrapper interface for SyncObjects which adds IDispose to the interface
		/// </summary>
		public static ISyncObjects WithComCleanup(this SyncObjects resource)
		{
			return resource.WithComCleanup<SyncObjects, ISyncObjects>();
		}

	/// <summary>
		/// Wrapper interface for SyncObjectEvents_Event which adds IDispose to the interface
		/// </summary>
		public static ISyncObjectEvents_Event WithComCleanup(this SyncObjectEvents_Event resource)
		{
			return resource.WithComCleanup<SyncObjectEvents_Event, ISyncObjectEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for SyncObject which adds IDispose to the interface
		/// </summary>
		public static ISyncObject WithComCleanup(this SyncObject resource)
		{
			return resource.WithComCleanup<SyncObject, ISyncObject>();
		}

	/// <summary>
		/// Wrapper interface for _SyncObject which adds IDispose to the interface
		/// </summary>
		public static I_SyncObject WithComCleanup(this _SyncObject resource)
		{
			return resource.WithComCleanup<_SyncObject, I_SyncObject>();
		}

	/// <summary>
		/// Wrapper interface for SyncObjectEvents which adds IDispose to the interface
		/// </summary>
		public static ISyncObjectEvents WithComCleanup(this SyncObjectEvents resource)
		{
			return resource.WithComCleanup<SyncObjectEvents, ISyncObjectEvents>();
		}

	/// <summary>
		/// Wrapper interface for AccountsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IAccountsEvents_Event WithComCleanup(this AccountsEvents_Event resource)
		{
			return resource.WithComCleanup<AccountsEvents_Event, IAccountsEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for Accounts which adds IDispose to the interface
		/// </summary>
		public static IAccounts WithComCleanup(this Accounts resource)
		{
			return resource.WithComCleanup<Accounts, IAccounts>();
		}

	/// <summary>
		/// Wrapper interface for _Accounts which adds IDispose to the interface
		/// </summary>
		public static I_Accounts WithComCleanup(this _Accounts resource)
		{
			return resource.WithComCleanup<_Accounts, I_Accounts>();
		}

	/// <summary>
		/// Wrapper interface for AccountsEvents which adds IDispose to the interface
		/// </summary>
		public static IAccountsEvents WithComCleanup(this AccountsEvents resource)
		{
			return resource.WithComCleanup<AccountsEvents, IAccountsEvents>();
		}

	/// <summary>
		/// Wrapper interface for StoresEvents_12_Event which adds IDispose to the interface
		/// </summary>
		public static IStoresEvents_12_Event WithComCleanup(this StoresEvents_12_Event resource)
		{
			return resource.WithComCleanup<StoresEvents_12_Event, IStoresEvents_12_Event>();
		}

	/// <summary>
		/// Wrapper interface for Stores which adds IDispose to the interface
		/// </summary>
		public static IStores WithComCleanup(this Stores resource)
		{
			return resource.WithComCleanup<Stores, IStores>();
		}

	/// <summary>
		/// Wrapper interface for _Stores which adds IDispose to the interface
		/// </summary>
		public static I_Stores WithComCleanup(this _Stores resource)
		{
			return resource.WithComCleanup<_Stores, I_Stores>();
		}

	/// <summary>
		/// Wrapper interface for StoresEvents_12 which adds IDispose to the interface
		/// </summary>
		public static IStoresEvents_12 WithComCleanup(this StoresEvents_12 resource)
		{
			return resource.WithComCleanup<StoresEvents_12, IStoresEvents_12>();
		}

	/// <summary>
		/// Wrapper interface for SelectNamesDialog which adds IDispose to the interface
		/// </summary>
		public static ISelectNamesDialog WithComCleanup(this SelectNamesDialog resource)
		{
			return resource.WithComCleanup<SelectNamesDialog, ISelectNamesDialog>();
		}

	/// <summary>
		/// Wrapper interface for _SelectNamesDialog which adds IDispose to the interface
		/// </summary>
		public static I_SelectNamesDialog WithComCleanup(this _SelectNamesDialog resource)
		{
			return resource.WithComCleanup<_SelectNamesDialog, I_SelectNamesDialog>();
		}

	/// <summary>
		/// Wrapper interface for SharingItem which adds IDispose to the interface
		/// </summary>
		public static ISharingItem WithComCleanup(this SharingItem resource)
		{
			return resource.WithComCleanup<SharingItem, ISharingItem>();
		}

	/// <summary>
		/// Wrapper interface for _SharingItem which adds IDispose to the interface
		/// </summary>
		public static I_SharingItem WithComCleanup(this _SharingItem resource)
		{
			return resource.WithComCleanup<_SharingItem, I_SharingItem>();
		}

	/// <summary>
		/// Wrapper interface for _Explorers which adds IDispose to the interface
		/// </summary>
		public static I_Explorers WithComCleanup(this _Explorers resource)
		{
			return resource.WithComCleanup<_Explorers, I_Explorers>();
		}

	/// <summary>
		/// Wrapper interface for ExplorerEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IExplorerEvents_Event WithComCleanup(this ExplorerEvents_Event resource)
		{
			return resource.WithComCleanup<ExplorerEvents_Event, IExplorerEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for ExplorerEvents_10_Event which adds IDispose to the interface
		/// </summary>
		public static IExplorerEvents_10_Event WithComCleanup(this ExplorerEvents_10_Event resource)
		{
			return resource.WithComCleanup<ExplorerEvents_10_Event, IExplorerEvents_10_Event>();
		}

	/// <summary>
		/// Wrapper interface for Explorer which adds IDispose to the interface
		/// </summary>
		public static IExplorer WithComCleanup(this Explorer resource)
		{
			return resource.WithComCleanup<Explorer, IExplorer>();
		}

	/// <summary>
		/// Wrapper interface for ExplorerEvents which adds IDispose to the interface
		/// </summary>
		public static IExplorerEvents WithComCleanup(this ExplorerEvents resource)
		{
			return resource.WithComCleanup<ExplorerEvents, IExplorerEvents>();
		}

	/// <summary>
		/// Wrapper interface for ExplorerEvents_10 which adds IDispose to the interface
		/// </summary>
		public static IExplorerEvents_10 WithComCleanup(this ExplorerEvents_10 resource)
		{
			return resource.WithComCleanup<ExplorerEvents_10, IExplorerEvents_10>();
		}

	/// <summary>
		/// Wrapper interface for _Inspectors which adds IDispose to the interface
		/// </summary>
		public static I_Inspectors WithComCleanup(this _Inspectors resource)
		{
			return resource.WithComCleanup<_Inspectors, I_Inspectors>();
		}

	/// <summary>
		/// Wrapper interface for InspectorEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IInspectorEvents_Event WithComCleanup(this InspectorEvents_Event resource)
		{
			return resource.WithComCleanup<InspectorEvents_Event, IInspectorEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for InspectorEvents_10_Event which adds IDispose to the interface
		/// </summary>
		public static IInspectorEvents_10_Event WithComCleanup(this InspectorEvents_10_Event resource)
		{
			return resource.WithComCleanup<InspectorEvents_10_Event, IInspectorEvents_10_Event>();
		}

	/// <summary>
		/// Wrapper interface for Inspector which adds IDispose to the interface
		/// </summary>
		public static IInspector WithComCleanup(this Inspector resource)
		{
			return resource.WithComCleanup<Inspector, IInspector>();
		}

	/// <summary>
		/// Wrapper interface for InspectorEvents which adds IDispose to the interface
		/// </summary>
		public static IInspectorEvents WithComCleanup(this InspectorEvents resource)
		{
			return resource.WithComCleanup<InspectorEvents, IInspectorEvents>();
		}

	/// <summary>
		/// Wrapper interface for InspectorEvents_10 which adds IDispose to the interface
		/// </summary>
		public static IInspectorEvents_10 WithComCleanup(this InspectorEvents_10 resource)
		{
			return resource.WithComCleanup<InspectorEvents_10, IInspectorEvents_10>();
		}

	/// <summary>
		/// Wrapper interface for Search which adds IDispose to the interface
		/// </summary>
		public static ISearch WithComCleanup(this Search resource)
		{
			return resource.WithComCleanup<Search, ISearch>();
		}

	/// <summary>
		/// Wrapper interface for _Results which adds IDispose to the interface
		/// </summary>
		public static I_Results WithComCleanup(this _Results resource)
		{
			return resource.WithComCleanup<_Results, I_Results>();
		}

	/// <summary>
		/// Wrapper interface for _Reminders which adds IDispose to the interface
		/// </summary>
		public static I_Reminders WithComCleanup(this _Reminders resource)
		{
			return resource.WithComCleanup<_Reminders, I_Reminders>();
		}

	/// <summary>
		/// Wrapper interface for _Reminder which adds IDispose to the interface
		/// </summary>
		public static I_Reminder WithComCleanup(this _Reminder resource)
		{
			return resource.WithComCleanup<_Reminder, I_Reminder>();
		}

	/// <summary>
		/// Wrapper interface for TimeZones which adds IDispose to the interface
		/// </summary>
		public static ITimeZones WithComCleanup(this TimeZones resource)
		{
			return resource.WithComCleanup<TimeZones, ITimeZones>();
		}

	/// <summary>
		/// Wrapper interface for _TimeZones which adds IDispose to the interface
		/// </summary>
		public static I_TimeZones WithComCleanup(this _TimeZones resource)
		{
			return resource.WithComCleanup<_TimeZones, I_TimeZones>();
		}

	/// <summary>
		/// Wrapper interface for _OlkTimeZoneControl which adds IDispose to the interface
		/// </summary>
		public static I_OlkTimeZoneControl WithComCleanup(this _OlkTimeZoneControl resource)
		{
			return resource.WithComCleanup<_OlkTimeZoneControl, I_OlkTimeZoneControl>();
		}

	/// <summary>
		/// Wrapper interface for OlkTimeZoneControlEvents which adds IDispose to the interface
		/// </summary>
		public static IOlkTimeZoneControlEvents WithComCleanup(this OlkTimeZoneControlEvents resource)
		{
			return resource.WithComCleanup<OlkTimeZoneControlEvents, IOlkTimeZoneControlEvents>();
		}

	/// <summary>
		/// Wrapper interface for OlkTimeZoneControlEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IOlkTimeZoneControlEvents_Event WithComCleanup(this OlkTimeZoneControlEvents_Event resource)
		{
			return resource.WithComCleanup<OlkTimeZoneControlEvents_Event, IOlkTimeZoneControlEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for OlkTimeZoneControl which adds IDispose to the interface
		/// </summary>
		public static IOlkTimeZoneControl WithComCleanup(this OlkTimeZoneControl resource)
		{
			return resource.WithComCleanup<OlkTimeZoneControl, IOlkTimeZoneControl>();
		}

	/// <summary>
		/// Wrapper interface for ApplicationEvents which adds IDispose to the interface
		/// </summary>
		public static IApplicationEvents WithComCleanup(this ApplicationEvents resource)
		{
			return resource.WithComCleanup<ApplicationEvents, IApplicationEvents>();
		}

	/// <summary>
		/// Wrapper interface for PropertyPages which adds IDispose to the interface
		/// </summary>
		public static IPropertyPages WithComCleanup(this PropertyPages resource)
		{
			return resource.WithComCleanup<PropertyPages, IPropertyPages>();
		}

	/// <summary>
		/// Wrapper interface for RecurrencePattern which adds IDispose to the interface
		/// </summary>
		public static IRecurrencePattern WithComCleanup(this RecurrencePattern resource)
		{
			return resource.WithComCleanup<RecurrencePattern, IRecurrencePattern>();
		}

	/// <summary>
		/// Wrapper interface for Exceptions which adds IDispose to the interface
		/// </summary>
		public static IExceptions WithComCleanup(this Exceptions resource)
		{
			return resource.WithComCleanup<Exceptions, IExceptions>();
		}

	/// <summary>
		/// Wrapper interface for Exception which adds IDispose to the interface
		/// </summary>
		public static IException WithComCleanup(this Exception resource)
		{
			return resource.WithComCleanup<Exception, IException>();
		}

	/// <summary>
		/// Wrapper interface for AppointmentItem which adds IDispose to the interface
		/// </summary>
		public static IAppointmentItem WithComCleanup(this AppointmentItem resource)
		{
			return resource.WithComCleanup<AppointmentItem, IAppointmentItem>();
		}

	/// <summary>
		/// Wrapper interface for _AppointmentItem which adds IDispose to the interface
		/// </summary>
		public static I_AppointmentItem WithComCleanup(this _AppointmentItem resource)
		{
			return resource.WithComCleanup<_AppointmentItem, I_AppointmentItem>();
		}

	/// <summary>
		/// Wrapper interface for MeetingItem which adds IDispose to the interface
		/// </summary>
		public static IMeetingItem WithComCleanup(this MeetingItem resource)
		{
			return resource.WithComCleanup<MeetingItem, IMeetingItem>();
		}

	/// <summary>
		/// Wrapper interface for _MeetingItem which adds IDispose to the interface
		/// </summary>
		public static I_MeetingItem WithComCleanup(this _MeetingItem resource)
		{
			return resource.WithComCleanup<_MeetingItem, I_MeetingItem>();
		}

	/// <summary>
		/// Wrapper interface for ExplorersEvents which adds IDispose to the interface
		/// </summary>
		public static IExplorersEvents WithComCleanup(this ExplorersEvents resource)
		{
			return resource.WithComCleanup<ExplorersEvents, IExplorersEvents>();
		}

	/// <summary>
		/// Wrapper interface for FoldersEvents which adds IDispose to the interface
		/// </summary>
		public static IFoldersEvents WithComCleanup(this FoldersEvents resource)
		{
			return resource.WithComCleanup<FoldersEvents, IFoldersEvents>();
		}

	/// <summary>
		/// Wrapper interface for InspectorsEvents which adds IDispose to the interface
		/// </summary>
		public static IInspectorsEvents WithComCleanup(this InspectorsEvents resource)
		{
			return resource.WithComCleanup<InspectorsEvents, IInspectorsEvents>();
		}

	/// <summary>
		/// Wrapper interface for ItemsEvents which adds IDispose to the interface
		/// </summary>
		public static IItemsEvents WithComCleanup(this ItemsEvents resource)
		{
			return resource.WithComCleanup<ItemsEvents, IItemsEvents>();
		}

	/// <summary>
		/// Wrapper interface for NameSpaceEvents which adds IDispose to the interface
		/// </summary>
		public static INameSpaceEvents WithComCleanup(this NameSpaceEvents resource)
		{
			return resource.WithComCleanup<NameSpaceEvents, INameSpaceEvents>();
		}

	/// <summary>
		/// Wrapper interface for OutlookBarGroup which adds IDispose to the interface
		/// </summary>
		public static IOutlookBarGroup WithComCleanup(this OutlookBarGroup resource)
		{
			return resource.WithComCleanup<OutlookBarGroup, IOutlookBarGroup>();
		}

	/// <summary>
		/// Wrapper interface for _OutlookBarShortcuts which adds IDispose to the interface
		/// </summary>
		public static I_OutlookBarShortcuts WithComCleanup(this _OutlookBarShortcuts resource)
		{
			return resource.WithComCleanup<_OutlookBarShortcuts, I_OutlookBarShortcuts>();
		}

	/// <summary>
		/// Wrapper interface for OutlookBarShortcut which adds IDispose to the interface
		/// </summary>
		public static IOutlookBarShortcut WithComCleanup(this OutlookBarShortcut resource)
		{
			return resource.WithComCleanup<OutlookBarShortcut, IOutlookBarShortcut>();
		}

	/// <summary>
		/// Wrapper interface for _OutlookBarGroups which adds IDispose to the interface
		/// </summary>
		public static I_OutlookBarGroups WithComCleanup(this _OutlookBarGroups resource)
		{
			return resource.WithComCleanup<_OutlookBarGroups, I_OutlookBarGroups>();
		}

	/// <summary>
		/// Wrapper interface for OutlookBarGroupsEvents which adds IDispose to the interface
		/// </summary>
		public static IOutlookBarGroupsEvents WithComCleanup(this OutlookBarGroupsEvents resource)
		{
			return resource.WithComCleanup<OutlookBarGroupsEvents, IOutlookBarGroupsEvents>();
		}

	/// <summary>
		/// Wrapper interface for _OutlookBarPane which adds IDispose to the interface
		/// </summary>
		public static I_OutlookBarPane WithComCleanup(this _OutlookBarPane resource)
		{
			return resource.WithComCleanup<_OutlookBarPane, I_OutlookBarPane>();
		}

	/// <summary>
		/// Wrapper interface for OutlookBarStorage which adds IDispose to the interface
		/// </summary>
		public static IOutlookBarStorage WithComCleanup(this OutlookBarStorage resource)
		{
			return resource.WithComCleanup<OutlookBarStorage, IOutlookBarStorage>();
		}

	/// <summary>
		/// Wrapper interface for OutlookBarPaneEvents which adds IDispose to the interface
		/// </summary>
		public static IOutlookBarPaneEvents WithComCleanup(this OutlookBarPaneEvents resource)
		{
			return resource.WithComCleanup<OutlookBarPaneEvents, IOutlookBarPaneEvents>();
		}

	/// <summary>
		/// Wrapper interface for OutlookBarShortcutsEvents which adds IDispose to the interface
		/// </summary>
		public static IOutlookBarShortcutsEvents WithComCleanup(this OutlookBarShortcutsEvents resource)
		{
			return resource.WithComCleanup<OutlookBarShortcutsEvents, IOutlookBarShortcutsEvents>();
		}

	/// <summary>
		/// Wrapper interface for PropertyPage which adds IDispose to the interface
		/// </summary>
		public static IPropertyPage WithComCleanup(this PropertyPage resource)
		{
			return resource.WithComCleanup<PropertyPage, IPropertyPage>();
		}

	/// <summary>
		/// Wrapper interface for PropertyPageSite which adds IDispose to the interface
		/// </summary>
		public static IPropertyPageSite WithComCleanup(this PropertyPageSite resource)
		{
			return resource.WithComCleanup<PropertyPageSite, IPropertyPageSite>();
		}

	/// <summary>
		/// Wrapper interface for Pages which adds IDispose to the interface
		/// </summary>
		public static IPages WithComCleanup(this Pages resource)
		{
			return resource.WithComCleanup<Pages, IPages>();
		}

	/// <summary>
		/// Wrapper interface for ApplicationEvents_10 which adds IDispose to the interface
		/// </summary>
		public static IApplicationEvents_10 WithComCleanup(this ApplicationEvents_10 resource)
		{
			return resource.WithComCleanup<ApplicationEvents_10, IApplicationEvents_10>();
		}

	/// <summary>
		/// Wrapper interface for ApplicationEvents_11 which adds IDispose to the interface
		/// </summary>
		public static IApplicationEvents_11 WithComCleanup(this ApplicationEvents_11 resource)
		{
			return resource.WithComCleanup<ApplicationEvents_11, IApplicationEvents_11>();
		}

	/// <summary>
		/// Wrapper interface for AttachmentSelection which adds IDispose to the interface
		/// </summary>
		public static IAttachmentSelection WithComCleanup(this AttachmentSelection resource)
		{
			return resource.WithComCleanup<AttachmentSelection, IAttachmentSelection>();
		}

	/// <summary>
		/// Wrapper interface for MAPIFolderEvents_12_Event which adds IDispose to the interface
		/// </summary>
		public static IMAPIFolderEvents_12_Event WithComCleanup(this MAPIFolderEvents_12_Event resource)
		{
			return resource.WithComCleanup<MAPIFolderEvents_12_Event, IMAPIFolderEvents_12_Event>();
		}

	/// <summary>
		/// Wrapper interface for Folder which adds IDispose to the interface
		/// </summary>
		public static IFolder WithComCleanup(this Folder resource)
		{
			return resource.WithComCleanup<Folder, IFolder>();
		}

	/// <summary>
		/// Wrapper interface for MAPIFolderEvents_12 which adds IDispose to the interface
		/// </summary>
		public static IMAPIFolderEvents_12 WithComCleanup(this MAPIFolderEvents_12 resource)
		{
			return resource.WithComCleanup<MAPIFolderEvents_12, IMAPIFolderEvents_12>();
		}

	/// <summary>
		/// Wrapper interface for ResultsEvents which adds IDispose to the interface
		/// </summary>
		public static IResultsEvents WithComCleanup(this ResultsEvents resource)
		{
			return resource.WithComCleanup<ResultsEvents, IResultsEvents>();
		}

	/// <summary>
		/// Wrapper interface for _ViewsEvents which adds IDispose to the interface
		/// </summary>
		public static I_ViewsEvents WithComCleanup(this _ViewsEvents resource)
		{
			return resource.WithComCleanup<_ViewsEvents, I_ViewsEvents>();
		}

	/// <summary>
		/// Wrapper interface for ReminderCollectionEvents which adds IDispose to the interface
		/// </summary>
		public static IReminderCollectionEvents WithComCleanup(this ReminderCollectionEvents resource)
		{
			return resource.WithComCleanup<ReminderCollectionEvents, IReminderCollectionEvents>();
		}

	/// <summary>
		/// Wrapper interface for _DocumentItem which adds IDispose to the interface
		/// </summary>
		public static I_DocumentItem WithComCleanup(this _DocumentItem resource)
		{
			return resource.WithComCleanup<_DocumentItem, I_DocumentItem>();
		}

	/// <summary>
		/// Wrapper interface for _NoteItem which adds IDispose to the interface
		/// </summary>
		public static I_NoteItem WithComCleanup(this _NoteItem resource)
		{
			return resource.WithComCleanup<_NoteItem, I_NoteItem>();
		}

	/// <summary>
		/// Wrapper interface for FormRegionEvents which adds IDispose to the interface
		/// </summary>
		public static IFormRegionEvents WithComCleanup(this FormRegionEvents resource)
		{
			return resource.WithComCleanup<FormRegionEvents, IFormRegionEvents>();
		}

	/// <summary>
		/// Wrapper interface for _ViewField which adds IDispose to the interface
		/// </summary>
		public static I_ViewField WithComCleanup(this _ViewField resource)
		{
			return resource.WithComCleanup<_ViewField, I_ViewField>();
		}

	/// <summary>
		/// Wrapper interface for ColumnFormat which adds IDispose to the interface
		/// </summary>
		public static IColumnFormat WithComCleanup(this ColumnFormat resource)
		{
			return resource.WithComCleanup<ColumnFormat, IColumnFormat>();
		}

	/// <summary>
		/// Wrapper interface for _ColumnFormat which adds IDispose to the interface
		/// </summary>
		public static I_ColumnFormat WithComCleanup(this _ColumnFormat resource)
		{
			return resource.WithComCleanup<_ColumnFormat, I_ColumnFormat>();
		}

	/// <summary>
		/// Wrapper interface for _ViewFields which adds IDispose to the interface
		/// </summary>
		public static I_ViewFields WithComCleanup(this _ViewFields resource)
		{
			return resource.WithComCleanup<_ViewFields, I_ViewFields>();
		}

	/// <summary>
		/// Wrapper interface for ViewField which adds IDispose to the interface
		/// </summary>
		public static IViewField WithComCleanup(this ViewField resource)
		{
			return resource.WithComCleanup<ViewField, IViewField>();
		}

	/// <summary>
		/// Wrapper interface for _IconView which adds IDispose to the interface
		/// </summary>
		public static I_IconView WithComCleanup(this _IconView resource)
		{
			return resource.WithComCleanup<_IconView, I_IconView>();
		}

	/// <summary>
		/// Wrapper interface for OrderFields which adds IDispose to the interface
		/// </summary>
		public static IOrderFields WithComCleanup(this OrderFields resource)
		{
			return resource.WithComCleanup<OrderFields, IOrderFields>();
		}

	/// <summary>
		/// Wrapper interface for _OrderFields which adds IDispose to the interface
		/// </summary>
		public static I_OrderFields WithComCleanup(this _OrderFields resource)
		{
			return resource.WithComCleanup<_OrderFields, I_OrderFields>();
		}

	/// <summary>
		/// Wrapper interface for _OrderField which adds IDispose to the interface
		/// </summary>
		public static I_OrderField WithComCleanup(this _OrderField resource)
		{
			return resource.WithComCleanup<_OrderField, I_OrderField>();
		}

	/// <summary>
		/// Wrapper interface for OrderField which adds IDispose to the interface
		/// </summary>
		public static IOrderField WithComCleanup(this OrderField resource)
		{
			return resource.WithComCleanup<OrderField, IOrderField>();
		}

	/// <summary>
		/// Wrapper interface for _CardView which adds IDispose to the interface
		/// </summary>
		public static I_CardView WithComCleanup(this _CardView resource)
		{
			return resource.WithComCleanup<_CardView, I_CardView>();
		}

	/// <summary>
		/// Wrapper interface for ViewFields which adds IDispose to the interface
		/// </summary>
		public static IViewFields WithComCleanup(this ViewFields resource)
		{
			return resource.WithComCleanup<ViewFields, IViewFields>();
		}

	/// <summary>
		/// Wrapper interface for ViewFont which adds IDispose to the interface
		/// </summary>
		public static IViewFont WithComCleanup(this ViewFont resource)
		{
			return resource.WithComCleanup<ViewFont, IViewFont>();
		}

	/// <summary>
		/// Wrapper interface for _ViewFont which adds IDispose to the interface
		/// </summary>
		public static I_ViewFont WithComCleanup(this _ViewFont resource)
		{
			return resource.WithComCleanup<_ViewFont, I_ViewFont>();
		}

	/// <summary>
		/// Wrapper interface for AutoFormatRules which adds IDispose to the interface
		/// </summary>
		public static IAutoFormatRules WithComCleanup(this AutoFormatRules resource)
		{
			return resource.WithComCleanup<AutoFormatRules, IAutoFormatRules>();
		}

	/// <summary>
		/// Wrapper interface for _AutoFormatRules which adds IDispose to the interface
		/// </summary>
		public static I_AutoFormatRules WithComCleanup(this _AutoFormatRules resource)
		{
			return resource.WithComCleanup<_AutoFormatRules, I_AutoFormatRules>();
		}

	/// <summary>
		/// Wrapper interface for AutoFormatRule which adds IDispose to the interface
		/// </summary>
		public static IAutoFormatRule WithComCleanup(this AutoFormatRule resource)
		{
			return resource.WithComCleanup<AutoFormatRule, IAutoFormatRule>();
		}

	/// <summary>
		/// Wrapper interface for _AutoFormatRule which adds IDispose to the interface
		/// </summary>
		public static I_AutoFormatRule WithComCleanup(this _AutoFormatRule resource)
		{
			return resource.WithComCleanup<_AutoFormatRule, I_AutoFormatRule>();
		}

	/// <summary>
		/// Wrapper interface for _TimelineView which adds IDispose to the interface
		/// </summary>
		public static I_TimelineView WithComCleanup(this _TimelineView resource)
		{
			return resource.WithComCleanup<_TimelineView, I_TimelineView>();
		}

	/// <summary>
		/// Wrapper interface for _MailModule which adds IDispose to the interface
		/// </summary>
		public static I_MailModule WithComCleanup(this _MailModule resource)
		{
			return resource.WithComCleanup<_MailModule, I_MailModule>();
		}

	/// <summary>
		/// Wrapper interface for _NavigationGroups which adds IDispose to the interface
		/// </summary>
		public static I_NavigationGroups WithComCleanup(this _NavigationGroups resource)
		{
			return resource.WithComCleanup<_NavigationGroups, I_NavigationGroups>();
		}

	/// <summary>
		/// Wrapper interface for _NavigationGroup which adds IDispose to the interface
		/// </summary>
		public static I_NavigationGroup WithComCleanup(this _NavigationGroup resource)
		{
			return resource.WithComCleanup<_NavigationGroup, I_NavigationGroup>();
		}

	/// <summary>
		/// Wrapper interface for NavigationFolders which adds IDispose to the interface
		/// </summary>
		public static INavigationFolders WithComCleanup(this NavigationFolders resource)
		{
			return resource.WithComCleanup<NavigationFolders, INavigationFolders>();
		}

	/// <summary>
		/// Wrapper interface for _NavigationFolders which adds IDispose to the interface
		/// </summary>
		public static I_NavigationFolders WithComCleanup(this _NavigationFolders resource)
		{
			return resource.WithComCleanup<_NavigationFolders, I_NavigationFolders>();
		}

	/// <summary>
		/// Wrapper interface for _NavigationFolder which adds IDispose to the interface
		/// </summary>
		public static I_NavigationFolder WithComCleanup(this _NavigationFolder resource)
		{
			return resource.WithComCleanup<_NavigationFolder, I_NavigationFolder>();
		}

	/// <summary>
		/// Wrapper interface for NavigationFolder which adds IDispose to the interface
		/// </summary>
		public static INavigationFolder WithComCleanup(this NavigationFolder resource)
		{
			return resource.WithComCleanup<NavigationFolder, INavigationFolder>();
		}

	/// <summary>
		/// Wrapper interface for NavigationGroup which adds IDispose to the interface
		/// </summary>
		public static INavigationGroup WithComCleanup(this NavigationGroup resource)
		{
			return resource.WithComCleanup<NavigationGroup, INavigationGroup>();
		}

	/// <summary>
		/// Wrapper interface for _CalendarModule which adds IDispose to the interface
		/// </summary>
		public static I_CalendarModule WithComCleanup(this _CalendarModule resource)
		{
			return resource.WithComCleanup<_CalendarModule, I_CalendarModule>();
		}

	/// <summary>
		/// Wrapper interface for _ContactsModule which adds IDispose to the interface
		/// </summary>
		public static I_ContactsModule WithComCleanup(this _ContactsModule resource)
		{
			return resource.WithComCleanup<_ContactsModule, I_ContactsModule>();
		}

	/// <summary>
		/// Wrapper interface for _TasksModule which adds IDispose to the interface
		/// </summary>
		public static I_TasksModule WithComCleanup(this _TasksModule resource)
		{
			return resource.WithComCleanup<_TasksModule, I_TasksModule>();
		}

	/// <summary>
		/// Wrapper interface for _JournalModule which adds IDispose to the interface
		/// </summary>
		public static I_JournalModule WithComCleanup(this _JournalModule resource)
		{
			return resource.WithComCleanup<_JournalModule, I_JournalModule>();
		}

	/// <summary>
		/// Wrapper interface for _NotesModule which adds IDispose to the interface
		/// </summary>
		public static I_NotesModule WithComCleanup(this _NotesModule resource)
		{
			return resource.WithComCleanup<_NotesModule, I_NotesModule>();
		}

	/// <summary>
		/// Wrapper interface for NavigationPaneEvents_12 which adds IDispose to the interface
		/// </summary>
		public static INavigationPaneEvents_12 WithComCleanup(this NavigationPaneEvents_12 resource)
		{
			return resource.WithComCleanup<NavigationPaneEvents_12, INavigationPaneEvents_12>();
		}

	/// <summary>
		/// Wrapper interface for NavigationGroupsEvents_12 which adds IDispose to the interface
		/// </summary>
		public static INavigationGroupsEvents_12 WithComCleanup(this NavigationGroupsEvents_12 resource)
		{
			return resource.WithComCleanup<NavigationGroupsEvents_12, INavigationGroupsEvents_12>();
		}

	/// <summary>
		/// Wrapper interface for _BusinessCardView which adds IDispose to the interface
		/// </summary>
		public static I_BusinessCardView WithComCleanup(this _BusinessCardView resource)
		{
			return resource.WithComCleanup<_BusinessCardView, I_BusinessCardView>();
		}

	/// <summary>
		/// Wrapper interface for _FormRegionStartup which adds IDispose to the interface
		/// </summary>
		public static I_FormRegionStartup WithComCleanup(this _FormRegionStartup resource)
		{
			return resource.WithComCleanup<_FormRegionStartup, I_FormRegionStartup>();
		}

	/// <summary>
		/// Wrapper interface for FormRegionEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IFormRegionEvents_Event WithComCleanup(this FormRegionEvents_Event resource)
		{
			return resource.WithComCleanup<FormRegionEvents_Event, IFormRegionEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for FormRegion which adds IDispose to the interface
		/// </summary>
		public static IFormRegion WithComCleanup(this FormRegion resource)
		{
			return resource.WithComCleanup<FormRegion, IFormRegion>();
		}

	/// <summary>
		/// Wrapper interface for _FormRegion which adds IDispose to the interface
		/// </summary>
		public static I_FormRegion WithComCleanup(this _FormRegion resource)
		{
			return resource.WithComCleanup<_FormRegion, I_FormRegion>();
		}

	/// <summary>
		/// Wrapper interface for _SolutionsModule which adds IDispose to the interface
		/// </summary>
		public static I_SolutionsModule WithComCleanup(this _SolutionsModule resource)
		{
			return resource.WithComCleanup<_SolutionsModule, I_SolutionsModule>();
		}

	/// <summary>
		/// Wrapper interface for _CalendarView which adds IDispose to the interface
		/// </summary>
		public static I_CalendarView WithComCleanup(this _CalendarView resource)
		{
			return resource.WithComCleanup<_CalendarView, I_CalendarView>();
		}

	/// <summary>
		/// Wrapper interface for _TableView which adds IDispose to the interface
		/// </summary>
		public static I_TableView WithComCleanup(this _TableView resource)
		{
			return resource.WithComCleanup<_TableView, I_TableView>();
		}

	/// <summary>
		/// Wrapper interface for _MobileItem which adds IDispose to the interface
		/// </summary>
		public static I_MobileItem WithComCleanup(this _MobileItem resource)
		{
			return resource.WithComCleanup<_MobileItem, I_MobileItem>();
		}

	/// <summary>
		/// Wrapper interface for MobileItem which adds IDispose to the interface
		/// </summary>
		public static IMobileItem WithComCleanup(this MobileItem resource)
		{
			return resource.WithComCleanup<MobileItem, IMobileItem>();
		}

	/// <summary>
		/// Wrapper interface for _JournalItem which adds IDispose to the interface
		/// </summary>
		public static I_JournalItem WithComCleanup(this _JournalItem resource)
		{
			return resource.WithComCleanup<_JournalItem, I_JournalItem>();
		}

	/// <summary>
		/// Wrapper interface for _PostItem which adds IDispose to the interface
		/// </summary>
		public static I_PostItem WithComCleanup(this _PostItem resource)
		{
			return resource.WithComCleanup<_PostItem, I_PostItem>();
		}

	/// <summary>
		/// Wrapper interface for _TaskItem which adds IDispose to the interface
		/// </summary>
		public static I_TaskItem WithComCleanup(this _TaskItem resource)
		{
			return resource.WithComCleanup<_TaskItem, I_TaskItem>();
		}

	/// <summary>
		/// Wrapper interface for TaskItem which adds IDispose to the interface
		/// </summary>
		public static ITaskItem WithComCleanup(this TaskItem resource)
		{
			return resource.WithComCleanup<TaskItem, ITaskItem>();
		}

	/// <summary>
		/// Wrapper interface for AccountSelectorEvents which adds IDispose to the interface
		/// </summary>
		public static IAccountSelectorEvents WithComCleanup(this AccountSelectorEvents resource)
		{
			return resource.WithComCleanup<AccountSelectorEvents, IAccountSelectorEvents>();
		}

	/// <summary>
		/// Wrapper interface for _DistListItem which adds IDispose to the interface
		/// </summary>
		public static I_DistListItem WithComCleanup(this _DistListItem resource)
		{
			return resource.WithComCleanup<_DistListItem, I_DistListItem>();
		}

	/// <summary>
		/// Wrapper interface for _ReportItem which adds IDispose to the interface
		/// </summary>
		public static I_ReportItem WithComCleanup(this _ReportItem resource)
		{
			return resource.WithComCleanup<_ReportItem, I_ReportItem>();
		}

	/// <summary>
		/// Wrapper interface for _RemoteItem which adds IDispose to the interface
		/// </summary>
		public static I_RemoteItem WithComCleanup(this _RemoteItem resource)
		{
			return resource.WithComCleanup<_RemoteItem, I_RemoteItem>();
		}

	/// <summary>
		/// Wrapper interface for _TaskRequestItem which adds IDispose to the interface
		/// </summary>
		public static I_TaskRequestItem WithComCleanup(this _TaskRequestItem resource)
		{
			return resource.WithComCleanup<_TaskRequestItem, I_TaskRequestItem>();
		}

	/// <summary>
		/// Wrapper interface for _TaskRequestAcceptItem which adds IDispose to the interface
		/// </summary>
		public static I_TaskRequestAcceptItem WithComCleanup(this _TaskRequestAcceptItem resource)
		{
			return resource.WithComCleanup<_TaskRequestAcceptItem, I_TaskRequestAcceptItem>();
		}

	/// <summary>
		/// Wrapper interface for _TaskRequestDeclineItem which adds IDispose to the interface
		/// </summary>
		public static I_TaskRequestDeclineItem WithComCleanup(this _TaskRequestDeclineItem resource)
		{
			return resource.WithComCleanup<_TaskRequestDeclineItem, I_TaskRequestDeclineItem>();
		}

	/// <summary>
		/// Wrapper interface for _TaskRequestUpdateItem which adds IDispose to the interface
		/// </summary>
		public static I_TaskRequestUpdateItem WithComCleanup(this _TaskRequestUpdateItem resource)
		{
			return resource.WithComCleanup<_TaskRequestUpdateItem, I_TaskRequestUpdateItem>();
		}

	/// <summary>
		/// Wrapper interface for _ConversationHeader which adds IDispose to the interface
		/// </summary>
		public static I_ConversationHeader WithComCleanup(this _ConversationHeader resource)
		{
			return resource.WithComCleanup<_ConversationHeader, I_ConversationHeader>();
		}

	/// <summary>
		/// Wrapper interface for ApplicationEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IApplicationEvents_Event WithComCleanup(this ApplicationEvents_Event resource)
		{
			return resource.WithComCleanup<ApplicationEvents_Event, IApplicationEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for ApplicationEvents_10_Event which adds IDispose to the interface
		/// </summary>
		public static IApplicationEvents_10_Event WithComCleanup(this ApplicationEvents_10_Event resource)
		{
			return resource.WithComCleanup<ApplicationEvents_10_Event, IApplicationEvents_10_Event>();
		}

	/// <summary>
		/// Wrapper interface for ApplicationEvents_11_Event which adds IDispose to the interface
		/// </summary>
		public static IApplicationEvents_11_Event WithComCleanup(this ApplicationEvents_11_Event resource)
		{
			return resource.WithComCleanup<ApplicationEvents_11_Event, IApplicationEvents_11_Event>();
		}

	/// <summary>
		/// Wrapper interface for Application which adds IDispose to the interface
		/// </summary>
		public static IApplication WithComCleanup(this Application resource)
		{
			return resource.WithComCleanup<Application, IApplication>();
		}

	/// <summary>
		/// Wrapper interface for DistListItem which adds IDispose to the interface
		/// </summary>
		public static IDistListItem WithComCleanup(this DistListItem resource)
		{
			return resource.WithComCleanup<DistListItem, IDistListItem>();
		}

	/// <summary>
		/// Wrapper interface for DocumentItem which adds IDispose to the interface
		/// </summary>
		public static IDocumentItem WithComCleanup(this DocumentItem resource)
		{
			return resource.WithComCleanup<DocumentItem, IDocumentItem>();
		}

	/// <summary>
		/// Wrapper interface for ExplorersEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IExplorersEvents_Event WithComCleanup(this ExplorersEvents_Event resource)
		{
			return resource.WithComCleanup<ExplorersEvents_Event, IExplorersEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for Explorers which adds IDispose to the interface
		/// </summary>
		public static IExplorers WithComCleanup(this Explorers resource)
		{
			return resource.WithComCleanup<Explorers, IExplorers>();
		}

	/// <summary>
		/// Wrapper interface for InspectorsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IInspectorsEvents_Event WithComCleanup(this InspectorsEvents_Event resource)
		{
			return resource.WithComCleanup<InspectorsEvents_Event, IInspectorsEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for Inspectors which adds IDispose to the interface
		/// </summary>
		public static IInspectors WithComCleanup(this Inspectors resource)
		{
			return resource.WithComCleanup<Inspectors, IInspectors>();
		}

	/// <summary>
		/// Wrapper interface for FoldersEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IFoldersEvents_Event WithComCleanup(this FoldersEvents_Event resource)
		{
			return resource.WithComCleanup<FoldersEvents_Event, IFoldersEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for Folders which adds IDispose to the interface
		/// </summary>
		public static IFolders WithComCleanup(this Folders resource)
		{
			return resource.WithComCleanup<Folders, IFolders>();
		}

	/// <summary>
		/// Wrapper interface for ItemsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IItemsEvents_Event WithComCleanup(this ItemsEvents_Event resource)
		{
			return resource.WithComCleanup<ItemsEvents_Event, IItemsEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for Items which adds IDispose to the interface
		/// </summary>
		public static IItems WithComCleanup(this Items resource)
		{
			return resource.WithComCleanup<Items, IItems>();
		}

	/// <summary>
		/// Wrapper interface for JournalItem which adds IDispose to the interface
		/// </summary>
		public static IJournalItem WithComCleanup(this JournalItem resource)
		{
			return resource.WithComCleanup<JournalItem, IJournalItem>();
		}

	/// <summary>
		/// Wrapper interface for NameSpaceEvents_Event which adds IDispose to the interface
		/// </summary>
		public static INameSpaceEvents_Event WithComCleanup(this NameSpaceEvents_Event resource)
		{
			return resource.WithComCleanup<NameSpaceEvents_Event, INameSpaceEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for NameSpace which adds IDispose to the interface
		/// </summary>
		public static INameSpace WithComCleanup(this NameSpace resource)
		{
			return resource.WithComCleanup<NameSpace, INameSpace>();
		}

	/// <summary>
		/// Wrapper interface for NoteItem which adds IDispose to the interface
		/// </summary>
		public static INoteItem WithComCleanup(this NoteItem resource)
		{
			return resource.WithComCleanup<NoteItem, INoteItem>();
		}

	/// <summary>
		/// Wrapper interface for OutlookBarGroupsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IOutlookBarGroupsEvents_Event WithComCleanup(this OutlookBarGroupsEvents_Event resource)
		{
			return resource.WithComCleanup<OutlookBarGroupsEvents_Event, IOutlookBarGroupsEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for OutlookBarGroups which adds IDispose to the interface
		/// </summary>
		public static IOutlookBarGroups WithComCleanup(this OutlookBarGroups resource)
		{
			return resource.WithComCleanup<OutlookBarGroups, IOutlookBarGroups>();
		}

	/// <summary>
		/// Wrapper interface for OutlookBarPaneEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IOutlookBarPaneEvents_Event WithComCleanup(this OutlookBarPaneEvents_Event resource)
		{
			return resource.WithComCleanup<OutlookBarPaneEvents_Event, IOutlookBarPaneEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for OutlookBarPane which adds IDispose to the interface
		/// </summary>
		public static IOutlookBarPane WithComCleanup(this OutlookBarPane resource)
		{
			return resource.WithComCleanup<OutlookBarPane, IOutlookBarPane>();
		}

	/// <summary>
		/// Wrapper interface for OutlookBarShortcutsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IOutlookBarShortcutsEvents_Event WithComCleanup(this OutlookBarShortcutsEvents_Event resource)
		{
			return resource.WithComCleanup<OutlookBarShortcutsEvents_Event, IOutlookBarShortcutsEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for OutlookBarShortcuts which adds IDispose to the interface
		/// </summary>
		public static IOutlookBarShortcuts WithComCleanup(this OutlookBarShortcuts resource)
		{
			return resource.WithComCleanup<OutlookBarShortcuts, IOutlookBarShortcuts>();
		}

	/// <summary>
		/// Wrapper interface for PostItem which adds IDispose to the interface
		/// </summary>
		public static IPostItem WithComCleanup(this PostItem resource)
		{
			return resource.WithComCleanup<PostItem, IPostItem>();
		}

	/// <summary>
		/// Wrapper interface for RemoteItem which adds IDispose to the interface
		/// </summary>
		public static IRemoteItem WithComCleanup(this RemoteItem resource)
		{
			return resource.WithComCleanup<RemoteItem, IRemoteItem>();
		}

	/// <summary>
		/// Wrapper interface for ReportItem which adds IDispose to the interface
		/// </summary>
		public static IReportItem WithComCleanup(this ReportItem resource)
		{
			return resource.WithComCleanup<ReportItem, IReportItem>();
		}

	/// <summary>
		/// Wrapper interface for TaskRequestAcceptItem which adds IDispose to the interface
		/// </summary>
		public static ITaskRequestAcceptItem WithComCleanup(this TaskRequestAcceptItem resource)
		{
			return resource.WithComCleanup<TaskRequestAcceptItem, ITaskRequestAcceptItem>();
		}

	/// <summary>
		/// Wrapper interface for TaskRequestDeclineItem which adds IDispose to the interface
		/// </summary>
		public static ITaskRequestDeclineItem WithComCleanup(this TaskRequestDeclineItem resource)
		{
			return resource.WithComCleanup<TaskRequestDeclineItem, ITaskRequestDeclineItem>();
		}

	/// <summary>
		/// Wrapper interface for TaskRequestItem which adds IDispose to the interface
		/// </summary>
		public static ITaskRequestItem WithComCleanup(this TaskRequestItem resource)
		{
			return resource.WithComCleanup<TaskRequestItem, ITaskRequestItem>();
		}

	/// <summary>
		/// Wrapper interface for TaskRequestUpdateItem which adds IDispose to the interface
		/// </summary>
		public static ITaskRequestUpdateItem WithComCleanup(this TaskRequestUpdateItem resource)
		{
			return resource.WithComCleanup<TaskRequestUpdateItem, ITaskRequestUpdateItem>();
		}

	/// <summary>
		/// Wrapper interface for ResultsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IResultsEvents_Event WithComCleanup(this ResultsEvents_Event resource)
		{
			return resource.WithComCleanup<ResultsEvents_Event, IResultsEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for Results which adds IDispose to the interface
		/// </summary>
		public static IResults WithComCleanup(this Results resource)
		{
			return resource.WithComCleanup<Results, IResults>();
		}

	/// <summary>
		/// Wrapper interface for _ViewsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static I_ViewsEvents_Event WithComCleanup(this _ViewsEvents_Event resource)
		{
			return resource.WithComCleanup<_ViewsEvents_Event, I_ViewsEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for Views which adds IDispose to the interface
		/// </summary>
		public static IViews WithComCleanup(this Views resource)
		{
			return resource.WithComCleanup<Views, IViews>();
		}

	/// <summary>
		/// Wrapper interface for Reminder which adds IDispose to the interface
		/// </summary>
		public static IReminder WithComCleanup(this Reminder resource)
		{
			return resource.WithComCleanup<Reminder, IReminder>();
		}

	/// <summary>
		/// Wrapper interface for ReminderCollectionEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IReminderCollectionEvents_Event WithComCleanup(this ReminderCollectionEvents_Event resource)
		{
			return resource.WithComCleanup<ReminderCollectionEvents_Event, IReminderCollectionEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for Reminders which adds IDispose to the interface
		/// </summary>
		public static IReminders WithComCleanup(this Reminders resource)
		{
			return resource.WithComCleanup<Reminders, IReminders>();
		}

	/// <summary>
		/// Wrapper interface for StorageItem which adds IDispose to the interface
		/// </summary>
		public static IStorageItem WithComCleanup(this StorageItem resource)
		{
			return resource.WithComCleanup<StorageItem, IStorageItem>();
		}

	/// <summary>
		/// Wrapper interface for NavigationPaneEvents_12_Event which adds IDispose to the interface
		/// </summary>
		public static INavigationPaneEvents_12_Event WithComCleanup(this NavigationPaneEvents_12_Event resource)
		{
			return resource.WithComCleanup<NavigationPaneEvents_12_Event, INavigationPaneEvents_12_Event>();
		}

	/// <summary>
		/// Wrapper interface for NavigationPane which adds IDispose to the interface
		/// </summary>
		public static INavigationPane WithComCleanup(this NavigationPane resource)
		{
			return resource.WithComCleanup<NavigationPane, INavigationPane>();
		}

	/// <summary>
		/// Wrapper interface for NavigationGroupsEvents_12_Event which adds IDispose to the interface
		/// </summary>
		public static INavigationGroupsEvents_12_Event WithComCleanup(this NavigationGroupsEvents_12_Event resource)
		{
			return resource.WithComCleanup<NavigationGroupsEvents_12_Event, INavigationGroupsEvents_12_Event>();
		}

	/// <summary>
		/// Wrapper interface for NavigationGroups which adds IDispose to the interface
		/// </summary>
		public static INavigationGroups WithComCleanup(this NavigationGroups resource)
		{
			return resource.WithComCleanup<NavigationGroups, INavigationGroups>();
		}

	/// <summary>
		/// Wrapper interface for DoNotUseMeFolder which adds IDispose to the interface
		/// </summary>
		public static IDoNotUseMeFolder WithComCleanup(this DoNotUseMeFolder resource)
		{
			return resource.WithComCleanup<DoNotUseMeFolder, IDoNotUseMeFolder>();
		}

	/// <summary>
		/// Wrapper interface for TimelineView which adds IDispose to the interface
		/// </summary>
		public static ITimelineView WithComCleanup(this TimelineView resource)
		{
			return resource.WithComCleanup<TimelineView, ITimelineView>();
		}

	/// <summary>
		/// Wrapper interface for MailModule which adds IDispose to the interface
		/// </summary>
		public static IMailModule WithComCleanup(this MailModule resource)
		{
			return resource.WithComCleanup<MailModule, IMailModule>();
		}

	/// <summary>
		/// Wrapper interface for CalendarModule which adds IDispose to the interface
		/// </summary>
		public static ICalendarModule WithComCleanup(this CalendarModule resource)
		{
			return resource.WithComCleanup<CalendarModule, ICalendarModule>();
		}

	/// <summary>
		/// Wrapper interface for ContactsModule which adds IDispose to the interface
		/// </summary>
		public static IContactsModule WithComCleanup(this ContactsModule resource)
		{
			return resource.WithComCleanup<ContactsModule, IContactsModule>();
		}

	/// <summary>
		/// Wrapper interface for TasksModule which adds IDispose to the interface
		/// </summary>
		public static ITasksModule WithComCleanup(this TasksModule resource)
		{
			return resource.WithComCleanup<TasksModule, ITasksModule>();
		}

	/// <summary>
		/// Wrapper interface for JournalModule which adds IDispose to the interface
		/// </summary>
		public static IJournalModule WithComCleanup(this JournalModule resource)
		{
			return resource.WithComCleanup<JournalModule, IJournalModule>();
		}

	/// <summary>
		/// Wrapper interface for NotesModule which adds IDispose to the interface
		/// </summary>
		public static INotesModule WithComCleanup(this NotesModule resource)
		{
			return resource.WithComCleanup<NotesModule, INotesModule>();
		}

	/// <summary>
		/// Wrapper interface for TableView which adds IDispose to the interface
		/// </summary>
		public static ITableView WithComCleanup(this TableView resource)
		{
			return resource.WithComCleanup<TableView, ITableView>();
		}

	/// <summary>
		/// Wrapper interface for IconView which adds IDispose to the interface
		/// </summary>
		public static IIconView WithComCleanup(this IconView resource)
		{
			return resource.WithComCleanup<IconView, IIconView>();
		}

	/// <summary>
		/// Wrapper interface for CardView which adds IDispose to the interface
		/// </summary>
		public static ICardView WithComCleanup(this CardView resource)
		{
			return resource.WithComCleanup<CardView, ICardView>();
		}

	/// <summary>
		/// Wrapper interface for CalendarView which adds IDispose to the interface
		/// </summary>
		public static ICalendarView WithComCleanup(this CalendarView resource)
		{
			return resource.WithComCleanup<CalendarView, ICalendarView>();
		}

	/// <summary>
		/// Wrapper interface for BusinessCardView which adds IDispose to the interface
		/// </summary>
		public static IBusinessCardView WithComCleanup(this BusinessCardView resource)
		{
			return resource.WithComCleanup<BusinessCardView, IBusinessCardView>();
		}

	/// <summary>
		/// Wrapper interface for FormRegionStartup which adds IDispose to the interface
		/// </summary>
		public static IFormRegionStartup WithComCleanup(this FormRegionStartup resource)
		{
			return resource.WithComCleanup<FormRegionStartup, IFormRegionStartup>();
		}

	/// <summary>
		/// Wrapper interface for TimeZone which adds IDispose to the interface
		/// </summary>
		public static ITimeZone WithComCleanup(this TimeZone resource)
		{
			return resource.WithComCleanup<TimeZone, ITimeZone>();
		}

	/// <summary>
		/// Wrapper interface for SolutionsModule which adds IDispose to the interface
		/// </summary>
		public static ISolutionsModule WithComCleanup(this SolutionsModule resource)
		{
			return resource.WithComCleanup<SolutionsModule, ISolutionsModule>();
		}

	/// <summary>
		/// Wrapper interface for Conversation which adds IDispose to the interface
		/// </summary>
		public static IConversation WithComCleanup(this Conversation resource)
		{
			return resource.WithComCleanup<Conversation, IConversation>();
		}

	/// <summary>
		/// Wrapper interface for AccountSelectorEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IAccountSelectorEvents_Event WithComCleanup(this AccountSelectorEvents_Event resource)
		{
			return resource.WithComCleanup<AccountSelectorEvents_Event, IAccountSelectorEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for AccountSelector which adds IDispose to the interface
		/// </summary>
		public static IAccountSelector WithComCleanup(this AccountSelector resource)
		{
			return resource.WithComCleanup<AccountSelector, IAccountSelector>();
		}

	/// <summary>
		/// Wrapper interface for ConversationHeader which adds IDispose to the interface
		/// </summary>
		public static IConversationHeader WithComCleanup(this ConversationHeader resource)
		{
			return resource.WithComCleanup<ConversationHeader, IConversationHeader>();
		}

	}
}