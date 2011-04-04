//Microsoft.Office.Interop.Outlook, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c
namespace Office.Contrib.Extensions
{
	/// <summary>
	/// Provides cleanup extension methods for interfaces exposed by Microsoft.Office.Interop.Outlook.dll
	/// </summary>
	public static class OutlookCleanupExtensions
	{
		/// <summary>
		/// Wrapper interface for _IRecipientControl which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_IRecipientControl WithComCleanup(this Microsoft.Office.Interop.Outlook._IRecipientControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._IRecipientControl, Outlook.Contrib.Interfaces.I_IRecipientControl>();
		}

		/// <summary>
		/// Wrapper interface for _DRecipientControl which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_DRecipientControl WithComCleanup(this Microsoft.Office.Interop.Outlook._DRecipientControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._DRecipientControl, Outlook.Contrib.Interfaces.I_DRecipientControl>();
		}

		/// <summary>
		/// Wrapper interface for _DRecipientControlEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_DRecipientControlEvents WithComCleanup(this Microsoft.Office.Interop.Outlook._DRecipientControlEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._DRecipientControlEvents, Outlook.Contrib.Interfaces.I_DRecipientControlEvents>();
		}

		/// <summary>
		/// Wrapper interface for _DRecipientControlEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_DRecipientControlEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook._DRecipientControlEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._DRecipientControlEvents_Event, Outlook.Contrib.Interfaces.I_DRecipientControlEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for _RecipientControl which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_RecipientControl WithComCleanup(this Microsoft.Office.Interop.Outlook._RecipientControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._RecipientControl, Outlook.Contrib.Interfaces.I_RecipientControl>();
		}

		/// <summary>
		/// Wrapper interface for _IDocSiteControl which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_IDocSiteControl WithComCleanup(this Microsoft.Office.Interop.Outlook._IDocSiteControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._IDocSiteControl, Outlook.Contrib.Interfaces.I_IDocSiteControl>();
		}

		/// <summary>
		/// Wrapper interface for _DDocSiteControl which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_DDocSiteControl WithComCleanup(this Microsoft.Office.Interop.Outlook._DDocSiteControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._DDocSiteControl, Outlook.Contrib.Interfaces.I_DDocSiteControl>();
		}

		/// <summary>
		/// Wrapper interface for _DDocSiteControlEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_DDocSiteControlEvents WithComCleanup(this Microsoft.Office.Interop.Outlook._DDocSiteControlEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._DDocSiteControlEvents, Outlook.Contrib.Interfaces.I_DDocSiteControlEvents>();
		}

		/// <summary>
		/// Wrapper interface for _DDocSiteControlEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_DDocSiteControlEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook._DDocSiteControlEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._DDocSiteControlEvents_Event, Outlook.Contrib.Interfaces.I_DDocSiteControlEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for _DocSiteControl which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_DocSiteControl WithComCleanup(this Microsoft.Office.Interop.Outlook._DocSiteControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._DocSiteControl, Outlook.Contrib.Interfaces.I_DocSiteControl>();
		}

		/// <summary>
		/// Wrapper interface for OlkControl which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkControl WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkControl, Outlook.Contrib.Interfaces.IOlkControl>();
		}

		/// <summary>
		/// Wrapper interface for _OlkTextBox which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OlkTextBox WithComCleanup(this Microsoft.Office.Interop.Outlook._OlkTextBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OlkTextBox, Outlook.Contrib.Interfaces.I_OlkTextBox>();
		}

		/// <summary>
		/// Wrapper interface for OlkTextBoxEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkTextBoxEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkTextBoxEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkTextBoxEvents, Outlook.Contrib.Interfaces.IOlkTextBoxEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkTextBoxEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkTextBoxEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkTextBoxEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkTextBoxEvents_Event, Outlook.Contrib.Interfaces.IOlkTextBoxEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkTextBox which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkTextBox WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkTextBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkTextBox, Outlook.Contrib.Interfaces.IOlkTextBox>();
		}

		/// <summary>
		/// Wrapper interface for _OlkLabel which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OlkLabel WithComCleanup(this Microsoft.Office.Interop.Outlook._OlkLabel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OlkLabel, Outlook.Contrib.Interfaces.I_OlkLabel>();
		}

		/// <summary>
		/// Wrapper interface for OlkLabelEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkLabelEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkLabelEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkLabelEvents, Outlook.Contrib.Interfaces.IOlkLabelEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkLabelEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkLabelEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkLabelEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkLabelEvents_Event, Outlook.Contrib.Interfaces.IOlkLabelEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkLabel which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkLabel WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkLabel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkLabel, Outlook.Contrib.Interfaces.IOlkLabel>();
		}

		/// <summary>
		/// Wrapper interface for _OlkCommandButton which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OlkCommandButton WithComCleanup(this Microsoft.Office.Interop.Outlook._OlkCommandButton resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OlkCommandButton, Outlook.Contrib.Interfaces.I_OlkCommandButton>();
		}

		/// <summary>
		/// Wrapper interface for OlkCommandButtonEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkCommandButtonEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkCommandButtonEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkCommandButtonEvents, Outlook.Contrib.Interfaces.IOlkCommandButtonEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkCommandButtonEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkCommandButtonEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkCommandButtonEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkCommandButtonEvents_Event, Outlook.Contrib.Interfaces.IOlkCommandButtonEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkCommandButton which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkCommandButton WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkCommandButton resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkCommandButton, Outlook.Contrib.Interfaces.IOlkCommandButton>();
		}

		/// <summary>
		/// Wrapper interface for _OlkCheckBox which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OlkCheckBox WithComCleanup(this Microsoft.Office.Interop.Outlook._OlkCheckBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OlkCheckBox, Outlook.Contrib.Interfaces.I_OlkCheckBox>();
		}

		/// <summary>
		/// Wrapper interface for OlkCheckBoxEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkCheckBoxEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkCheckBoxEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkCheckBoxEvents, Outlook.Contrib.Interfaces.IOlkCheckBoxEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkCheckBoxEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkCheckBoxEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkCheckBoxEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkCheckBoxEvents_Event, Outlook.Contrib.Interfaces.IOlkCheckBoxEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkCheckBox which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkCheckBox WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkCheckBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkCheckBox, Outlook.Contrib.Interfaces.IOlkCheckBox>();
		}

		/// <summary>
		/// Wrapper interface for _OlkOptionButton which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OlkOptionButton WithComCleanup(this Microsoft.Office.Interop.Outlook._OlkOptionButton resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OlkOptionButton, Outlook.Contrib.Interfaces.I_OlkOptionButton>();
		}

		/// <summary>
		/// Wrapper interface for OlkOptionButtonEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkOptionButtonEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkOptionButtonEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkOptionButtonEvents, Outlook.Contrib.Interfaces.IOlkOptionButtonEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkOptionButtonEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkOptionButtonEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkOptionButtonEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkOptionButtonEvents_Event, Outlook.Contrib.Interfaces.IOlkOptionButtonEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkOptionButton which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkOptionButton WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkOptionButton resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkOptionButton, Outlook.Contrib.Interfaces.IOlkOptionButton>();
		}

		/// <summary>
		/// Wrapper interface for _OlkComboBox which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OlkComboBox WithComCleanup(this Microsoft.Office.Interop.Outlook._OlkComboBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OlkComboBox, Outlook.Contrib.Interfaces.I_OlkComboBox>();
		}

		/// <summary>
		/// Wrapper interface for OlkComboBoxEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkComboBoxEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkComboBoxEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkComboBoxEvents, Outlook.Contrib.Interfaces.IOlkComboBoxEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkComboBoxEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkComboBoxEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkComboBoxEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkComboBoxEvents_Event, Outlook.Contrib.Interfaces.IOlkComboBoxEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkComboBox which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkComboBox WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkComboBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkComboBox, Outlook.Contrib.Interfaces.IOlkComboBox>();
		}

		/// <summary>
		/// Wrapper interface for _OlkListBox which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OlkListBox WithComCleanup(this Microsoft.Office.Interop.Outlook._OlkListBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OlkListBox, Outlook.Contrib.Interfaces.I_OlkListBox>();
		}

		/// <summary>
		/// Wrapper interface for OlkListBoxEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkListBoxEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkListBoxEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkListBoxEvents, Outlook.Contrib.Interfaces.IOlkListBoxEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkListBoxEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkListBoxEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkListBoxEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkListBoxEvents_Event, Outlook.Contrib.Interfaces.IOlkListBoxEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkListBox which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkListBox WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkListBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkListBox, Outlook.Contrib.Interfaces.IOlkListBox>();
		}

		/// <summary>
		/// Wrapper interface for _OlkInfoBar which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OlkInfoBar WithComCleanup(this Microsoft.Office.Interop.Outlook._OlkInfoBar resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OlkInfoBar, Outlook.Contrib.Interfaces.I_OlkInfoBar>();
		}

		/// <summary>
		/// Wrapper interface for OlkInfoBarEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkInfoBarEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkInfoBarEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkInfoBarEvents, Outlook.Contrib.Interfaces.IOlkInfoBarEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkInfoBarEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkInfoBarEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkInfoBarEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkInfoBarEvents_Event, Outlook.Contrib.Interfaces.IOlkInfoBarEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkInfoBar which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkInfoBar WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkInfoBar resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkInfoBar, Outlook.Contrib.Interfaces.IOlkInfoBar>();
		}

		/// <summary>
		/// Wrapper interface for _OlkContactPhoto which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OlkContactPhoto WithComCleanup(this Microsoft.Office.Interop.Outlook._OlkContactPhoto resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OlkContactPhoto, Outlook.Contrib.Interfaces.I_OlkContactPhoto>();
		}

		/// <summary>
		/// Wrapper interface for OlkContactPhotoEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkContactPhotoEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkContactPhotoEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkContactPhotoEvents, Outlook.Contrib.Interfaces.IOlkContactPhotoEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkContactPhotoEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkContactPhotoEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkContactPhotoEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkContactPhotoEvents_Event, Outlook.Contrib.Interfaces.IOlkContactPhotoEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkContactPhoto which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkContactPhoto WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkContactPhoto resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkContactPhoto, Outlook.Contrib.Interfaces.IOlkContactPhoto>();
		}

		/// <summary>
		/// Wrapper interface for _OlkBusinessCardControl which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OlkBusinessCardControl WithComCleanup(this Microsoft.Office.Interop.Outlook._OlkBusinessCardControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OlkBusinessCardControl, Outlook.Contrib.Interfaces.I_OlkBusinessCardControl>();
		}

		/// <summary>
		/// Wrapper interface for OlkBusinessCardControlEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkBusinessCardControlEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkBusinessCardControlEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkBusinessCardControlEvents, Outlook.Contrib.Interfaces.IOlkBusinessCardControlEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkBusinessCardControlEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkBusinessCardControlEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkBusinessCardControlEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkBusinessCardControlEvents_Event, Outlook.Contrib.Interfaces.IOlkBusinessCardControlEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkBusinessCardControl which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkBusinessCardControl WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkBusinessCardControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkBusinessCardControl, Outlook.Contrib.Interfaces.IOlkBusinessCardControl>();
		}

		/// <summary>
		/// Wrapper interface for _OlkPageControl which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OlkPageControl WithComCleanup(this Microsoft.Office.Interop.Outlook._OlkPageControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OlkPageControl, Outlook.Contrib.Interfaces.I_OlkPageControl>();
		}

		/// <summary>
		/// Wrapper interface for OlkPageControlEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkPageControlEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkPageControlEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkPageControlEvents, Outlook.Contrib.Interfaces.IOlkPageControlEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkPageControlEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkPageControlEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkPageControlEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkPageControlEvents_Event, Outlook.Contrib.Interfaces.IOlkPageControlEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkPageControl which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkPageControl WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkPageControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkPageControl, Outlook.Contrib.Interfaces.IOlkPageControl>();
		}

		/// <summary>
		/// Wrapper interface for _OlkDateControl which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OlkDateControl WithComCleanup(this Microsoft.Office.Interop.Outlook._OlkDateControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OlkDateControl, Outlook.Contrib.Interfaces.I_OlkDateControl>();
		}

		/// <summary>
		/// Wrapper interface for OlkDateControlEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkDateControlEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkDateControlEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkDateControlEvents, Outlook.Contrib.Interfaces.IOlkDateControlEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkDateControlEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkDateControlEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkDateControlEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkDateControlEvents_Event, Outlook.Contrib.Interfaces.IOlkDateControlEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkDateControl which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkDateControl WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkDateControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkDateControl, Outlook.Contrib.Interfaces.IOlkDateControl>();
		}

		/// <summary>
		/// Wrapper interface for _OlkTimeControl which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OlkTimeControl WithComCleanup(this Microsoft.Office.Interop.Outlook._OlkTimeControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OlkTimeControl, Outlook.Contrib.Interfaces.I_OlkTimeControl>();
		}

		/// <summary>
		/// Wrapper interface for OlkTimeControlEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkTimeControlEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkTimeControlEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkTimeControlEvents, Outlook.Contrib.Interfaces.IOlkTimeControlEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkTimeControlEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkTimeControlEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkTimeControlEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkTimeControlEvents_Event, Outlook.Contrib.Interfaces.IOlkTimeControlEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkTimeControl which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkTimeControl WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkTimeControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkTimeControl, Outlook.Contrib.Interfaces.IOlkTimeControl>();
		}

		/// <summary>
		/// Wrapper interface for _OlkCategory which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OlkCategory WithComCleanup(this Microsoft.Office.Interop.Outlook._OlkCategory resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OlkCategory, Outlook.Contrib.Interfaces.I_OlkCategory>();
		}

		/// <summary>
		/// Wrapper interface for OlkCategoryEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkCategoryEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkCategoryEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkCategoryEvents, Outlook.Contrib.Interfaces.IOlkCategoryEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkCategoryEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkCategoryEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkCategoryEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkCategoryEvents_Event, Outlook.Contrib.Interfaces.IOlkCategoryEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkCategory which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkCategory WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkCategory resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkCategory, Outlook.Contrib.Interfaces.IOlkCategory>();
		}

		/// <summary>
		/// Wrapper interface for _OlkFrameHeader which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OlkFrameHeader WithComCleanup(this Microsoft.Office.Interop.Outlook._OlkFrameHeader resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OlkFrameHeader, Outlook.Contrib.Interfaces.I_OlkFrameHeader>();
		}

		/// <summary>
		/// Wrapper interface for OlkFrameHeaderEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkFrameHeaderEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkFrameHeaderEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkFrameHeaderEvents, Outlook.Contrib.Interfaces.IOlkFrameHeaderEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkFrameHeaderEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkFrameHeaderEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkFrameHeaderEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkFrameHeaderEvents_Event, Outlook.Contrib.Interfaces.IOlkFrameHeaderEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkFrameHeader which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkFrameHeader WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkFrameHeader resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkFrameHeader, Outlook.Contrib.Interfaces.IOlkFrameHeader>();
		}

		/// <summary>
		/// Wrapper interface for _OlkSenderPhoto which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OlkSenderPhoto WithComCleanup(this Microsoft.Office.Interop.Outlook._OlkSenderPhoto resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OlkSenderPhoto, Outlook.Contrib.Interfaces.I_OlkSenderPhoto>();
		}

		/// <summary>
		/// Wrapper interface for OlkSenderPhotoEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkSenderPhotoEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkSenderPhotoEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkSenderPhotoEvents, Outlook.Contrib.Interfaces.IOlkSenderPhotoEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkSenderPhotoEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkSenderPhotoEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkSenderPhotoEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkSenderPhotoEvents_Event, Outlook.Contrib.Interfaces.IOlkSenderPhotoEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkSenderPhoto which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkSenderPhoto WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkSenderPhoto resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkSenderPhoto, Outlook.Contrib.Interfaces.IOlkSenderPhoto>();
		}

		/// <summary>
		/// Wrapper interface for _TimeZone which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_TimeZone WithComCleanup(this Microsoft.Office.Interop.Outlook._TimeZone resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._TimeZone, Outlook.Contrib.Interfaces.I_TimeZone>();
		}

		/// <summary>
		/// Wrapper interface for _Application which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Application WithComCleanup(this Microsoft.Office.Interop.Outlook._Application resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Application, Outlook.Contrib.Interfaces.I_Application>();
		}

		/// <summary>
		/// Wrapper interface for _NameSpace which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_NameSpace WithComCleanup(this Microsoft.Office.Interop.Outlook._NameSpace resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._NameSpace, Outlook.Contrib.Interfaces.I_NameSpace>();
		}

		/// <summary>
		/// Wrapper interface for Recipient which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IRecipient WithComCleanup(this Microsoft.Office.Interop.Outlook.Recipient resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Recipient, Outlook.Contrib.Interfaces.IRecipient>();
		}

		/// <summary>
		/// Wrapper interface for AddressEntry which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IAddressEntry WithComCleanup(this Microsoft.Office.Interop.Outlook.AddressEntry resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.AddressEntry, Outlook.Contrib.Interfaces.IAddressEntry>();
		}

		/// <summary>
		/// Wrapper interface for AddressEntries which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IAddressEntries WithComCleanup(this Microsoft.Office.Interop.Outlook.AddressEntries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.AddressEntries, Outlook.Contrib.Interfaces.IAddressEntries>();
		}

		/// <summary>
		/// Wrapper interface for _ContactItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_ContactItem WithComCleanup(this Microsoft.Office.Interop.Outlook._ContactItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._ContactItem, Outlook.Contrib.Interfaces.I_ContactItem>();
		}

		/// <summary>
		/// Wrapper interface for Actions which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IActions WithComCleanup(this Microsoft.Office.Interop.Outlook.Actions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Actions, Outlook.Contrib.Interfaces.IActions>();
		}

		/// <summary>
		/// Wrapper interface for Action which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IAction WithComCleanup(this Microsoft.Office.Interop.Outlook.Action resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Action, Outlook.Contrib.Interfaces.IAction>();
		}

		/// <summary>
		/// Wrapper interface for Attachments which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IAttachments WithComCleanup(this Microsoft.Office.Interop.Outlook.Attachments resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Attachments, Outlook.Contrib.Interfaces.IAttachments>();
		}

		/// <summary>
		/// Wrapper interface for Attachment which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IAttachment WithComCleanup(this Microsoft.Office.Interop.Outlook.Attachment resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Attachment, Outlook.Contrib.Interfaces.IAttachment>();
		}

		/// <summary>
		/// Wrapper interface for PropertyAccessor which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IPropertyAccessor WithComCleanup(this Microsoft.Office.Interop.Outlook.PropertyAccessor resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.PropertyAccessor, Outlook.Contrib.Interfaces.IPropertyAccessor>();
		}

		/// <summary>
		/// Wrapper interface for _PropertyAccessor which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_PropertyAccessor WithComCleanup(this Microsoft.Office.Interop.Outlook._PropertyAccessor resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._PropertyAccessor, Outlook.Contrib.Interfaces.I_PropertyAccessor>();
		}

		/// <summary>
		/// Wrapper interface for FormDescription which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IFormDescription WithComCleanup(this Microsoft.Office.Interop.Outlook.FormDescription resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.FormDescription, Outlook.Contrib.Interfaces.IFormDescription>();
		}

		/// <summary>
		/// Wrapper interface for _Inspector which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Inspector WithComCleanup(this Microsoft.Office.Interop.Outlook._Inspector resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Inspector, Outlook.Contrib.Interfaces.I_Inspector>();
		}

		/// <summary>
		/// Wrapper interface for _AttachmentSelection which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_AttachmentSelection WithComCleanup(this Microsoft.Office.Interop.Outlook._AttachmentSelection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._AttachmentSelection, Outlook.Contrib.Interfaces.I_AttachmentSelection>();
		}

		/// <summary>
		/// Wrapper interface for Selection which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ISelection WithComCleanup(this Microsoft.Office.Interop.Outlook.Selection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Selection, Outlook.Contrib.Interfaces.ISelection>();
		}

		/// <summary>
		/// Wrapper interface for UserProperties which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IUserProperties WithComCleanup(this Microsoft.Office.Interop.Outlook.UserProperties resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.UserProperties, Outlook.Contrib.Interfaces.IUserProperties>();
		}

		/// <summary>
		/// Wrapper interface for UserProperty which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IUserProperty WithComCleanup(this Microsoft.Office.Interop.Outlook.UserProperty resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.UserProperty, Outlook.Contrib.Interfaces.IUserProperty>();
		}

		/// <summary>
		/// Wrapper interface for MAPIFolder which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IMAPIFolder WithComCleanup(this Microsoft.Office.Interop.Outlook.MAPIFolder resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.MAPIFolder, Outlook.Contrib.Interfaces.IMAPIFolder>();
		}

		/// <summary>
		/// Wrapper interface for _Folders which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Folders WithComCleanup(this Microsoft.Office.Interop.Outlook._Folders resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Folders, Outlook.Contrib.Interfaces.I_Folders>();
		}

		/// <summary>
		/// Wrapper interface for _Items which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Items WithComCleanup(this Microsoft.Office.Interop.Outlook._Items resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Items, Outlook.Contrib.Interfaces.I_Items>();
		}

		/// <summary>
		/// Wrapper interface for _Explorer which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Explorer WithComCleanup(this Microsoft.Office.Interop.Outlook._Explorer resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Explorer, Outlook.Contrib.Interfaces.I_Explorer>();
		}

		/// <summary>
		/// Wrapper interface for Panes which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IPanes WithComCleanup(this Microsoft.Office.Interop.Outlook.Panes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Panes, Outlook.Contrib.Interfaces.IPanes>();
		}

		/// <summary>
		/// Wrapper interface for _NavigationPane which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_NavigationPane WithComCleanup(this Microsoft.Office.Interop.Outlook._NavigationPane resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._NavigationPane, Outlook.Contrib.Interfaces.I_NavigationPane>();
		}

		/// <summary>
		/// Wrapper interface for NavigationModule which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.INavigationModule WithComCleanup(this Microsoft.Office.Interop.Outlook.NavigationModule resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.NavigationModule, Outlook.Contrib.Interfaces.INavigationModule>();
		}

		/// <summary>
		/// Wrapper interface for _NavigationModule which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_NavigationModule WithComCleanup(this Microsoft.Office.Interop.Outlook._NavigationModule resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._NavigationModule, Outlook.Contrib.Interfaces.I_NavigationModule>();
		}

		/// <summary>
		/// Wrapper interface for NavigationModules which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.INavigationModules WithComCleanup(this Microsoft.Office.Interop.Outlook.NavigationModules resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.NavigationModules, Outlook.Contrib.Interfaces.INavigationModules>();
		}

		/// <summary>
		/// Wrapper interface for _NavigationModules which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_NavigationModules WithComCleanup(this Microsoft.Office.Interop.Outlook._NavigationModules resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._NavigationModules, Outlook.Contrib.Interfaces.I_NavigationModules>();
		}

		/// <summary>
		/// Wrapper interface for _AccountSelector which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_AccountSelector WithComCleanup(this Microsoft.Office.Interop.Outlook._AccountSelector resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._AccountSelector, Outlook.Contrib.Interfaces.I_AccountSelector>();
		}

		/// <summary>
		/// Wrapper interface for _Account which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Account WithComCleanup(this Microsoft.Office.Interop.Outlook._Account resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Account, Outlook.Contrib.Interfaces.I_Account>();
		}

		/// <summary>
		/// Wrapper interface for Store which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IStore WithComCleanup(this Microsoft.Office.Interop.Outlook.Store resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Store, Outlook.Contrib.Interfaces.IStore>();
		}

		/// <summary>
		/// Wrapper interface for _Store which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Store WithComCleanup(this Microsoft.Office.Interop.Outlook._Store resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Store, Outlook.Contrib.Interfaces.I_Store>();
		}

		/// <summary>
		/// Wrapper interface for Rules which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IRules WithComCleanup(this Microsoft.Office.Interop.Outlook.Rules resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Rules, Outlook.Contrib.Interfaces.IRules>();
		}

		/// <summary>
		/// Wrapper interface for _Rules which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Rules WithComCleanup(this Microsoft.Office.Interop.Outlook._Rules resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Rules, Outlook.Contrib.Interfaces.I_Rules>();
		}

		/// <summary>
		/// Wrapper interface for _Rule which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Rule WithComCleanup(this Microsoft.Office.Interop.Outlook._Rule resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Rule, Outlook.Contrib.Interfaces.I_Rule>();
		}

		/// <summary>
		/// Wrapper interface for RuleActions which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IRuleActions WithComCleanup(this Microsoft.Office.Interop.Outlook.RuleActions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.RuleActions, Outlook.Contrib.Interfaces.IRuleActions>();
		}

		/// <summary>
		/// Wrapper interface for _RuleActions which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_RuleActions WithComCleanup(this Microsoft.Office.Interop.Outlook._RuleActions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._RuleActions, Outlook.Contrib.Interfaces.I_RuleActions>();
		}

		/// <summary>
		/// Wrapper interface for _RuleAction which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_RuleAction WithComCleanup(this Microsoft.Office.Interop.Outlook._RuleAction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._RuleAction, Outlook.Contrib.Interfaces.I_RuleAction>();
		}

		/// <summary>
		/// Wrapper interface for MoveOrCopyRuleAction which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IMoveOrCopyRuleAction WithComCleanup(this Microsoft.Office.Interop.Outlook.MoveOrCopyRuleAction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.MoveOrCopyRuleAction, Outlook.Contrib.Interfaces.IMoveOrCopyRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for _MoveOrCopyRuleAction which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_MoveOrCopyRuleAction WithComCleanup(this Microsoft.Office.Interop.Outlook._MoveOrCopyRuleAction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._MoveOrCopyRuleAction, Outlook.Contrib.Interfaces.I_MoveOrCopyRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for RuleAction which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IRuleAction WithComCleanup(this Microsoft.Office.Interop.Outlook.RuleAction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.RuleAction, Outlook.Contrib.Interfaces.IRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for SendRuleAction which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ISendRuleAction WithComCleanup(this Microsoft.Office.Interop.Outlook.SendRuleAction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.SendRuleAction, Outlook.Contrib.Interfaces.ISendRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for _SendRuleAction which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_SendRuleAction WithComCleanup(this Microsoft.Office.Interop.Outlook._SendRuleAction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._SendRuleAction, Outlook.Contrib.Interfaces.I_SendRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for Recipients which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IRecipients WithComCleanup(this Microsoft.Office.Interop.Outlook.Recipients resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Recipients, Outlook.Contrib.Interfaces.IRecipients>();
		}

		/// <summary>
		/// Wrapper interface for AssignToCategoryRuleAction which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IAssignToCategoryRuleAction WithComCleanup(this Microsoft.Office.Interop.Outlook.AssignToCategoryRuleAction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.AssignToCategoryRuleAction, Outlook.Contrib.Interfaces.IAssignToCategoryRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for _AssignToCategoryRuleAction which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_AssignToCategoryRuleAction WithComCleanup(this Microsoft.Office.Interop.Outlook._AssignToCategoryRuleAction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._AssignToCategoryRuleAction, Outlook.Contrib.Interfaces.I_AssignToCategoryRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for PlaySoundRuleAction which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IPlaySoundRuleAction WithComCleanup(this Microsoft.Office.Interop.Outlook.PlaySoundRuleAction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.PlaySoundRuleAction, Outlook.Contrib.Interfaces.IPlaySoundRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for _PlaySoundRuleAction which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_PlaySoundRuleAction WithComCleanup(this Microsoft.Office.Interop.Outlook._PlaySoundRuleAction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._PlaySoundRuleAction, Outlook.Contrib.Interfaces.I_PlaySoundRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for MarkAsTaskRuleAction which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IMarkAsTaskRuleAction WithComCleanup(this Microsoft.Office.Interop.Outlook.MarkAsTaskRuleAction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.MarkAsTaskRuleAction, Outlook.Contrib.Interfaces.IMarkAsTaskRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for _MarkAsTaskRuleAction which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_MarkAsTaskRuleAction WithComCleanup(this Microsoft.Office.Interop.Outlook._MarkAsTaskRuleAction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._MarkAsTaskRuleAction, Outlook.Contrib.Interfaces.I_MarkAsTaskRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for NewItemAlertRuleAction which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.INewItemAlertRuleAction WithComCleanup(this Microsoft.Office.Interop.Outlook.NewItemAlertRuleAction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.NewItemAlertRuleAction, Outlook.Contrib.Interfaces.INewItemAlertRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for _NewItemAlertRuleAction which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_NewItemAlertRuleAction WithComCleanup(this Microsoft.Office.Interop.Outlook._NewItemAlertRuleAction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._NewItemAlertRuleAction, Outlook.Contrib.Interfaces.I_NewItemAlertRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for RuleConditions which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IRuleConditions WithComCleanup(this Microsoft.Office.Interop.Outlook.RuleConditions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.RuleConditions, Outlook.Contrib.Interfaces.IRuleConditions>();
		}

		/// <summary>
		/// Wrapper interface for _RuleConditions which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_RuleConditions WithComCleanup(this Microsoft.Office.Interop.Outlook._RuleConditions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._RuleConditions, Outlook.Contrib.Interfaces.I_RuleConditions>();
		}

		/// <summary>
		/// Wrapper interface for _RuleCondition which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_RuleCondition WithComCleanup(this Microsoft.Office.Interop.Outlook._RuleCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._RuleCondition, Outlook.Contrib.Interfaces.I_RuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for RuleCondition which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IRuleCondition WithComCleanup(this Microsoft.Office.Interop.Outlook.RuleCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.RuleCondition, Outlook.Contrib.Interfaces.IRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for ImportanceRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IImportanceRuleCondition WithComCleanup(this Microsoft.Office.Interop.Outlook.ImportanceRuleCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ImportanceRuleCondition, Outlook.Contrib.Interfaces.IImportanceRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for _ImportanceRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_ImportanceRuleCondition WithComCleanup(this Microsoft.Office.Interop.Outlook._ImportanceRuleCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._ImportanceRuleCondition, Outlook.Contrib.Interfaces.I_ImportanceRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for AccountRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IAccountRuleCondition WithComCleanup(this Microsoft.Office.Interop.Outlook.AccountRuleCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.AccountRuleCondition, Outlook.Contrib.Interfaces.IAccountRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for _AccountRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_AccountRuleCondition WithComCleanup(this Microsoft.Office.Interop.Outlook._AccountRuleCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._AccountRuleCondition, Outlook.Contrib.Interfaces.I_AccountRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for Account which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IAccount WithComCleanup(this Microsoft.Office.Interop.Outlook.Account resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Account, Outlook.Contrib.Interfaces.IAccount>();
		}

		/// <summary>
		/// Wrapper interface for TextRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ITextRuleCondition WithComCleanup(this Microsoft.Office.Interop.Outlook.TextRuleCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.TextRuleCondition, Outlook.Contrib.Interfaces.ITextRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for _TextRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_TextRuleCondition WithComCleanup(this Microsoft.Office.Interop.Outlook._TextRuleCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._TextRuleCondition, Outlook.Contrib.Interfaces.I_TextRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for CategoryRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ICategoryRuleCondition WithComCleanup(this Microsoft.Office.Interop.Outlook.CategoryRuleCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.CategoryRuleCondition, Outlook.Contrib.Interfaces.ICategoryRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for _CategoryRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_CategoryRuleCondition WithComCleanup(this Microsoft.Office.Interop.Outlook._CategoryRuleCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._CategoryRuleCondition, Outlook.Contrib.Interfaces.I_CategoryRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for FormNameRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IFormNameRuleCondition WithComCleanup(this Microsoft.Office.Interop.Outlook.FormNameRuleCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.FormNameRuleCondition, Outlook.Contrib.Interfaces.IFormNameRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for _FormNameRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_FormNameRuleCondition WithComCleanup(this Microsoft.Office.Interop.Outlook._FormNameRuleCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._FormNameRuleCondition, Outlook.Contrib.Interfaces.I_FormNameRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for ToOrFromRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IToOrFromRuleCondition WithComCleanup(this Microsoft.Office.Interop.Outlook.ToOrFromRuleCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ToOrFromRuleCondition, Outlook.Contrib.Interfaces.IToOrFromRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for _ToOrFromRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_ToOrFromRuleCondition WithComCleanup(this Microsoft.Office.Interop.Outlook._ToOrFromRuleCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._ToOrFromRuleCondition, Outlook.Contrib.Interfaces.I_ToOrFromRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for AddressRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IAddressRuleCondition WithComCleanup(this Microsoft.Office.Interop.Outlook.AddressRuleCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.AddressRuleCondition, Outlook.Contrib.Interfaces.IAddressRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for _AddressRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_AddressRuleCondition WithComCleanup(this Microsoft.Office.Interop.Outlook._AddressRuleCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._AddressRuleCondition, Outlook.Contrib.Interfaces.I_AddressRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for SenderInAddressListRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ISenderInAddressListRuleCondition WithComCleanup(this Microsoft.Office.Interop.Outlook.SenderInAddressListRuleCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.SenderInAddressListRuleCondition, Outlook.Contrib.Interfaces.ISenderInAddressListRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for _SenderInAddressListRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_SenderInAddressListRuleCondition WithComCleanup(this Microsoft.Office.Interop.Outlook._SenderInAddressListRuleCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._SenderInAddressListRuleCondition, Outlook.Contrib.Interfaces.I_SenderInAddressListRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for AddressList which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IAddressList WithComCleanup(this Microsoft.Office.Interop.Outlook.AddressList resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.AddressList, Outlook.Contrib.Interfaces.IAddressList>();
		}

		/// <summary>
		/// Wrapper interface for FromRssFeedRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IFromRssFeedRuleCondition WithComCleanup(this Microsoft.Office.Interop.Outlook.FromRssFeedRuleCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.FromRssFeedRuleCondition, Outlook.Contrib.Interfaces.IFromRssFeedRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for _FromRssFeedRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_FromRssFeedRuleCondition WithComCleanup(this Microsoft.Office.Interop.Outlook._FromRssFeedRuleCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._FromRssFeedRuleCondition, Outlook.Contrib.Interfaces.I_FromRssFeedRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for Rule which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IRule WithComCleanup(this Microsoft.Office.Interop.Outlook.Rule resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Rule, Outlook.Contrib.Interfaces.IRule>();
		}

		/// <summary>
		/// Wrapper interface for Categories which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ICategories WithComCleanup(this Microsoft.Office.Interop.Outlook.Categories resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Categories, Outlook.Contrib.Interfaces.ICategories>();
		}

		/// <summary>
		/// Wrapper interface for _Categories which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Categories WithComCleanup(this Microsoft.Office.Interop.Outlook._Categories resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Categories, Outlook.Contrib.Interfaces.I_Categories>();
		}

		/// <summary>
		/// Wrapper interface for _Category which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Category WithComCleanup(this Microsoft.Office.Interop.Outlook._Category resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Category, Outlook.Contrib.Interfaces.I_Category>();
		}

		/// <summary>
		/// Wrapper interface for Category which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ICategory WithComCleanup(this Microsoft.Office.Interop.Outlook.Category resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Category, Outlook.Contrib.Interfaces.ICategory>();
		}

		/// <summary>
		/// Wrapper interface for View which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IView WithComCleanup(this Microsoft.Office.Interop.Outlook.View resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.View, Outlook.Contrib.Interfaces.IView>();
		}

		/// <summary>
		/// Wrapper interface for _Views which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Views WithComCleanup(this Microsoft.Office.Interop.Outlook._Views resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Views, Outlook.Contrib.Interfaces.I_Views>();
		}

		/// <summary>
		/// Wrapper interface for _StorageItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_StorageItem WithComCleanup(this Microsoft.Office.Interop.Outlook._StorageItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._StorageItem, Outlook.Contrib.Interfaces.I_StorageItem>();
		}

		/// <summary>
		/// Wrapper interface for Table which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ITable WithComCleanup(this Microsoft.Office.Interop.Outlook.Table resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Table, Outlook.Contrib.Interfaces.ITable>();
		}

		/// <summary>
		/// Wrapper interface for _Table which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Table WithComCleanup(this Microsoft.Office.Interop.Outlook._Table resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Table, Outlook.Contrib.Interfaces.I_Table>();
		}

		/// <summary>
		/// Wrapper interface for Row which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IRow WithComCleanup(this Microsoft.Office.Interop.Outlook.Row resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Row, Outlook.Contrib.Interfaces.IRow>();
		}

		/// <summary>
		/// Wrapper interface for _Row which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Row WithComCleanup(this Microsoft.Office.Interop.Outlook._Row resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Row, Outlook.Contrib.Interfaces.I_Row>();
		}

		/// <summary>
		/// Wrapper interface for Columns which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IColumns WithComCleanup(this Microsoft.Office.Interop.Outlook.Columns resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Columns, Outlook.Contrib.Interfaces.IColumns>();
		}

		/// <summary>
		/// Wrapper interface for _Columns which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Columns WithComCleanup(this Microsoft.Office.Interop.Outlook._Columns resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Columns, Outlook.Contrib.Interfaces.I_Columns>();
		}

		/// <summary>
		/// Wrapper interface for _Column which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Column WithComCleanup(this Microsoft.Office.Interop.Outlook._Column resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Column, Outlook.Contrib.Interfaces.I_Column>();
		}

		/// <summary>
		/// Wrapper interface for Column which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IColumn WithComCleanup(this Microsoft.Office.Interop.Outlook.Column resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Column, Outlook.Contrib.Interfaces.IColumn>();
		}

		/// <summary>
		/// Wrapper interface for CalendarSharing which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ICalendarSharing WithComCleanup(this Microsoft.Office.Interop.Outlook.CalendarSharing resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.CalendarSharing, Outlook.Contrib.Interfaces.ICalendarSharing>();
		}

		/// <summary>
		/// Wrapper interface for _CalendarSharing which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_CalendarSharing WithComCleanup(this Microsoft.Office.Interop.Outlook._CalendarSharing resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._CalendarSharing, Outlook.Contrib.Interfaces.I_CalendarSharing>();
		}

		/// <summary>
		/// Wrapper interface for ItemEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IItemEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.ItemEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ItemEvents_Event, Outlook.Contrib.Interfaces.IItemEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for ItemEvents_10_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IItemEvents_10_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.ItemEvents_10_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ItemEvents_10_Event, Outlook.Contrib.Interfaces.IItemEvents_10_Event>();
		}

		/// <summary>
		/// Wrapper interface for MailItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IMailItem WithComCleanup(this Microsoft.Office.Interop.Outlook.MailItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.MailItem, Outlook.Contrib.Interfaces.IMailItem>();
		}

		/// <summary>
		/// Wrapper interface for _MailItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_MailItem WithComCleanup(this Microsoft.Office.Interop.Outlook._MailItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._MailItem, Outlook.Contrib.Interfaces.I_MailItem>();
		}

		/// <summary>
		/// Wrapper interface for Links which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ILinks WithComCleanup(this Microsoft.Office.Interop.Outlook.Links resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Links, Outlook.Contrib.Interfaces.ILinks>();
		}

		/// <summary>
		/// Wrapper interface for Link which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ILink WithComCleanup(this Microsoft.Office.Interop.Outlook.Link resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Link, Outlook.Contrib.Interfaces.ILink>();
		}

		/// <summary>
		/// Wrapper interface for ItemProperties which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IItemProperties WithComCleanup(this Microsoft.Office.Interop.Outlook.ItemProperties resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ItemProperties, Outlook.Contrib.Interfaces.IItemProperties>();
		}

		/// <summary>
		/// Wrapper interface for ItemProperty which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IItemProperty WithComCleanup(this Microsoft.Office.Interop.Outlook.ItemProperty resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ItemProperty, Outlook.Contrib.Interfaces.IItemProperty>();
		}

		/// <summary>
		/// Wrapper interface for Conflicts which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IConflicts WithComCleanup(this Microsoft.Office.Interop.Outlook.Conflicts resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Conflicts, Outlook.Contrib.Interfaces.IConflicts>();
		}

		/// <summary>
		/// Wrapper interface for Conflict which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IConflict WithComCleanup(this Microsoft.Office.Interop.Outlook.Conflict resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Conflict, Outlook.Contrib.Interfaces.IConflict>();
		}

		/// <summary>
		/// Wrapper interface for ContactItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IContactItem WithComCleanup(this Microsoft.Office.Interop.Outlook.ContactItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ContactItem, Outlook.Contrib.Interfaces.IContactItem>();
		}

		/// <summary>
		/// Wrapper interface for ItemEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IItemEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.ItemEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ItemEvents, Outlook.Contrib.Interfaces.IItemEvents>();
		}

		/// <summary>
		/// Wrapper interface for ItemEvents_10 which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IItemEvents_10 WithComCleanup(this Microsoft.Office.Interop.Outlook.ItemEvents_10 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ItemEvents_10, Outlook.Contrib.Interfaces.IItemEvents_10>();
		}

		/// <summary>
		/// Wrapper interface for _Conversation which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Conversation WithComCleanup(this Microsoft.Office.Interop.Outlook._Conversation resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Conversation, Outlook.Contrib.Interfaces.I_Conversation>();
		}

		/// <summary>
		/// Wrapper interface for SimpleItems which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ISimpleItems WithComCleanup(this Microsoft.Office.Interop.Outlook.SimpleItems resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.SimpleItems, Outlook.Contrib.Interfaces.ISimpleItems>();
		}

		/// <summary>
		/// Wrapper interface for _SimpleItems which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_SimpleItems WithComCleanup(this Microsoft.Office.Interop.Outlook._SimpleItems resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._SimpleItems, Outlook.Contrib.Interfaces.I_SimpleItems>();
		}

		/// <summary>
		/// Wrapper interface for UserDefinedProperties which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IUserDefinedProperties WithComCleanup(this Microsoft.Office.Interop.Outlook.UserDefinedProperties resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.UserDefinedProperties, Outlook.Contrib.Interfaces.IUserDefinedProperties>();
		}

		/// <summary>
		/// Wrapper interface for _UserDefinedProperties which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_UserDefinedProperties WithComCleanup(this Microsoft.Office.Interop.Outlook._UserDefinedProperties resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._UserDefinedProperties, Outlook.Contrib.Interfaces.I_UserDefinedProperties>();
		}

		/// <summary>
		/// Wrapper interface for _UserDefinedProperty which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_UserDefinedProperty WithComCleanup(this Microsoft.Office.Interop.Outlook._UserDefinedProperty resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._UserDefinedProperty, Outlook.Contrib.Interfaces.I_UserDefinedProperty>();
		}

		/// <summary>
		/// Wrapper interface for UserDefinedProperty which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IUserDefinedProperty WithComCleanup(this Microsoft.Office.Interop.Outlook.UserDefinedProperty resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.UserDefinedProperty, Outlook.Contrib.Interfaces.IUserDefinedProperty>();
		}

		/// <summary>
		/// Wrapper interface for ExchangeUser which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IExchangeUser WithComCleanup(this Microsoft.Office.Interop.Outlook.ExchangeUser resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ExchangeUser, Outlook.Contrib.Interfaces.IExchangeUser>();
		}

		/// <summary>
		/// Wrapper interface for _ExchangeUser which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_ExchangeUser WithComCleanup(this Microsoft.Office.Interop.Outlook._ExchangeUser resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._ExchangeUser, Outlook.Contrib.Interfaces.I_ExchangeUser>();
		}

		/// <summary>
		/// Wrapper interface for ExchangeDistributionList which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IExchangeDistributionList WithComCleanup(this Microsoft.Office.Interop.Outlook.ExchangeDistributionList resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ExchangeDistributionList, Outlook.Contrib.Interfaces.IExchangeDistributionList>();
		}

		/// <summary>
		/// Wrapper interface for _ExchangeDistributionList which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_ExchangeDistributionList WithComCleanup(this Microsoft.Office.Interop.Outlook._ExchangeDistributionList resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._ExchangeDistributionList, Outlook.Contrib.Interfaces.I_ExchangeDistributionList>();
		}

		/// <summary>
		/// Wrapper interface for AddressLists which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IAddressLists WithComCleanup(this Microsoft.Office.Interop.Outlook.AddressLists resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.AddressLists, Outlook.Contrib.Interfaces.IAddressLists>();
		}

		/// <summary>
		/// Wrapper interface for SyncObjects which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ISyncObjects WithComCleanup(this Microsoft.Office.Interop.Outlook.SyncObjects resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.SyncObjects, Outlook.Contrib.Interfaces.ISyncObjects>();
		}

		/// <summary>
		/// Wrapper interface for SyncObjectEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ISyncObjectEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.SyncObjectEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.SyncObjectEvents_Event, Outlook.Contrib.Interfaces.ISyncObjectEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for SyncObject which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ISyncObject WithComCleanup(this Microsoft.Office.Interop.Outlook.SyncObject resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.SyncObject, Outlook.Contrib.Interfaces.ISyncObject>();
		}

		/// <summary>
		/// Wrapper interface for _SyncObject which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_SyncObject WithComCleanup(this Microsoft.Office.Interop.Outlook._SyncObject resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._SyncObject, Outlook.Contrib.Interfaces.I_SyncObject>();
		}

		/// <summary>
		/// Wrapper interface for SyncObjectEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ISyncObjectEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.SyncObjectEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.SyncObjectEvents, Outlook.Contrib.Interfaces.ISyncObjectEvents>();
		}

		/// <summary>
		/// Wrapper interface for AccountsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IAccountsEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.AccountsEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.AccountsEvents_Event, Outlook.Contrib.Interfaces.IAccountsEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Accounts which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IAccounts WithComCleanup(this Microsoft.Office.Interop.Outlook.Accounts resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Accounts, Outlook.Contrib.Interfaces.IAccounts>();
		}

		/// <summary>
		/// Wrapper interface for _Accounts which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Accounts WithComCleanup(this Microsoft.Office.Interop.Outlook._Accounts resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Accounts, Outlook.Contrib.Interfaces.I_Accounts>();
		}

		/// <summary>
		/// Wrapper interface for AccountsEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IAccountsEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.AccountsEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.AccountsEvents, Outlook.Contrib.Interfaces.IAccountsEvents>();
		}

		/// <summary>
		/// Wrapper interface for StoresEvents_12_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IStoresEvents_12_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.StoresEvents_12_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.StoresEvents_12_Event, Outlook.Contrib.Interfaces.IStoresEvents_12_Event>();
		}

		/// <summary>
		/// Wrapper interface for Stores which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IStores WithComCleanup(this Microsoft.Office.Interop.Outlook.Stores resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Stores, Outlook.Contrib.Interfaces.IStores>();
		}

		/// <summary>
		/// Wrapper interface for _Stores which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Stores WithComCleanup(this Microsoft.Office.Interop.Outlook._Stores resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Stores, Outlook.Contrib.Interfaces.I_Stores>();
		}

		/// <summary>
		/// Wrapper interface for StoresEvents_12 which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IStoresEvents_12 WithComCleanup(this Microsoft.Office.Interop.Outlook.StoresEvents_12 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.StoresEvents_12, Outlook.Contrib.Interfaces.IStoresEvents_12>();
		}

		/// <summary>
		/// Wrapper interface for SelectNamesDialog which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ISelectNamesDialog WithComCleanup(this Microsoft.Office.Interop.Outlook.SelectNamesDialog resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.SelectNamesDialog, Outlook.Contrib.Interfaces.ISelectNamesDialog>();
		}

		/// <summary>
		/// Wrapper interface for _SelectNamesDialog which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_SelectNamesDialog WithComCleanup(this Microsoft.Office.Interop.Outlook._SelectNamesDialog resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._SelectNamesDialog, Outlook.Contrib.Interfaces.I_SelectNamesDialog>();
		}

		/// <summary>
		/// Wrapper interface for SharingItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ISharingItem WithComCleanup(this Microsoft.Office.Interop.Outlook.SharingItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.SharingItem, Outlook.Contrib.Interfaces.ISharingItem>();
		}

		/// <summary>
		/// Wrapper interface for _SharingItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_SharingItem WithComCleanup(this Microsoft.Office.Interop.Outlook._SharingItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._SharingItem, Outlook.Contrib.Interfaces.I_SharingItem>();
		}

		/// <summary>
		/// Wrapper interface for _Explorers which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Explorers WithComCleanup(this Microsoft.Office.Interop.Outlook._Explorers resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Explorers, Outlook.Contrib.Interfaces.I_Explorers>();
		}

		/// <summary>
		/// Wrapper interface for ExplorerEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IExplorerEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.ExplorerEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ExplorerEvents_Event, Outlook.Contrib.Interfaces.IExplorerEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for ExplorerEvents_10_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IExplorerEvents_10_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.ExplorerEvents_10_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ExplorerEvents_10_Event, Outlook.Contrib.Interfaces.IExplorerEvents_10_Event>();
		}

		/// <summary>
		/// Wrapper interface for Explorer which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IExplorer WithComCleanup(this Microsoft.Office.Interop.Outlook.Explorer resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Explorer, Outlook.Contrib.Interfaces.IExplorer>();
		}

		/// <summary>
		/// Wrapper interface for ExplorerEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IExplorerEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.ExplorerEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ExplorerEvents, Outlook.Contrib.Interfaces.IExplorerEvents>();
		}

		/// <summary>
		/// Wrapper interface for ExplorerEvents_10 which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IExplorerEvents_10 WithComCleanup(this Microsoft.Office.Interop.Outlook.ExplorerEvents_10 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ExplorerEvents_10, Outlook.Contrib.Interfaces.IExplorerEvents_10>();
		}

		/// <summary>
		/// Wrapper interface for _Inspectors which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Inspectors WithComCleanup(this Microsoft.Office.Interop.Outlook._Inspectors resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Inspectors, Outlook.Contrib.Interfaces.I_Inspectors>();
		}

		/// <summary>
		/// Wrapper interface for InspectorEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IInspectorEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.InspectorEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.InspectorEvents_Event, Outlook.Contrib.Interfaces.IInspectorEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for InspectorEvents_10_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IInspectorEvents_10_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.InspectorEvents_10_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.InspectorEvents_10_Event, Outlook.Contrib.Interfaces.IInspectorEvents_10_Event>();
		}

		/// <summary>
		/// Wrapper interface for Inspector which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IInspector WithComCleanup(this Microsoft.Office.Interop.Outlook.Inspector resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Inspector, Outlook.Contrib.Interfaces.IInspector>();
		}

		/// <summary>
		/// Wrapper interface for InspectorEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IInspectorEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.InspectorEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.InspectorEvents, Outlook.Contrib.Interfaces.IInspectorEvents>();
		}

		/// <summary>
		/// Wrapper interface for InspectorEvents_10 which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IInspectorEvents_10 WithComCleanup(this Microsoft.Office.Interop.Outlook.InspectorEvents_10 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.InspectorEvents_10, Outlook.Contrib.Interfaces.IInspectorEvents_10>();
		}

		/// <summary>
		/// Wrapper interface for Search which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ISearch WithComCleanup(this Microsoft.Office.Interop.Outlook.Search resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Search, Outlook.Contrib.Interfaces.ISearch>();
		}

		/// <summary>
		/// Wrapper interface for _Results which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Results WithComCleanup(this Microsoft.Office.Interop.Outlook._Results resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Results, Outlook.Contrib.Interfaces.I_Results>();
		}

		/// <summary>
		/// Wrapper interface for _Reminders which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Reminders WithComCleanup(this Microsoft.Office.Interop.Outlook._Reminders resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Reminders, Outlook.Contrib.Interfaces.I_Reminders>();
		}

		/// <summary>
		/// Wrapper interface for _Reminder which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_Reminder WithComCleanup(this Microsoft.Office.Interop.Outlook._Reminder resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._Reminder, Outlook.Contrib.Interfaces.I_Reminder>();
		}

		/// <summary>
		/// Wrapper interface for TimeZones which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ITimeZones WithComCleanup(this Microsoft.Office.Interop.Outlook.TimeZones resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.TimeZones, Outlook.Contrib.Interfaces.ITimeZones>();
		}

		/// <summary>
		/// Wrapper interface for _TimeZones which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_TimeZones WithComCleanup(this Microsoft.Office.Interop.Outlook._TimeZones resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._TimeZones, Outlook.Contrib.Interfaces.I_TimeZones>();
		}

		/// <summary>
		/// Wrapper interface for _OlkTimeZoneControl which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OlkTimeZoneControl WithComCleanup(this Microsoft.Office.Interop.Outlook._OlkTimeZoneControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OlkTimeZoneControl, Outlook.Contrib.Interfaces.I_OlkTimeZoneControl>();
		}

		/// <summary>
		/// Wrapper interface for OlkTimeZoneControlEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkTimeZoneControlEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkTimeZoneControlEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkTimeZoneControlEvents, Outlook.Contrib.Interfaces.IOlkTimeZoneControlEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkTimeZoneControlEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkTimeZoneControlEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkTimeZoneControlEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkTimeZoneControlEvents_Event, Outlook.Contrib.Interfaces.IOlkTimeZoneControlEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkTimeZoneControl which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOlkTimeZoneControl WithComCleanup(this Microsoft.Office.Interop.Outlook.OlkTimeZoneControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OlkTimeZoneControl, Outlook.Contrib.Interfaces.IOlkTimeZoneControl>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IApplicationEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.ApplicationEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ApplicationEvents, Outlook.Contrib.Interfaces.IApplicationEvents>();
		}

		/// <summary>
		/// Wrapper interface for PropertyPages which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IPropertyPages WithComCleanup(this Microsoft.Office.Interop.Outlook.PropertyPages resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.PropertyPages, Outlook.Contrib.Interfaces.IPropertyPages>();
		}

		/// <summary>
		/// Wrapper interface for RecurrencePattern which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IRecurrencePattern WithComCleanup(this Microsoft.Office.Interop.Outlook.RecurrencePattern resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.RecurrencePattern, Outlook.Contrib.Interfaces.IRecurrencePattern>();
		}

		/// <summary>
		/// Wrapper interface for Exceptions which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IExceptions WithComCleanup(this Microsoft.Office.Interop.Outlook.Exceptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Exceptions, Outlook.Contrib.Interfaces.IExceptions>();
		}

		/// <summary>
		/// Wrapper interface for Exception which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IException WithComCleanup(this Microsoft.Office.Interop.Outlook.Exception resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Exception, Outlook.Contrib.Interfaces.IException>();
		}

		/// <summary>
		/// Wrapper interface for AppointmentItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IAppointmentItem WithComCleanup(this Microsoft.Office.Interop.Outlook.AppointmentItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.AppointmentItem, Outlook.Contrib.Interfaces.IAppointmentItem>();
		}

		/// <summary>
		/// Wrapper interface for _AppointmentItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_AppointmentItem WithComCleanup(this Microsoft.Office.Interop.Outlook._AppointmentItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._AppointmentItem, Outlook.Contrib.Interfaces.I_AppointmentItem>();
		}

		/// <summary>
		/// Wrapper interface for MeetingItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IMeetingItem WithComCleanup(this Microsoft.Office.Interop.Outlook.MeetingItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.MeetingItem, Outlook.Contrib.Interfaces.IMeetingItem>();
		}

		/// <summary>
		/// Wrapper interface for _MeetingItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_MeetingItem WithComCleanup(this Microsoft.Office.Interop.Outlook._MeetingItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._MeetingItem, Outlook.Contrib.Interfaces.I_MeetingItem>();
		}

		/// <summary>
		/// Wrapper interface for ExplorersEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IExplorersEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.ExplorersEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ExplorersEvents, Outlook.Contrib.Interfaces.IExplorersEvents>();
		}

		/// <summary>
		/// Wrapper interface for FoldersEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IFoldersEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.FoldersEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.FoldersEvents, Outlook.Contrib.Interfaces.IFoldersEvents>();
		}

		/// <summary>
		/// Wrapper interface for InspectorsEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IInspectorsEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.InspectorsEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.InspectorsEvents, Outlook.Contrib.Interfaces.IInspectorsEvents>();
		}

		/// <summary>
		/// Wrapper interface for ItemsEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IItemsEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.ItemsEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ItemsEvents, Outlook.Contrib.Interfaces.IItemsEvents>();
		}

		/// <summary>
		/// Wrapper interface for NameSpaceEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.INameSpaceEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.NameSpaceEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.NameSpaceEvents, Outlook.Contrib.Interfaces.INameSpaceEvents>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarGroup which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOutlookBarGroup WithComCleanup(this Microsoft.Office.Interop.Outlook.OutlookBarGroup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OutlookBarGroup, Outlook.Contrib.Interfaces.IOutlookBarGroup>();
		}

		/// <summary>
		/// Wrapper interface for _OutlookBarShortcuts which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OutlookBarShortcuts WithComCleanup(this Microsoft.Office.Interop.Outlook._OutlookBarShortcuts resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OutlookBarShortcuts, Outlook.Contrib.Interfaces.I_OutlookBarShortcuts>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarShortcut which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOutlookBarShortcut WithComCleanup(this Microsoft.Office.Interop.Outlook.OutlookBarShortcut resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OutlookBarShortcut, Outlook.Contrib.Interfaces.IOutlookBarShortcut>();
		}

		/// <summary>
		/// Wrapper interface for _OutlookBarGroups which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OutlookBarGroups WithComCleanup(this Microsoft.Office.Interop.Outlook._OutlookBarGroups resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OutlookBarGroups, Outlook.Contrib.Interfaces.I_OutlookBarGroups>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarGroupsEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOutlookBarGroupsEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.OutlookBarGroupsEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OutlookBarGroupsEvents, Outlook.Contrib.Interfaces.IOutlookBarGroupsEvents>();
		}

		/// <summary>
		/// Wrapper interface for _OutlookBarPane which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OutlookBarPane WithComCleanup(this Microsoft.Office.Interop.Outlook._OutlookBarPane resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OutlookBarPane, Outlook.Contrib.Interfaces.I_OutlookBarPane>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarStorage which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOutlookBarStorage WithComCleanup(this Microsoft.Office.Interop.Outlook.OutlookBarStorage resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OutlookBarStorage, Outlook.Contrib.Interfaces.IOutlookBarStorage>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarPaneEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOutlookBarPaneEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.OutlookBarPaneEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OutlookBarPaneEvents, Outlook.Contrib.Interfaces.IOutlookBarPaneEvents>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarShortcutsEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOutlookBarShortcutsEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.OutlookBarShortcutsEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OutlookBarShortcutsEvents, Outlook.Contrib.Interfaces.IOutlookBarShortcutsEvents>();
		}

		/// <summary>
		/// Wrapper interface for PropertyPage which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IPropertyPage WithComCleanup(this Microsoft.Office.Interop.Outlook.PropertyPage resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.PropertyPage, Outlook.Contrib.Interfaces.IPropertyPage>();
		}

		/// <summary>
		/// Wrapper interface for PropertyPageSite which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IPropertyPageSite WithComCleanup(this Microsoft.Office.Interop.Outlook.PropertyPageSite resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.PropertyPageSite, Outlook.Contrib.Interfaces.IPropertyPageSite>();
		}

		/// <summary>
		/// Wrapper interface for Pages which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IPages WithComCleanup(this Microsoft.Office.Interop.Outlook.Pages resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Pages, Outlook.Contrib.Interfaces.IPages>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents_10 which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IApplicationEvents_10 WithComCleanup(this Microsoft.Office.Interop.Outlook.ApplicationEvents_10 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ApplicationEvents_10, Outlook.Contrib.Interfaces.IApplicationEvents_10>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents_11 which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IApplicationEvents_11 WithComCleanup(this Microsoft.Office.Interop.Outlook.ApplicationEvents_11 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ApplicationEvents_11, Outlook.Contrib.Interfaces.IApplicationEvents_11>();
		}

		/// <summary>
		/// Wrapper interface for AttachmentSelection which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IAttachmentSelection WithComCleanup(this Microsoft.Office.Interop.Outlook.AttachmentSelection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.AttachmentSelection, Outlook.Contrib.Interfaces.IAttachmentSelection>();
		}

		/// <summary>
		/// Wrapper interface for MAPIFolderEvents_12_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IMAPIFolderEvents_12_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.MAPIFolderEvents_12_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.MAPIFolderEvents_12_Event, Outlook.Contrib.Interfaces.IMAPIFolderEvents_12_Event>();
		}

		/// <summary>
		/// Wrapper interface for Folder which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IFolder WithComCleanup(this Microsoft.Office.Interop.Outlook.Folder resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Folder, Outlook.Contrib.Interfaces.IFolder>();
		}

		/// <summary>
		/// Wrapper interface for MAPIFolderEvents_12 which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IMAPIFolderEvents_12 WithComCleanup(this Microsoft.Office.Interop.Outlook.MAPIFolderEvents_12 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.MAPIFolderEvents_12, Outlook.Contrib.Interfaces.IMAPIFolderEvents_12>();
		}

		/// <summary>
		/// Wrapper interface for ResultsEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IResultsEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.ResultsEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ResultsEvents, Outlook.Contrib.Interfaces.IResultsEvents>();
		}

		/// <summary>
		/// Wrapper interface for _ViewsEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_ViewsEvents WithComCleanup(this Microsoft.Office.Interop.Outlook._ViewsEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._ViewsEvents, Outlook.Contrib.Interfaces.I_ViewsEvents>();
		}

		/// <summary>
		/// Wrapper interface for ReminderCollectionEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IReminderCollectionEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.ReminderCollectionEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ReminderCollectionEvents, Outlook.Contrib.Interfaces.IReminderCollectionEvents>();
		}

		/// <summary>
		/// Wrapper interface for _DocumentItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_DocumentItem WithComCleanup(this Microsoft.Office.Interop.Outlook._DocumentItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._DocumentItem, Outlook.Contrib.Interfaces.I_DocumentItem>();
		}

		/// <summary>
		/// Wrapper interface for _NoteItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_NoteItem WithComCleanup(this Microsoft.Office.Interop.Outlook._NoteItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._NoteItem, Outlook.Contrib.Interfaces.I_NoteItem>();
		}

		/// <summary>
		/// Wrapper interface for FormRegionEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IFormRegionEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.FormRegionEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.FormRegionEvents, Outlook.Contrib.Interfaces.IFormRegionEvents>();
		}

		/// <summary>
		/// Wrapper interface for _ViewField which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_ViewField WithComCleanup(this Microsoft.Office.Interop.Outlook._ViewField resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._ViewField, Outlook.Contrib.Interfaces.I_ViewField>();
		}

		/// <summary>
		/// Wrapper interface for ColumnFormat which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IColumnFormat WithComCleanup(this Microsoft.Office.Interop.Outlook.ColumnFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ColumnFormat, Outlook.Contrib.Interfaces.IColumnFormat>();
		}

		/// <summary>
		/// Wrapper interface for _ColumnFormat which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_ColumnFormat WithComCleanup(this Microsoft.Office.Interop.Outlook._ColumnFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._ColumnFormat, Outlook.Contrib.Interfaces.I_ColumnFormat>();
		}

		/// <summary>
		/// Wrapper interface for _ViewFields which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_ViewFields WithComCleanup(this Microsoft.Office.Interop.Outlook._ViewFields resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._ViewFields, Outlook.Contrib.Interfaces.I_ViewFields>();
		}

		/// <summary>
		/// Wrapper interface for ViewField which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IViewField WithComCleanup(this Microsoft.Office.Interop.Outlook.ViewField resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ViewField, Outlook.Contrib.Interfaces.IViewField>();
		}

		/// <summary>
		/// Wrapper interface for _IconView which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_IconView WithComCleanup(this Microsoft.Office.Interop.Outlook._IconView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._IconView, Outlook.Contrib.Interfaces.I_IconView>();
		}

		/// <summary>
		/// Wrapper interface for OrderFields which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOrderFields WithComCleanup(this Microsoft.Office.Interop.Outlook.OrderFields resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OrderFields, Outlook.Contrib.Interfaces.IOrderFields>();
		}

		/// <summary>
		/// Wrapper interface for _OrderFields which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OrderFields WithComCleanup(this Microsoft.Office.Interop.Outlook._OrderFields resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OrderFields, Outlook.Contrib.Interfaces.I_OrderFields>();
		}

		/// <summary>
		/// Wrapper interface for _OrderField which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_OrderField WithComCleanup(this Microsoft.Office.Interop.Outlook._OrderField resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._OrderField, Outlook.Contrib.Interfaces.I_OrderField>();
		}

		/// <summary>
		/// Wrapper interface for OrderField which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOrderField WithComCleanup(this Microsoft.Office.Interop.Outlook.OrderField resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OrderField, Outlook.Contrib.Interfaces.IOrderField>();
		}

		/// <summary>
		/// Wrapper interface for _CardView which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_CardView WithComCleanup(this Microsoft.Office.Interop.Outlook._CardView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._CardView, Outlook.Contrib.Interfaces.I_CardView>();
		}

		/// <summary>
		/// Wrapper interface for ViewFields which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IViewFields WithComCleanup(this Microsoft.Office.Interop.Outlook.ViewFields resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ViewFields, Outlook.Contrib.Interfaces.IViewFields>();
		}

		/// <summary>
		/// Wrapper interface for ViewFont which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IViewFont WithComCleanup(this Microsoft.Office.Interop.Outlook.ViewFont resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ViewFont, Outlook.Contrib.Interfaces.IViewFont>();
		}

		/// <summary>
		/// Wrapper interface for _ViewFont which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_ViewFont WithComCleanup(this Microsoft.Office.Interop.Outlook._ViewFont resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._ViewFont, Outlook.Contrib.Interfaces.I_ViewFont>();
		}

		/// <summary>
		/// Wrapper interface for AutoFormatRules which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IAutoFormatRules WithComCleanup(this Microsoft.Office.Interop.Outlook.AutoFormatRules resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.AutoFormatRules, Outlook.Contrib.Interfaces.IAutoFormatRules>();
		}

		/// <summary>
		/// Wrapper interface for _AutoFormatRules which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_AutoFormatRules WithComCleanup(this Microsoft.Office.Interop.Outlook._AutoFormatRules resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._AutoFormatRules, Outlook.Contrib.Interfaces.I_AutoFormatRules>();
		}

		/// <summary>
		/// Wrapper interface for AutoFormatRule which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IAutoFormatRule WithComCleanup(this Microsoft.Office.Interop.Outlook.AutoFormatRule resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.AutoFormatRule, Outlook.Contrib.Interfaces.IAutoFormatRule>();
		}

		/// <summary>
		/// Wrapper interface for _AutoFormatRule which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_AutoFormatRule WithComCleanup(this Microsoft.Office.Interop.Outlook._AutoFormatRule resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._AutoFormatRule, Outlook.Contrib.Interfaces.I_AutoFormatRule>();
		}

		/// <summary>
		/// Wrapper interface for _TimelineView which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_TimelineView WithComCleanup(this Microsoft.Office.Interop.Outlook._TimelineView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._TimelineView, Outlook.Contrib.Interfaces.I_TimelineView>();
		}

		/// <summary>
		/// Wrapper interface for _MailModule which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_MailModule WithComCleanup(this Microsoft.Office.Interop.Outlook._MailModule resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._MailModule, Outlook.Contrib.Interfaces.I_MailModule>();
		}

		/// <summary>
		/// Wrapper interface for _NavigationGroups which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_NavigationGroups WithComCleanup(this Microsoft.Office.Interop.Outlook._NavigationGroups resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._NavigationGroups, Outlook.Contrib.Interfaces.I_NavigationGroups>();
		}

		/// <summary>
		/// Wrapper interface for _NavigationGroup which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_NavigationGroup WithComCleanup(this Microsoft.Office.Interop.Outlook._NavigationGroup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._NavigationGroup, Outlook.Contrib.Interfaces.I_NavigationGroup>();
		}

		/// <summary>
		/// Wrapper interface for NavigationFolders which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.INavigationFolders WithComCleanup(this Microsoft.Office.Interop.Outlook.NavigationFolders resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.NavigationFolders, Outlook.Contrib.Interfaces.INavigationFolders>();
		}

		/// <summary>
		/// Wrapper interface for _NavigationFolders which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_NavigationFolders WithComCleanup(this Microsoft.Office.Interop.Outlook._NavigationFolders resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._NavigationFolders, Outlook.Contrib.Interfaces.I_NavigationFolders>();
		}

		/// <summary>
		/// Wrapper interface for _NavigationFolder which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_NavigationFolder WithComCleanup(this Microsoft.Office.Interop.Outlook._NavigationFolder resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._NavigationFolder, Outlook.Contrib.Interfaces.I_NavigationFolder>();
		}

		/// <summary>
		/// Wrapper interface for NavigationFolder which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.INavigationFolder WithComCleanup(this Microsoft.Office.Interop.Outlook.NavigationFolder resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.NavigationFolder, Outlook.Contrib.Interfaces.INavigationFolder>();
		}

		/// <summary>
		/// Wrapper interface for NavigationGroup which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.INavigationGroup WithComCleanup(this Microsoft.Office.Interop.Outlook.NavigationGroup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.NavigationGroup, Outlook.Contrib.Interfaces.INavigationGroup>();
		}

		/// <summary>
		/// Wrapper interface for _CalendarModule which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_CalendarModule WithComCleanup(this Microsoft.Office.Interop.Outlook._CalendarModule resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._CalendarModule, Outlook.Contrib.Interfaces.I_CalendarModule>();
		}

		/// <summary>
		/// Wrapper interface for _ContactsModule which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_ContactsModule WithComCleanup(this Microsoft.Office.Interop.Outlook._ContactsModule resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._ContactsModule, Outlook.Contrib.Interfaces.I_ContactsModule>();
		}

		/// <summary>
		/// Wrapper interface for _TasksModule which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_TasksModule WithComCleanup(this Microsoft.Office.Interop.Outlook._TasksModule resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._TasksModule, Outlook.Contrib.Interfaces.I_TasksModule>();
		}

		/// <summary>
		/// Wrapper interface for _JournalModule which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_JournalModule WithComCleanup(this Microsoft.Office.Interop.Outlook._JournalModule resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._JournalModule, Outlook.Contrib.Interfaces.I_JournalModule>();
		}

		/// <summary>
		/// Wrapper interface for _NotesModule which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_NotesModule WithComCleanup(this Microsoft.Office.Interop.Outlook._NotesModule resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._NotesModule, Outlook.Contrib.Interfaces.I_NotesModule>();
		}

		/// <summary>
		/// Wrapper interface for NavigationPaneEvents_12 which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.INavigationPaneEvents_12 WithComCleanup(this Microsoft.Office.Interop.Outlook.NavigationPaneEvents_12 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.NavigationPaneEvents_12, Outlook.Contrib.Interfaces.INavigationPaneEvents_12>();
		}

		/// <summary>
		/// Wrapper interface for NavigationGroupsEvents_12 which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.INavigationGroupsEvents_12 WithComCleanup(this Microsoft.Office.Interop.Outlook.NavigationGroupsEvents_12 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.NavigationGroupsEvents_12, Outlook.Contrib.Interfaces.INavigationGroupsEvents_12>();
		}

		/// <summary>
		/// Wrapper interface for _BusinessCardView which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_BusinessCardView WithComCleanup(this Microsoft.Office.Interop.Outlook._BusinessCardView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._BusinessCardView, Outlook.Contrib.Interfaces.I_BusinessCardView>();
		}

		/// <summary>
		/// Wrapper interface for _FormRegionStartup which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_FormRegionStartup WithComCleanup(this Microsoft.Office.Interop.Outlook._FormRegionStartup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._FormRegionStartup, Outlook.Contrib.Interfaces.I_FormRegionStartup>();
		}

		/// <summary>
		/// Wrapper interface for FormRegionEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IFormRegionEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.FormRegionEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.FormRegionEvents_Event, Outlook.Contrib.Interfaces.IFormRegionEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for FormRegion which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IFormRegion WithComCleanup(this Microsoft.Office.Interop.Outlook.FormRegion resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.FormRegion, Outlook.Contrib.Interfaces.IFormRegion>();
		}

		/// <summary>
		/// Wrapper interface for _FormRegion which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_FormRegion WithComCleanup(this Microsoft.Office.Interop.Outlook._FormRegion resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._FormRegion, Outlook.Contrib.Interfaces.I_FormRegion>();
		}

		/// <summary>
		/// Wrapper interface for _SolutionsModule which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_SolutionsModule WithComCleanup(this Microsoft.Office.Interop.Outlook._SolutionsModule resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._SolutionsModule, Outlook.Contrib.Interfaces.I_SolutionsModule>();
		}

		/// <summary>
		/// Wrapper interface for _CalendarView which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_CalendarView WithComCleanup(this Microsoft.Office.Interop.Outlook._CalendarView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._CalendarView, Outlook.Contrib.Interfaces.I_CalendarView>();
		}

		/// <summary>
		/// Wrapper interface for _TableView which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_TableView WithComCleanup(this Microsoft.Office.Interop.Outlook._TableView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._TableView, Outlook.Contrib.Interfaces.I_TableView>();
		}

		/// <summary>
		/// Wrapper interface for _MobileItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_MobileItem WithComCleanup(this Microsoft.Office.Interop.Outlook._MobileItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._MobileItem, Outlook.Contrib.Interfaces.I_MobileItem>();
		}

		/// <summary>
		/// Wrapper interface for MobileItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IMobileItem WithComCleanup(this Microsoft.Office.Interop.Outlook.MobileItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.MobileItem, Outlook.Contrib.Interfaces.IMobileItem>();
		}

		/// <summary>
		/// Wrapper interface for _JournalItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_JournalItem WithComCleanup(this Microsoft.Office.Interop.Outlook._JournalItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._JournalItem, Outlook.Contrib.Interfaces.I_JournalItem>();
		}

		/// <summary>
		/// Wrapper interface for _PostItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_PostItem WithComCleanup(this Microsoft.Office.Interop.Outlook._PostItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._PostItem, Outlook.Contrib.Interfaces.I_PostItem>();
		}

		/// <summary>
		/// Wrapper interface for _TaskItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_TaskItem WithComCleanup(this Microsoft.Office.Interop.Outlook._TaskItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._TaskItem, Outlook.Contrib.Interfaces.I_TaskItem>();
		}

		/// <summary>
		/// Wrapper interface for TaskItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ITaskItem WithComCleanup(this Microsoft.Office.Interop.Outlook.TaskItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.TaskItem, Outlook.Contrib.Interfaces.ITaskItem>();
		}

		/// <summary>
		/// Wrapper interface for AccountSelectorEvents which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IAccountSelectorEvents WithComCleanup(this Microsoft.Office.Interop.Outlook.AccountSelectorEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.AccountSelectorEvents, Outlook.Contrib.Interfaces.IAccountSelectorEvents>();
		}

		/// <summary>
		/// Wrapper interface for _DistListItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_DistListItem WithComCleanup(this Microsoft.Office.Interop.Outlook._DistListItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._DistListItem, Outlook.Contrib.Interfaces.I_DistListItem>();
		}

		/// <summary>
		/// Wrapper interface for _ReportItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_ReportItem WithComCleanup(this Microsoft.Office.Interop.Outlook._ReportItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._ReportItem, Outlook.Contrib.Interfaces.I_ReportItem>();
		}

		/// <summary>
		/// Wrapper interface for _RemoteItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_RemoteItem WithComCleanup(this Microsoft.Office.Interop.Outlook._RemoteItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._RemoteItem, Outlook.Contrib.Interfaces.I_RemoteItem>();
		}

		/// <summary>
		/// Wrapper interface for _TaskRequestItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_TaskRequestItem WithComCleanup(this Microsoft.Office.Interop.Outlook._TaskRequestItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._TaskRequestItem, Outlook.Contrib.Interfaces.I_TaskRequestItem>();
		}

		/// <summary>
		/// Wrapper interface for _TaskRequestAcceptItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_TaskRequestAcceptItem WithComCleanup(this Microsoft.Office.Interop.Outlook._TaskRequestAcceptItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._TaskRequestAcceptItem, Outlook.Contrib.Interfaces.I_TaskRequestAcceptItem>();
		}

		/// <summary>
		/// Wrapper interface for _TaskRequestDeclineItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_TaskRequestDeclineItem WithComCleanup(this Microsoft.Office.Interop.Outlook._TaskRequestDeclineItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._TaskRequestDeclineItem, Outlook.Contrib.Interfaces.I_TaskRequestDeclineItem>();
		}

		/// <summary>
		/// Wrapper interface for _TaskRequestUpdateItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_TaskRequestUpdateItem WithComCleanup(this Microsoft.Office.Interop.Outlook._TaskRequestUpdateItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._TaskRequestUpdateItem, Outlook.Contrib.Interfaces.I_TaskRequestUpdateItem>();
		}

		/// <summary>
		/// Wrapper interface for _ConversationHeader which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_ConversationHeader WithComCleanup(this Microsoft.Office.Interop.Outlook._ConversationHeader resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._ConversationHeader, Outlook.Contrib.Interfaces.I_ConversationHeader>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IApplicationEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.ApplicationEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ApplicationEvents_Event, Outlook.Contrib.Interfaces.IApplicationEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents_10_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IApplicationEvents_10_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.ApplicationEvents_10_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ApplicationEvents_10_Event, Outlook.Contrib.Interfaces.IApplicationEvents_10_Event>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents_11_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IApplicationEvents_11_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.ApplicationEvents_11_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ApplicationEvents_11_Event, Outlook.Contrib.Interfaces.IApplicationEvents_11_Event>();
		}

		/// <summary>
		/// Wrapper interface for Application which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IApplication WithComCleanup(this Microsoft.Office.Interop.Outlook.Application resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Application, Outlook.Contrib.Interfaces.IApplication>();
		}

		/// <summary>
		/// Wrapper interface for DistListItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IDistListItem WithComCleanup(this Microsoft.Office.Interop.Outlook.DistListItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.DistListItem, Outlook.Contrib.Interfaces.IDistListItem>();
		}

		/// <summary>
		/// Wrapper interface for DocumentItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IDocumentItem WithComCleanup(this Microsoft.Office.Interop.Outlook.DocumentItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.DocumentItem, Outlook.Contrib.Interfaces.IDocumentItem>();
		}

		/// <summary>
		/// Wrapper interface for ExplorersEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IExplorersEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.ExplorersEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ExplorersEvents_Event, Outlook.Contrib.Interfaces.IExplorersEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Explorers which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IExplorers WithComCleanup(this Microsoft.Office.Interop.Outlook.Explorers resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Explorers, Outlook.Contrib.Interfaces.IExplorers>();
		}

		/// <summary>
		/// Wrapper interface for InspectorsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IInspectorsEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.InspectorsEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.InspectorsEvents_Event, Outlook.Contrib.Interfaces.IInspectorsEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Inspectors which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IInspectors WithComCleanup(this Microsoft.Office.Interop.Outlook.Inspectors resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Inspectors, Outlook.Contrib.Interfaces.IInspectors>();
		}

		/// <summary>
		/// Wrapper interface for FoldersEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IFoldersEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.FoldersEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.FoldersEvents_Event, Outlook.Contrib.Interfaces.IFoldersEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Folders which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IFolders WithComCleanup(this Microsoft.Office.Interop.Outlook.Folders resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Folders, Outlook.Contrib.Interfaces.IFolders>();
		}

		/// <summary>
		/// Wrapper interface for ItemsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IItemsEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.ItemsEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ItemsEvents_Event, Outlook.Contrib.Interfaces.IItemsEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Items which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IItems WithComCleanup(this Microsoft.Office.Interop.Outlook.Items resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Items, Outlook.Contrib.Interfaces.IItems>();
		}

		/// <summary>
		/// Wrapper interface for JournalItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IJournalItem WithComCleanup(this Microsoft.Office.Interop.Outlook.JournalItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.JournalItem, Outlook.Contrib.Interfaces.IJournalItem>();
		}

		/// <summary>
		/// Wrapper interface for NameSpaceEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.INameSpaceEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.NameSpaceEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.NameSpaceEvents_Event, Outlook.Contrib.Interfaces.INameSpaceEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for NameSpace which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.INameSpace WithComCleanup(this Microsoft.Office.Interop.Outlook.NameSpace resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.NameSpace, Outlook.Contrib.Interfaces.INameSpace>();
		}

		/// <summary>
		/// Wrapper interface for NoteItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.INoteItem WithComCleanup(this Microsoft.Office.Interop.Outlook.NoteItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.NoteItem, Outlook.Contrib.Interfaces.INoteItem>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarGroupsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOutlookBarGroupsEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.OutlookBarGroupsEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OutlookBarGroupsEvents_Event, Outlook.Contrib.Interfaces.IOutlookBarGroupsEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarGroups which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOutlookBarGroups WithComCleanup(this Microsoft.Office.Interop.Outlook.OutlookBarGroups resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OutlookBarGroups, Outlook.Contrib.Interfaces.IOutlookBarGroups>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarPaneEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOutlookBarPaneEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.OutlookBarPaneEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OutlookBarPaneEvents_Event, Outlook.Contrib.Interfaces.IOutlookBarPaneEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarPane which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOutlookBarPane WithComCleanup(this Microsoft.Office.Interop.Outlook.OutlookBarPane resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OutlookBarPane, Outlook.Contrib.Interfaces.IOutlookBarPane>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarShortcutsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOutlookBarShortcutsEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.OutlookBarShortcutsEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OutlookBarShortcutsEvents_Event, Outlook.Contrib.Interfaces.IOutlookBarShortcutsEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarShortcuts which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IOutlookBarShortcuts WithComCleanup(this Microsoft.Office.Interop.Outlook.OutlookBarShortcuts resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.OutlookBarShortcuts, Outlook.Contrib.Interfaces.IOutlookBarShortcuts>();
		}

		/// <summary>
		/// Wrapper interface for PostItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IPostItem WithComCleanup(this Microsoft.Office.Interop.Outlook.PostItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.PostItem, Outlook.Contrib.Interfaces.IPostItem>();
		}

		/// <summary>
		/// Wrapper interface for RemoteItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IRemoteItem WithComCleanup(this Microsoft.Office.Interop.Outlook.RemoteItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.RemoteItem, Outlook.Contrib.Interfaces.IRemoteItem>();
		}

		/// <summary>
		/// Wrapper interface for ReportItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IReportItem WithComCleanup(this Microsoft.Office.Interop.Outlook.ReportItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ReportItem, Outlook.Contrib.Interfaces.IReportItem>();
		}

		/// <summary>
		/// Wrapper interface for TaskRequestAcceptItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ITaskRequestAcceptItem WithComCleanup(this Microsoft.Office.Interop.Outlook.TaskRequestAcceptItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.TaskRequestAcceptItem, Outlook.Contrib.Interfaces.ITaskRequestAcceptItem>();
		}

		/// <summary>
		/// Wrapper interface for TaskRequestDeclineItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ITaskRequestDeclineItem WithComCleanup(this Microsoft.Office.Interop.Outlook.TaskRequestDeclineItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.TaskRequestDeclineItem, Outlook.Contrib.Interfaces.ITaskRequestDeclineItem>();
		}

		/// <summary>
		/// Wrapper interface for TaskRequestItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ITaskRequestItem WithComCleanup(this Microsoft.Office.Interop.Outlook.TaskRequestItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.TaskRequestItem, Outlook.Contrib.Interfaces.ITaskRequestItem>();
		}

		/// <summary>
		/// Wrapper interface for TaskRequestUpdateItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ITaskRequestUpdateItem WithComCleanup(this Microsoft.Office.Interop.Outlook.TaskRequestUpdateItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.TaskRequestUpdateItem, Outlook.Contrib.Interfaces.ITaskRequestUpdateItem>();
		}

		/// <summary>
		/// Wrapper interface for ResultsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IResultsEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.ResultsEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ResultsEvents_Event, Outlook.Contrib.Interfaces.IResultsEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Results which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IResults WithComCleanup(this Microsoft.Office.Interop.Outlook.Results resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Results, Outlook.Contrib.Interfaces.IResults>();
		}

		/// <summary>
		/// Wrapper interface for _ViewsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.I_ViewsEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook._ViewsEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook._ViewsEvents_Event, Outlook.Contrib.Interfaces.I_ViewsEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Views which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IViews WithComCleanup(this Microsoft.Office.Interop.Outlook.Views resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Views, Outlook.Contrib.Interfaces.IViews>();
		}

		/// <summary>
		/// Wrapper interface for Reminder which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IReminder WithComCleanup(this Microsoft.Office.Interop.Outlook.Reminder resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Reminder, Outlook.Contrib.Interfaces.IReminder>();
		}

		/// <summary>
		/// Wrapper interface for ReminderCollectionEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IReminderCollectionEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.ReminderCollectionEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ReminderCollectionEvents_Event, Outlook.Contrib.Interfaces.IReminderCollectionEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Reminders which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IReminders WithComCleanup(this Microsoft.Office.Interop.Outlook.Reminders resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Reminders, Outlook.Contrib.Interfaces.IReminders>();
		}

		/// <summary>
		/// Wrapper interface for StorageItem which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IStorageItem WithComCleanup(this Microsoft.Office.Interop.Outlook.StorageItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.StorageItem, Outlook.Contrib.Interfaces.IStorageItem>();
		}

		/// <summary>
		/// Wrapper interface for NavigationPaneEvents_12_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.INavigationPaneEvents_12_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.NavigationPaneEvents_12_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.NavigationPaneEvents_12_Event, Outlook.Contrib.Interfaces.INavigationPaneEvents_12_Event>();
		}

		/// <summary>
		/// Wrapper interface for NavigationPane which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.INavigationPane WithComCleanup(this Microsoft.Office.Interop.Outlook.NavigationPane resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.NavigationPane, Outlook.Contrib.Interfaces.INavigationPane>();
		}

		/// <summary>
		/// Wrapper interface for NavigationGroupsEvents_12_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.INavigationGroupsEvents_12_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.NavigationGroupsEvents_12_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.NavigationGroupsEvents_12_Event, Outlook.Contrib.Interfaces.INavigationGroupsEvents_12_Event>();
		}

		/// <summary>
		/// Wrapper interface for NavigationGroups which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.INavigationGroups WithComCleanup(this Microsoft.Office.Interop.Outlook.NavigationGroups resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.NavigationGroups, Outlook.Contrib.Interfaces.INavigationGroups>();
		}

		/// <summary>
		/// Wrapper interface for DoNotUseMeFolder which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IDoNotUseMeFolder WithComCleanup(this Microsoft.Office.Interop.Outlook.DoNotUseMeFolder resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.DoNotUseMeFolder, Outlook.Contrib.Interfaces.IDoNotUseMeFolder>();
		}

		/// <summary>
		/// Wrapper interface for TimelineView which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ITimelineView WithComCleanup(this Microsoft.Office.Interop.Outlook.TimelineView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.TimelineView, Outlook.Contrib.Interfaces.ITimelineView>();
		}

		/// <summary>
		/// Wrapper interface for MailModule which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IMailModule WithComCleanup(this Microsoft.Office.Interop.Outlook.MailModule resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.MailModule, Outlook.Contrib.Interfaces.IMailModule>();
		}

		/// <summary>
		/// Wrapper interface for CalendarModule which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ICalendarModule WithComCleanup(this Microsoft.Office.Interop.Outlook.CalendarModule resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.CalendarModule, Outlook.Contrib.Interfaces.ICalendarModule>();
		}

		/// <summary>
		/// Wrapper interface for ContactsModule which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IContactsModule WithComCleanup(this Microsoft.Office.Interop.Outlook.ContactsModule resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ContactsModule, Outlook.Contrib.Interfaces.IContactsModule>();
		}

		/// <summary>
		/// Wrapper interface for TasksModule which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ITasksModule WithComCleanup(this Microsoft.Office.Interop.Outlook.TasksModule resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.TasksModule, Outlook.Contrib.Interfaces.ITasksModule>();
		}

		/// <summary>
		/// Wrapper interface for JournalModule which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IJournalModule WithComCleanup(this Microsoft.Office.Interop.Outlook.JournalModule resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.JournalModule, Outlook.Contrib.Interfaces.IJournalModule>();
		}

		/// <summary>
		/// Wrapper interface for NotesModule which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.INotesModule WithComCleanup(this Microsoft.Office.Interop.Outlook.NotesModule resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.NotesModule, Outlook.Contrib.Interfaces.INotesModule>();
		}

		/// <summary>
		/// Wrapper interface for TableView which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ITableView WithComCleanup(this Microsoft.Office.Interop.Outlook.TableView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.TableView, Outlook.Contrib.Interfaces.ITableView>();
		}

		/// <summary>
		/// Wrapper interface for IconView which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IIconView WithComCleanup(this Microsoft.Office.Interop.Outlook.IconView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.IconView, Outlook.Contrib.Interfaces.IIconView>();
		}

		/// <summary>
		/// Wrapper interface for CardView which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ICardView WithComCleanup(this Microsoft.Office.Interop.Outlook.CardView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.CardView, Outlook.Contrib.Interfaces.ICardView>();
		}

		/// <summary>
		/// Wrapper interface for CalendarView which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ICalendarView WithComCleanup(this Microsoft.Office.Interop.Outlook.CalendarView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.CalendarView, Outlook.Contrib.Interfaces.ICalendarView>();
		}

		/// <summary>
		/// Wrapper interface for BusinessCardView which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IBusinessCardView WithComCleanup(this Microsoft.Office.Interop.Outlook.BusinessCardView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.BusinessCardView, Outlook.Contrib.Interfaces.IBusinessCardView>();
		}

		/// <summary>
		/// Wrapper interface for FormRegionStartup which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IFormRegionStartup WithComCleanup(this Microsoft.Office.Interop.Outlook.FormRegionStartup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.FormRegionStartup, Outlook.Contrib.Interfaces.IFormRegionStartup>();
		}

		/// <summary>
		/// Wrapper interface for TimeZone which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ITimeZone WithComCleanup(this Microsoft.Office.Interop.Outlook.TimeZone resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.TimeZone, Outlook.Contrib.Interfaces.ITimeZone>();
		}

		/// <summary>
		/// Wrapper interface for SolutionsModule which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.ISolutionsModule WithComCleanup(this Microsoft.Office.Interop.Outlook.SolutionsModule resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.SolutionsModule, Outlook.Contrib.Interfaces.ISolutionsModule>();
		}

		/// <summary>
		/// Wrapper interface for Conversation which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IConversation WithComCleanup(this Microsoft.Office.Interop.Outlook.Conversation resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.Conversation, Outlook.Contrib.Interfaces.IConversation>();
		}

		/// <summary>
		/// Wrapper interface for AccountSelectorEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IAccountSelectorEvents_Event WithComCleanup(this Microsoft.Office.Interop.Outlook.AccountSelectorEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.AccountSelectorEvents_Event, Outlook.Contrib.Interfaces.IAccountSelectorEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for AccountSelector which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IAccountSelector WithComCleanup(this Microsoft.Office.Interop.Outlook.AccountSelector resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.AccountSelector, Outlook.Contrib.Interfaces.IAccountSelector>();
		}

		/// <summary>
		/// Wrapper interface for ConversationHeader which adds IDispose to the interface
		/// </summary>
		public static Outlook.Contrib.Interfaces.IConversationHeader WithComCleanup(this Microsoft.Office.Interop.Outlook.ConversationHeader resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Outlook.ConversationHeader, Outlook.Contrib.Interfaces.IConversationHeader>();
		}

	}
}