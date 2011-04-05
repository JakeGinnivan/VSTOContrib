using VSTOContrib.Extensions.Proxies;

//Microsoft.Office.Interop.Outlook, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c
namespace VSTOContrib.Outlook.Extensions.Proxies
{
	/// <summary>
	/// Provides cleanup extension methods for interfaces exposed by Microsoft.Office.Interop.Outlook.dll
	/// </summary>
	public static class OutlookCleanupExtensions
	{
		/// <summary>
		/// Wrapper interface for _IRecipientControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_IRecipientControl WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._IRecipientControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._IRecipientControl, Interfaces.I_IRecipientControl>();
		}

		/// <summary>
		/// Wrapper interface for _DRecipientControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_DRecipientControl WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._DRecipientControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._DRecipientControl, Interfaces.I_DRecipientControl>();
		}

		/// <summary>
		/// Wrapper interface for _DRecipientControlEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_DRecipientControlEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._DRecipientControlEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._DRecipientControlEvents, Interfaces.I_DRecipientControlEvents>();
		}

		/// <summary>
		/// Wrapper interface for _DRecipientControlEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_DRecipientControlEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._DRecipientControlEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._DRecipientControlEvents_Event, Interfaces.I_DRecipientControlEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for _RecipientControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_RecipientControl WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._RecipientControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._RecipientControl, Interfaces.I_RecipientControl>();
		}

		/// <summary>
		/// Wrapper interface for _IDocSiteControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_IDocSiteControl WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._IDocSiteControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._IDocSiteControl, Interfaces.I_IDocSiteControl>();
		}

		/// <summary>
		/// Wrapper interface for _DDocSiteControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_DDocSiteControl WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._DDocSiteControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._DDocSiteControl, Interfaces.I_DDocSiteControl>();
		}

		/// <summary>
		/// Wrapper interface for _DDocSiteControlEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_DDocSiteControlEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._DDocSiteControlEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._DDocSiteControlEvents, Interfaces.I_DDocSiteControlEvents>();
		}

		/// <summary>
		/// Wrapper interface for _DDocSiteControlEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_DDocSiteControlEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._DDocSiteControlEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._DDocSiteControlEvents_Event, Interfaces.I_DDocSiteControlEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for _DocSiteControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_DocSiteControl WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._DocSiteControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._DocSiteControl, Interfaces.I_DocSiteControl>();
		}

		/// <summary>
		/// Wrapper interface for OlkControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkControl WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkControl, Interfaces.IOlkControl>();
		}

		/// <summary>
		/// Wrapper interface for _OlkTextBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OlkTextBox WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OlkTextBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OlkTextBox, Interfaces.I_OlkTextBox>();
		}

		/// <summary>
		/// Wrapper interface for OlkTextBoxEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkTextBoxEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkTextBoxEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkTextBoxEvents, Interfaces.IOlkTextBoxEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkTextBoxEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkTextBoxEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkTextBoxEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkTextBoxEvents_Event, Interfaces.IOlkTextBoxEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkTextBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkTextBox WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkTextBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkTextBox, Interfaces.IOlkTextBox>();
		}

		/// <summary>
		/// Wrapper interface for _OlkLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OlkLabel WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OlkLabel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OlkLabel, Interfaces.I_OlkLabel>();
		}

		/// <summary>
		/// Wrapper interface for OlkLabelEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkLabelEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkLabelEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkLabelEvents, Interfaces.IOlkLabelEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkLabelEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkLabelEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkLabelEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkLabelEvents_Event, Interfaces.IOlkLabelEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkLabel WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkLabel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkLabel, Interfaces.IOlkLabel>();
		}

		/// <summary>
		/// Wrapper interface for _OlkCommandButton which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OlkCommandButton WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OlkCommandButton resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OlkCommandButton, Interfaces.I_OlkCommandButton>();
		}

		/// <summary>
		/// Wrapper interface for OlkCommandButtonEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkCommandButtonEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkCommandButtonEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkCommandButtonEvents, Interfaces.IOlkCommandButtonEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkCommandButtonEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkCommandButtonEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkCommandButtonEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkCommandButtonEvents_Event, Interfaces.IOlkCommandButtonEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkCommandButton which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkCommandButton WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkCommandButton resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkCommandButton, Interfaces.IOlkCommandButton>();
		}

		/// <summary>
		/// Wrapper interface for _OlkCheckBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OlkCheckBox WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OlkCheckBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OlkCheckBox, Interfaces.I_OlkCheckBox>();
		}

		/// <summary>
		/// Wrapper interface for OlkCheckBoxEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkCheckBoxEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkCheckBoxEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkCheckBoxEvents, Interfaces.IOlkCheckBoxEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkCheckBoxEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkCheckBoxEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkCheckBoxEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkCheckBoxEvents_Event, Interfaces.IOlkCheckBoxEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkCheckBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkCheckBox WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkCheckBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkCheckBox, Interfaces.IOlkCheckBox>();
		}

		/// <summary>
		/// Wrapper interface for _OlkOptionButton which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OlkOptionButton WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OlkOptionButton resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OlkOptionButton, Interfaces.I_OlkOptionButton>();
		}

		/// <summary>
		/// Wrapper interface for OlkOptionButtonEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkOptionButtonEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkOptionButtonEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkOptionButtonEvents, Interfaces.IOlkOptionButtonEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkOptionButtonEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkOptionButtonEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkOptionButtonEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkOptionButtonEvents_Event, Interfaces.IOlkOptionButtonEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkOptionButton which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkOptionButton WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkOptionButton resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkOptionButton, Interfaces.IOlkOptionButton>();
		}

		/// <summary>
		/// Wrapper interface for _OlkComboBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OlkComboBox WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OlkComboBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OlkComboBox, Interfaces.I_OlkComboBox>();
		}

		/// <summary>
		/// Wrapper interface for OlkComboBoxEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkComboBoxEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkComboBoxEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkComboBoxEvents, Interfaces.IOlkComboBoxEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkComboBoxEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkComboBoxEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkComboBoxEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkComboBoxEvents_Event, Interfaces.IOlkComboBoxEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkComboBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkComboBox WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkComboBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkComboBox, Interfaces.IOlkComboBox>();
		}

		/// <summary>
		/// Wrapper interface for _OlkListBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OlkListBox WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OlkListBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OlkListBox, Interfaces.I_OlkListBox>();
		}

		/// <summary>
		/// Wrapper interface for OlkListBoxEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkListBoxEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkListBoxEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkListBoxEvents, Interfaces.IOlkListBoxEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkListBoxEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkListBoxEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkListBoxEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkListBoxEvents_Event, Interfaces.IOlkListBoxEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkListBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkListBox WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkListBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkListBox, Interfaces.IOlkListBox>();
		}

		/// <summary>
		/// Wrapper interface for _OlkInfoBar which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OlkInfoBar WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OlkInfoBar resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OlkInfoBar, Interfaces.I_OlkInfoBar>();
		}

		/// <summary>
		/// Wrapper interface for OlkInfoBarEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkInfoBarEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkInfoBarEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkInfoBarEvents, Interfaces.IOlkInfoBarEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkInfoBarEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkInfoBarEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkInfoBarEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkInfoBarEvents_Event, Interfaces.IOlkInfoBarEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkInfoBar which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkInfoBar WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkInfoBar resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkInfoBar, Interfaces.IOlkInfoBar>();
		}

		/// <summary>
		/// Wrapper interface for _OlkContactPhoto which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OlkContactPhoto WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OlkContactPhoto resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OlkContactPhoto, Interfaces.I_OlkContactPhoto>();
		}

		/// <summary>
		/// Wrapper interface for OlkContactPhotoEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkContactPhotoEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkContactPhotoEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkContactPhotoEvents, Interfaces.IOlkContactPhotoEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkContactPhotoEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkContactPhotoEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkContactPhotoEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkContactPhotoEvents_Event, Interfaces.IOlkContactPhotoEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkContactPhoto which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkContactPhoto WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkContactPhoto resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkContactPhoto, Interfaces.IOlkContactPhoto>();
		}

		/// <summary>
		/// Wrapper interface for _OlkBusinessCardControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OlkBusinessCardControl WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OlkBusinessCardControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OlkBusinessCardControl, Interfaces.I_OlkBusinessCardControl>();
		}

		/// <summary>
		/// Wrapper interface for OlkBusinessCardControlEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkBusinessCardControlEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkBusinessCardControlEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkBusinessCardControlEvents, Interfaces.IOlkBusinessCardControlEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkBusinessCardControlEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkBusinessCardControlEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkBusinessCardControlEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkBusinessCardControlEvents_Event, Interfaces.IOlkBusinessCardControlEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkBusinessCardControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkBusinessCardControl WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkBusinessCardControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkBusinessCardControl, Interfaces.IOlkBusinessCardControl>();
		}

		/// <summary>
		/// Wrapper interface for _OlkPageControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OlkPageControl WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OlkPageControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OlkPageControl, Interfaces.I_OlkPageControl>();
		}

		/// <summary>
		/// Wrapper interface for OlkPageControlEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkPageControlEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkPageControlEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkPageControlEvents, Interfaces.IOlkPageControlEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkPageControlEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkPageControlEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkPageControlEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkPageControlEvents_Event, Interfaces.IOlkPageControlEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkPageControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkPageControl WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkPageControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkPageControl, Interfaces.IOlkPageControl>();
		}

		/// <summary>
		/// Wrapper interface for _OlkDateControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OlkDateControl WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OlkDateControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OlkDateControl, Interfaces.I_OlkDateControl>();
		}

		/// <summary>
		/// Wrapper interface for OlkDateControlEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkDateControlEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkDateControlEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkDateControlEvents, Interfaces.IOlkDateControlEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkDateControlEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkDateControlEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkDateControlEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkDateControlEvents_Event, Interfaces.IOlkDateControlEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkDateControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkDateControl WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkDateControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkDateControl, Interfaces.IOlkDateControl>();
		}

		/// <summary>
		/// Wrapper interface for _OlkTimeControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OlkTimeControl WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OlkTimeControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OlkTimeControl, Interfaces.I_OlkTimeControl>();
		}

		/// <summary>
		/// Wrapper interface for OlkTimeControlEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkTimeControlEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkTimeControlEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkTimeControlEvents, Interfaces.IOlkTimeControlEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkTimeControlEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkTimeControlEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkTimeControlEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkTimeControlEvents_Event, Interfaces.IOlkTimeControlEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkTimeControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkTimeControl WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkTimeControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkTimeControl, Interfaces.IOlkTimeControl>();
		}

		/// <summary>
		/// Wrapper interface for _OlkCategory which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OlkCategory WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OlkCategory resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OlkCategory, Interfaces.I_OlkCategory>();
		}

		/// <summary>
		/// Wrapper interface for OlkCategoryEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkCategoryEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkCategoryEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkCategoryEvents, Interfaces.IOlkCategoryEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkCategoryEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkCategoryEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkCategoryEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkCategoryEvents_Event, Interfaces.IOlkCategoryEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkCategory which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkCategory WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkCategory resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkCategory, Interfaces.IOlkCategory>();
		}

		/// <summary>
		/// Wrapper interface for _OlkFrameHeader which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OlkFrameHeader WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OlkFrameHeader resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OlkFrameHeader, Interfaces.I_OlkFrameHeader>();
		}

		/// <summary>
		/// Wrapper interface for OlkFrameHeaderEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkFrameHeaderEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkFrameHeaderEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkFrameHeaderEvents, Interfaces.IOlkFrameHeaderEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkFrameHeaderEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkFrameHeaderEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkFrameHeaderEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkFrameHeaderEvents_Event, Interfaces.IOlkFrameHeaderEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkFrameHeader which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkFrameHeader WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkFrameHeader resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkFrameHeader, Interfaces.IOlkFrameHeader>();
		}

		/// <summary>
		/// Wrapper interface for _OlkSenderPhoto which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OlkSenderPhoto WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OlkSenderPhoto resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OlkSenderPhoto, Interfaces.I_OlkSenderPhoto>();
		}

		/// <summary>
		/// Wrapper interface for OlkSenderPhotoEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkSenderPhotoEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkSenderPhotoEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkSenderPhotoEvents, Interfaces.IOlkSenderPhotoEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkSenderPhotoEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkSenderPhotoEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkSenderPhotoEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkSenderPhotoEvents_Event, Interfaces.IOlkSenderPhotoEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkSenderPhoto which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkSenderPhoto WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkSenderPhoto resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkSenderPhoto, Interfaces.IOlkSenderPhoto>();
		}

		/// <summary>
		/// Wrapper interface for _TimeZone which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_TimeZone WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._TimeZone resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._TimeZone, Interfaces.I_TimeZone>();
		}

		/// <summary>
		/// Wrapper interface for _Application which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Application WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Application resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Application, Interfaces.I_Application>();
		}

		/// <summary>
		/// Wrapper interface for _NameSpace which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_NameSpace WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._NameSpace resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._NameSpace, Interfaces.I_NameSpace>();
		}

		/// <summary>
		/// Wrapper interface for Recipient which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRecipient WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Recipient resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Recipient, Interfaces.IRecipient>();
		}

		/// <summary>
		/// Wrapper interface for AddressEntry which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAddressEntry WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.AddressEntry resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.AddressEntry, Interfaces.IAddressEntry>();
		}

		/// <summary>
		/// Wrapper interface for AddressEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAddressEntries WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.AddressEntries resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.AddressEntries, Interfaces.IAddressEntries>();
		}

		/// <summary>
		/// Wrapper interface for _ContactItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_ContactItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._ContactItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._ContactItem, Interfaces.I_ContactItem>();
		}

		/// <summary>
		/// Wrapper interface for Actions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IActions WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Actions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Actions, Interfaces.IActions>();
		}

		/// <summary>
		/// Wrapper interface for Action which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAction WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Action resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Action, Interfaces.IAction>();
		}

		/// <summary>
		/// Wrapper interface for Attachments which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAttachments WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Attachments resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Attachments, Interfaces.IAttachments>();
		}

		/// <summary>
		/// Wrapper interface for Attachment which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAttachment WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Attachment resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Attachment, Interfaces.IAttachment>();
		}

		/// <summary>
		/// Wrapper interface for PropertyAccessor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPropertyAccessor WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.PropertyAccessor resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.PropertyAccessor, Interfaces.IPropertyAccessor>();
		}

		/// <summary>
		/// Wrapper interface for _PropertyAccessor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_PropertyAccessor WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._PropertyAccessor resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._PropertyAccessor, Interfaces.I_PropertyAccessor>();
		}

		/// <summary>
		/// Wrapper interface for FormDescription which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFormDescription WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.FormDescription resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.FormDescription, Interfaces.IFormDescription>();
		}

		/// <summary>
		/// Wrapper interface for _Inspector which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Inspector WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Inspector resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Inspector, Interfaces.I_Inspector>();
		}

		/// <summary>
		/// Wrapper interface for _AttachmentSelection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_AttachmentSelection WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._AttachmentSelection resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._AttachmentSelection, Interfaces.I_AttachmentSelection>();
		}

		/// <summary>
		/// Wrapper interface for Selection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISelection WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Selection resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Selection, Interfaces.ISelection>();
		}

		/// <summary>
		/// Wrapper interface for UserProperties which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IUserProperties WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.UserProperties resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.UserProperties, Interfaces.IUserProperties>();
		}

		/// <summary>
		/// Wrapper interface for UserProperty which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IUserProperty WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.UserProperty resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.UserProperty, Interfaces.IUserProperty>();
		}

		/// <summary>
		/// Wrapper interface for MAPIFolder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMAPIFolder WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.MAPIFolder resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.MAPIFolder, Interfaces.IMAPIFolder>();
		}

		/// <summary>
		/// Wrapper interface for _Folders which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Folders WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Folders resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Folders, Interfaces.I_Folders>();
		}

		/// <summary>
		/// Wrapper interface for _Items which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Items WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Items resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Items, Interfaces.I_Items>();
		}

		/// <summary>
		/// Wrapper interface for _Explorer which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Explorer WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Explorer resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Explorer, Interfaces.I_Explorer>();
		}

		/// <summary>
		/// Wrapper interface for Panes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPanes WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Panes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Panes, Interfaces.IPanes>();
		}

		/// <summary>
		/// Wrapper interface for _NavigationPane which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_NavigationPane WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._NavigationPane resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._NavigationPane, Interfaces.I_NavigationPane>();
		}

		/// <summary>
		/// Wrapper interface for NavigationModule which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INavigationModule WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.NavigationModule resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.NavigationModule, Interfaces.INavigationModule>();
		}

		/// <summary>
		/// Wrapper interface for _NavigationModule which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_NavigationModule WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._NavigationModule resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._NavigationModule, Interfaces.I_NavigationModule>();
		}

		/// <summary>
		/// Wrapper interface for NavigationModules which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INavigationModules WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.NavigationModules resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.NavigationModules, Interfaces.INavigationModules>();
		}

		/// <summary>
		/// Wrapper interface for _NavigationModules which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_NavigationModules WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._NavigationModules resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._NavigationModules, Interfaces.I_NavigationModules>();
		}

		/// <summary>
		/// Wrapper interface for _AccountSelector which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_AccountSelector WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._AccountSelector resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._AccountSelector, Interfaces.I_AccountSelector>();
		}

		/// <summary>
		/// Wrapper interface for _Account which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Account WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Account resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Account, Interfaces.I_Account>();
		}

		/// <summary>
		/// Wrapper interface for Store which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IStore WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Store resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Store, Interfaces.IStore>();
		}

		/// <summary>
		/// Wrapper interface for _Store which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Store WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Store resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Store, Interfaces.I_Store>();
		}

		/// <summary>
		/// Wrapper interface for Rules which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRules WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Rules resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Rules, Interfaces.IRules>();
		}

		/// <summary>
		/// Wrapper interface for _Rules which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Rules WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Rules resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Rules, Interfaces.I_Rules>();
		}

		/// <summary>
		/// Wrapper interface for _Rule which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Rule WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Rule resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Rule, Interfaces.I_Rule>();
		}

		/// <summary>
		/// Wrapper interface for RuleActions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRuleActions WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.RuleActions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.RuleActions, Interfaces.IRuleActions>();
		}

		/// <summary>
		/// Wrapper interface for _RuleActions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_RuleActions WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._RuleActions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._RuleActions, Interfaces.I_RuleActions>();
		}

		/// <summary>
		/// Wrapper interface for _RuleAction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_RuleAction WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._RuleAction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._RuleAction, Interfaces.I_RuleAction>();
		}

		/// <summary>
		/// Wrapper interface for MoveOrCopyRuleAction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMoveOrCopyRuleAction WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.MoveOrCopyRuleAction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.MoveOrCopyRuleAction, Interfaces.IMoveOrCopyRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for _MoveOrCopyRuleAction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_MoveOrCopyRuleAction WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._MoveOrCopyRuleAction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._MoveOrCopyRuleAction, Interfaces.I_MoveOrCopyRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for RuleAction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRuleAction WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.RuleAction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.RuleAction, Interfaces.IRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for SendRuleAction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISendRuleAction WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.SendRuleAction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.SendRuleAction, Interfaces.ISendRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for _SendRuleAction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_SendRuleAction WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._SendRuleAction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._SendRuleAction, Interfaces.I_SendRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for Recipients which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRecipients WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Recipients resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Recipients, Interfaces.IRecipients>();
		}

		/// <summary>
		/// Wrapper interface for AssignToCategoryRuleAction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAssignToCategoryRuleAction WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.AssignToCategoryRuleAction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.AssignToCategoryRuleAction, Interfaces.IAssignToCategoryRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for _AssignToCategoryRuleAction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_AssignToCategoryRuleAction WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._AssignToCategoryRuleAction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._AssignToCategoryRuleAction, Interfaces.I_AssignToCategoryRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for PlaySoundRuleAction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPlaySoundRuleAction WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.PlaySoundRuleAction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.PlaySoundRuleAction, Interfaces.IPlaySoundRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for _PlaySoundRuleAction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_PlaySoundRuleAction WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._PlaySoundRuleAction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._PlaySoundRuleAction, Interfaces.I_PlaySoundRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for MarkAsTaskRuleAction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMarkAsTaskRuleAction WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.MarkAsTaskRuleAction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.MarkAsTaskRuleAction, Interfaces.IMarkAsTaskRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for _MarkAsTaskRuleAction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_MarkAsTaskRuleAction WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._MarkAsTaskRuleAction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._MarkAsTaskRuleAction, Interfaces.I_MarkAsTaskRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for NewItemAlertRuleAction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INewItemAlertRuleAction WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.NewItemAlertRuleAction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.NewItemAlertRuleAction, Interfaces.INewItemAlertRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for _NewItemAlertRuleAction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_NewItemAlertRuleAction WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._NewItemAlertRuleAction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._NewItemAlertRuleAction, Interfaces.I_NewItemAlertRuleAction>();
		}

		/// <summary>
		/// Wrapper interface for RuleConditions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRuleConditions WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.RuleConditions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.RuleConditions, Interfaces.IRuleConditions>();
		}

		/// <summary>
		/// Wrapper interface for _RuleConditions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_RuleConditions WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._RuleConditions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._RuleConditions, Interfaces.I_RuleConditions>();
		}

		/// <summary>
		/// Wrapper interface for _RuleCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_RuleCondition WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._RuleCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._RuleCondition, Interfaces.I_RuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for RuleCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRuleCondition WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.RuleCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.RuleCondition, Interfaces.IRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for ImportanceRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IImportanceRuleCondition WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ImportanceRuleCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ImportanceRuleCondition, Interfaces.IImportanceRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for _ImportanceRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_ImportanceRuleCondition WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._ImportanceRuleCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._ImportanceRuleCondition, Interfaces.I_ImportanceRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for AccountRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAccountRuleCondition WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.AccountRuleCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.AccountRuleCondition, Interfaces.IAccountRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for _AccountRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_AccountRuleCondition WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._AccountRuleCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._AccountRuleCondition, Interfaces.I_AccountRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for Account which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAccount WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Account resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Account, Interfaces.IAccount>();
		}

		/// <summary>
		/// Wrapper interface for TextRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextRuleCondition WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.TextRuleCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.TextRuleCondition, Interfaces.ITextRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for _TextRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_TextRuleCondition WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._TextRuleCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._TextRuleCondition, Interfaces.I_TextRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for CategoryRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICategoryRuleCondition WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.CategoryRuleCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.CategoryRuleCondition, Interfaces.ICategoryRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for _CategoryRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CategoryRuleCondition WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._CategoryRuleCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._CategoryRuleCondition, Interfaces.I_CategoryRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for FormNameRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFormNameRuleCondition WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.FormNameRuleCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.FormNameRuleCondition, Interfaces.IFormNameRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for _FormNameRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_FormNameRuleCondition WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._FormNameRuleCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._FormNameRuleCondition, Interfaces.I_FormNameRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for ToOrFromRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IToOrFromRuleCondition WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ToOrFromRuleCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ToOrFromRuleCondition, Interfaces.IToOrFromRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for _ToOrFromRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_ToOrFromRuleCondition WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._ToOrFromRuleCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._ToOrFromRuleCondition, Interfaces.I_ToOrFromRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for AddressRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAddressRuleCondition WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.AddressRuleCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.AddressRuleCondition, Interfaces.IAddressRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for _AddressRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_AddressRuleCondition WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._AddressRuleCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._AddressRuleCondition, Interfaces.I_AddressRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for SenderInAddressListRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISenderInAddressListRuleCondition WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.SenderInAddressListRuleCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.SenderInAddressListRuleCondition, Interfaces.ISenderInAddressListRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for _SenderInAddressListRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_SenderInAddressListRuleCondition WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._SenderInAddressListRuleCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._SenderInAddressListRuleCondition, Interfaces.I_SenderInAddressListRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for AddressList which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAddressList WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.AddressList resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.AddressList, Interfaces.IAddressList>();
		}

		/// <summary>
		/// Wrapper interface for FromRssFeedRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFromRssFeedRuleCondition WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.FromRssFeedRuleCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.FromRssFeedRuleCondition, Interfaces.IFromRssFeedRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for _FromRssFeedRuleCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_FromRssFeedRuleCondition WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._FromRssFeedRuleCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._FromRssFeedRuleCondition, Interfaces.I_FromRssFeedRuleCondition>();
		}

		/// <summary>
		/// Wrapper interface for Rule which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRule WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Rule resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Rule, Interfaces.IRule>();
		}

		/// <summary>
		/// Wrapper interface for Categories which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICategories WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Categories resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Categories, Interfaces.ICategories>();
		}

		/// <summary>
		/// Wrapper interface for _Categories which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Categories WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Categories resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Categories, Interfaces.I_Categories>();
		}

		/// <summary>
		/// Wrapper interface for _Category which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Category WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Category resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Category, Interfaces.I_Category>();
		}

		/// <summary>
		/// Wrapper interface for Category which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICategory WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Category resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Category, Interfaces.ICategory>();
		}

		/// <summary>
		/// Wrapper interface for View which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IView WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.View resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.View, Interfaces.IView>();
		}

		/// <summary>
		/// Wrapper interface for _Views which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Views WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Views resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Views, Interfaces.I_Views>();
		}

		/// <summary>
		/// Wrapper interface for _StorageItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_StorageItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._StorageItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._StorageItem, Interfaces.I_StorageItem>();
		}

		/// <summary>
		/// Wrapper interface for Table which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITable WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Table resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Table, Interfaces.ITable>();
		}

		/// <summary>
		/// Wrapper interface for _Table which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Table WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Table resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Table, Interfaces.I_Table>();
		}

		/// <summary>
		/// Wrapper interface for Row which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRow WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Row resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Row, Interfaces.IRow>();
		}

		/// <summary>
		/// Wrapper interface for _Row which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Row WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Row resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Row, Interfaces.I_Row>();
		}

		/// <summary>
		/// Wrapper interface for Columns which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColumns WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Columns resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Columns, Interfaces.IColumns>();
		}

		/// <summary>
		/// Wrapper interface for _Columns which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Columns WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Columns resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Columns, Interfaces.I_Columns>();
		}

		/// <summary>
		/// Wrapper interface for _Column which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Column WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Column resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Column, Interfaces.I_Column>();
		}

		/// <summary>
		/// Wrapper interface for Column which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColumn WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Column resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Column, Interfaces.IColumn>();
		}

		/// <summary>
		/// Wrapper interface for CalendarSharing which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICalendarSharing WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.CalendarSharing resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.CalendarSharing, Interfaces.ICalendarSharing>();
		}

		/// <summary>
		/// Wrapper interface for _CalendarSharing which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CalendarSharing WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._CalendarSharing resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._CalendarSharing, Interfaces.I_CalendarSharing>();
		}

		/// <summary>
		/// Wrapper interface for ItemEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IItemEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ItemEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ItemEvents_Event, Interfaces.IItemEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for ItemEvents_10_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IItemEvents_10_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ItemEvents_10_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ItemEvents_10_Event, Interfaces.IItemEvents_10_Event>();
		}

		/// <summary>
		/// Wrapper interface for MailItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.MailItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.MailItem, Interfaces.IMailItem>();
		}

		/// <summary>
		/// Wrapper interface for _MailItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_MailItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._MailItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._MailItem, Interfaces.I_MailItem>();
		}

		/// <summary>
		/// Wrapper interface for Links which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILinks WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Links resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Links, Interfaces.ILinks>();
		}

		/// <summary>
		/// Wrapper interface for Link which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILink WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Link resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Link, Interfaces.ILink>();
		}

		/// <summary>
		/// Wrapper interface for ItemProperties which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IItemProperties WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ItemProperties resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ItemProperties, Interfaces.IItemProperties>();
		}

		/// <summary>
		/// Wrapper interface for ItemProperty which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IItemProperty WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ItemProperty resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ItemProperty, Interfaces.IItemProperty>();
		}

		/// <summary>
		/// Wrapper interface for Conflicts which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IConflicts WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Conflicts resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Conflicts, Interfaces.IConflicts>();
		}

		/// <summary>
		/// Wrapper interface for Conflict which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IConflict WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Conflict resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Conflict, Interfaces.IConflict>();
		}

		/// <summary>
		/// Wrapper interface for ContactItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IContactItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ContactItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ContactItem, Interfaces.IContactItem>();
		}

		/// <summary>
		/// Wrapper interface for ItemEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IItemEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ItemEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ItemEvents, Interfaces.IItemEvents>();
		}

		/// <summary>
		/// Wrapper interface for ItemEvents_10 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IItemEvents_10 WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ItemEvents_10 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ItemEvents_10, Interfaces.IItemEvents_10>();
		}

		/// <summary>
		/// Wrapper interface for _Conversation which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Conversation WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Conversation resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Conversation, Interfaces.I_Conversation>();
		}

		/// <summary>
		/// Wrapper interface for SimpleItems which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISimpleItems WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.SimpleItems resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.SimpleItems, Interfaces.ISimpleItems>();
		}

		/// <summary>
		/// Wrapper interface for _SimpleItems which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_SimpleItems WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._SimpleItems resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._SimpleItems, Interfaces.I_SimpleItems>();
		}

		/// <summary>
		/// Wrapper interface for UserDefinedProperties which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IUserDefinedProperties WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.UserDefinedProperties resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.UserDefinedProperties, Interfaces.IUserDefinedProperties>();
		}

		/// <summary>
		/// Wrapper interface for _UserDefinedProperties which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_UserDefinedProperties WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._UserDefinedProperties resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._UserDefinedProperties, Interfaces.I_UserDefinedProperties>();
		}

		/// <summary>
		/// Wrapper interface for _UserDefinedProperty which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_UserDefinedProperty WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._UserDefinedProperty resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._UserDefinedProperty, Interfaces.I_UserDefinedProperty>();
		}

		/// <summary>
		/// Wrapper interface for UserDefinedProperty which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IUserDefinedProperty WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.UserDefinedProperty resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.UserDefinedProperty, Interfaces.IUserDefinedProperty>();
		}

		/// <summary>
		/// Wrapper interface for ExchangeUser which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IExchangeUser WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ExchangeUser resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ExchangeUser, Interfaces.IExchangeUser>();
		}

		/// <summary>
		/// Wrapper interface for _ExchangeUser which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_ExchangeUser WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._ExchangeUser resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._ExchangeUser, Interfaces.I_ExchangeUser>();
		}

		/// <summary>
		/// Wrapper interface for ExchangeDistributionList which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IExchangeDistributionList WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ExchangeDistributionList resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ExchangeDistributionList, Interfaces.IExchangeDistributionList>();
		}

		/// <summary>
		/// Wrapper interface for _ExchangeDistributionList which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_ExchangeDistributionList WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._ExchangeDistributionList resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._ExchangeDistributionList, Interfaces.I_ExchangeDistributionList>();
		}

		/// <summary>
		/// Wrapper interface for AddressLists which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAddressLists WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.AddressLists resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.AddressLists, Interfaces.IAddressLists>();
		}

		/// <summary>
		/// Wrapper interface for SyncObjects which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISyncObjects WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.SyncObjects resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.SyncObjects, Interfaces.ISyncObjects>();
		}

		/// <summary>
		/// Wrapper interface for SyncObjectEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISyncObjectEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.SyncObjectEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.SyncObjectEvents_Event, Interfaces.ISyncObjectEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for SyncObject which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISyncObject WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.SyncObject resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.SyncObject, Interfaces.ISyncObject>();
		}

		/// <summary>
		/// Wrapper interface for _SyncObject which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_SyncObject WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._SyncObject resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._SyncObject, Interfaces.I_SyncObject>();
		}

		/// <summary>
		/// Wrapper interface for SyncObjectEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISyncObjectEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.SyncObjectEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.SyncObjectEvents, Interfaces.ISyncObjectEvents>();
		}

		/// <summary>
		/// Wrapper interface for AccountsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAccountsEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.AccountsEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.AccountsEvents_Event, Interfaces.IAccountsEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Accounts which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAccounts WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Accounts resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Accounts, Interfaces.IAccounts>();
		}

		/// <summary>
		/// Wrapper interface for _Accounts which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Accounts WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Accounts resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Accounts, Interfaces.I_Accounts>();
		}

		/// <summary>
		/// Wrapper interface for AccountsEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAccountsEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.AccountsEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.AccountsEvents, Interfaces.IAccountsEvents>();
		}

		/// <summary>
		/// Wrapper interface for StoresEvents_12_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IStoresEvents_12_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.StoresEvents_12_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.StoresEvents_12_Event, Interfaces.IStoresEvents_12_Event>();
		}

		/// <summary>
		/// Wrapper interface for Stores which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IStores WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Stores resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Stores, Interfaces.IStores>();
		}

		/// <summary>
		/// Wrapper interface for _Stores which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Stores WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Stores resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Stores, Interfaces.I_Stores>();
		}

		/// <summary>
		/// Wrapper interface for StoresEvents_12 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IStoresEvents_12 WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.StoresEvents_12 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.StoresEvents_12, Interfaces.IStoresEvents_12>();
		}

		/// <summary>
		/// Wrapper interface for SelectNamesDialog which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISelectNamesDialog WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.SelectNamesDialog resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.SelectNamesDialog, Interfaces.ISelectNamesDialog>();
		}

		/// <summary>
		/// Wrapper interface for _SelectNamesDialog which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_SelectNamesDialog WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._SelectNamesDialog resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._SelectNamesDialog, Interfaces.I_SelectNamesDialog>();
		}

		/// <summary>
		/// Wrapper interface for SharingItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISharingItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.SharingItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.SharingItem, Interfaces.ISharingItem>();
		}

		/// <summary>
		/// Wrapper interface for _SharingItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_SharingItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._SharingItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._SharingItem, Interfaces.I_SharingItem>();
		}

		/// <summary>
		/// Wrapper interface for _Explorers which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Explorers WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Explorers resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Explorers, Interfaces.I_Explorers>();
		}

		/// <summary>
		/// Wrapper interface for ExplorerEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IExplorerEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ExplorerEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ExplorerEvents_Event, Interfaces.IExplorerEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for ExplorerEvents_10_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IExplorerEvents_10_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ExplorerEvents_10_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ExplorerEvents_10_Event, Interfaces.IExplorerEvents_10_Event>();
		}

		/// <summary>
		/// Wrapper interface for Explorer which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IExplorer WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Explorer resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Explorer, Interfaces.IExplorer>();
		}

		/// <summary>
		/// Wrapper interface for ExplorerEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IExplorerEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ExplorerEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ExplorerEvents, Interfaces.IExplorerEvents>();
		}

		/// <summary>
		/// Wrapper interface for ExplorerEvents_10 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IExplorerEvents_10 WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ExplorerEvents_10 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ExplorerEvents_10, Interfaces.IExplorerEvents_10>();
		}

		/// <summary>
		/// Wrapper interface for _Inspectors which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Inspectors WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Inspectors resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Inspectors, Interfaces.I_Inspectors>();
		}

		/// <summary>
		/// Wrapper interface for InspectorEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IInspectorEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.InspectorEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.InspectorEvents_Event, Interfaces.IInspectorEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for InspectorEvents_10_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IInspectorEvents_10_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.InspectorEvents_10_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.InspectorEvents_10_Event, Interfaces.IInspectorEvents_10_Event>();
		}

		/// <summary>
		/// Wrapper interface for Inspector which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IInspector WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Inspector resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Inspector, Interfaces.IInspector>();
		}

		/// <summary>
		/// Wrapper interface for InspectorEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IInspectorEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.InspectorEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.InspectorEvents, Interfaces.IInspectorEvents>();
		}

		/// <summary>
		/// Wrapper interface for InspectorEvents_10 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IInspectorEvents_10 WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.InspectorEvents_10 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.InspectorEvents_10, Interfaces.IInspectorEvents_10>();
		}

		/// <summary>
		/// Wrapper interface for Search which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISearch WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Search resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Search, Interfaces.ISearch>();
		}

		/// <summary>
		/// Wrapper interface for _Results which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Results WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Results resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Results, Interfaces.I_Results>();
		}

		/// <summary>
		/// Wrapper interface for _Reminders which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Reminders WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Reminders resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Reminders, Interfaces.I_Reminders>();
		}

		/// <summary>
		/// Wrapper interface for _Reminder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Reminder WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._Reminder resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._Reminder, Interfaces.I_Reminder>();
		}

		/// <summary>
		/// Wrapper interface for TimeZones which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITimeZones WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.TimeZones resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.TimeZones, Interfaces.ITimeZones>();
		}

		/// <summary>
		/// Wrapper interface for _TimeZones which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_TimeZones WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._TimeZones resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._TimeZones, Interfaces.I_TimeZones>();
		}

		/// <summary>
		/// Wrapper interface for _OlkTimeZoneControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OlkTimeZoneControl WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OlkTimeZoneControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OlkTimeZoneControl, Interfaces.I_OlkTimeZoneControl>();
		}

		/// <summary>
		/// Wrapper interface for OlkTimeZoneControlEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkTimeZoneControlEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkTimeZoneControlEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkTimeZoneControlEvents, Interfaces.IOlkTimeZoneControlEvents>();
		}

		/// <summary>
		/// Wrapper interface for OlkTimeZoneControlEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkTimeZoneControlEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkTimeZoneControlEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkTimeZoneControlEvents_Event, Interfaces.IOlkTimeZoneControlEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OlkTimeZoneControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOlkTimeZoneControl WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OlkTimeZoneControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OlkTimeZoneControl, Interfaces.IOlkTimeZoneControl>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ApplicationEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ApplicationEvents, Interfaces.IApplicationEvents>();
		}

		/// <summary>
		/// Wrapper interface for PropertyPages which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPropertyPages WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.PropertyPages resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.PropertyPages, Interfaces.IPropertyPages>();
		}

		/// <summary>
		/// Wrapper interface for RecurrencePattern which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRecurrencePattern WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.RecurrencePattern resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.RecurrencePattern, Interfaces.IRecurrencePattern>();
		}

		/// <summary>
		/// Wrapper interface for Exceptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IExceptions WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Exceptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Exceptions, Interfaces.IExceptions>();
		}

		/// <summary>
		/// Wrapper interface for Exception which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IException WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Exception resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Exception, Interfaces.IException>();
		}

		/// <summary>
		/// Wrapper interface for AppointmentItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAppointmentItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.AppointmentItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.AppointmentItem, Interfaces.IAppointmentItem>();
		}

		/// <summary>
		/// Wrapper interface for _AppointmentItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_AppointmentItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._AppointmentItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._AppointmentItem, Interfaces.I_AppointmentItem>();
		}

		/// <summary>
		/// Wrapper interface for MeetingItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMeetingItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.MeetingItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.MeetingItem, Interfaces.IMeetingItem>();
		}

		/// <summary>
		/// Wrapper interface for _MeetingItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_MeetingItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._MeetingItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._MeetingItem, Interfaces.I_MeetingItem>();
		}

		/// <summary>
		/// Wrapper interface for ExplorersEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IExplorersEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ExplorersEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ExplorersEvents, Interfaces.IExplorersEvents>();
		}

		/// <summary>
		/// Wrapper interface for FoldersEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFoldersEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.FoldersEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.FoldersEvents, Interfaces.IFoldersEvents>();
		}

		/// <summary>
		/// Wrapper interface for InspectorsEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IInspectorsEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.InspectorsEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.InspectorsEvents, Interfaces.IInspectorsEvents>();
		}

		/// <summary>
		/// Wrapper interface for ItemsEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IItemsEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ItemsEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ItemsEvents, Interfaces.IItemsEvents>();
		}

		/// <summary>
		/// Wrapper interface for NameSpaceEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INameSpaceEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.NameSpaceEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.NameSpaceEvents, Interfaces.INameSpaceEvents>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarGroup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOutlookBarGroup WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OutlookBarGroup resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OutlookBarGroup, Interfaces.IOutlookBarGroup>();
		}

		/// <summary>
		/// Wrapper interface for _OutlookBarShortcuts which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OutlookBarShortcuts WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OutlookBarShortcuts resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OutlookBarShortcuts, Interfaces.I_OutlookBarShortcuts>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarShortcut which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOutlookBarShortcut WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OutlookBarShortcut resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OutlookBarShortcut, Interfaces.IOutlookBarShortcut>();
		}

		/// <summary>
		/// Wrapper interface for _OutlookBarGroups which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OutlookBarGroups WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OutlookBarGroups resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OutlookBarGroups, Interfaces.I_OutlookBarGroups>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarGroupsEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOutlookBarGroupsEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OutlookBarGroupsEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OutlookBarGroupsEvents, Interfaces.IOutlookBarGroupsEvents>();
		}

		/// <summary>
		/// Wrapper interface for _OutlookBarPane which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OutlookBarPane WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OutlookBarPane resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OutlookBarPane, Interfaces.I_OutlookBarPane>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarStorage which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOutlookBarStorage WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OutlookBarStorage resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OutlookBarStorage, Interfaces.IOutlookBarStorage>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarPaneEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOutlookBarPaneEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OutlookBarPaneEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OutlookBarPaneEvents, Interfaces.IOutlookBarPaneEvents>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarShortcutsEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOutlookBarShortcutsEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OutlookBarShortcutsEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OutlookBarShortcutsEvents, Interfaces.IOutlookBarShortcutsEvents>();
		}

		/// <summary>
		/// Wrapper interface for PropertyPage which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPropertyPage WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.PropertyPage resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.PropertyPage, Interfaces.IPropertyPage>();
		}

		/// <summary>
		/// Wrapper interface for PropertyPageSite which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPropertyPageSite WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.PropertyPageSite resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.PropertyPageSite, Interfaces.IPropertyPageSite>();
		}

		/// <summary>
		/// Wrapper interface for Pages which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPages WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Pages resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Pages, Interfaces.IPages>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents_10 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents_10 WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ApplicationEvents_10 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ApplicationEvents_10, Interfaces.IApplicationEvents_10>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents_11 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents_11 WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ApplicationEvents_11 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ApplicationEvents_11, Interfaces.IApplicationEvents_11>();
		}

		/// <summary>
		/// Wrapper interface for AttachmentSelection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAttachmentSelection WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.AttachmentSelection resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.AttachmentSelection, Interfaces.IAttachmentSelection>();
		}

		/// <summary>
		/// Wrapper interface for MAPIFolderEvents_12_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMAPIFolderEvents_12_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.MAPIFolderEvents_12_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.MAPIFolderEvents_12_Event, Interfaces.IMAPIFolderEvents_12_Event>();
		}

		/// <summary>
		/// Wrapper interface for Folder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFolder WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Folder resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Folder, Interfaces.IFolder>();
		}

		/// <summary>
		/// Wrapper interface for MAPIFolderEvents_12 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMAPIFolderEvents_12 WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.MAPIFolderEvents_12 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.MAPIFolderEvents_12, Interfaces.IMAPIFolderEvents_12>();
		}

		/// <summary>
		/// Wrapper interface for ResultsEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IResultsEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ResultsEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ResultsEvents, Interfaces.IResultsEvents>();
		}

		/// <summary>
		/// Wrapper interface for _ViewsEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_ViewsEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._ViewsEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._ViewsEvents, Interfaces.I_ViewsEvents>();
		}

		/// <summary>
		/// Wrapper interface for ReminderCollectionEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IReminderCollectionEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ReminderCollectionEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ReminderCollectionEvents, Interfaces.IReminderCollectionEvents>();
		}

		/// <summary>
		/// Wrapper interface for _DocumentItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_DocumentItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._DocumentItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._DocumentItem, Interfaces.I_DocumentItem>();
		}

		/// <summary>
		/// Wrapper interface for _NoteItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_NoteItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._NoteItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._NoteItem, Interfaces.I_NoteItem>();
		}

		/// <summary>
		/// Wrapper interface for FormRegionEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFormRegionEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.FormRegionEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.FormRegionEvents, Interfaces.IFormRegionEvents>();
		}

		/// <summary>
		/// Wrapper interface for _ViewField which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_ViewField WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._ViewField resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._ViewField, Interfaces.I_ViewField>();
		}

		/// <summary>
		/// Wrapper interface for ColumnFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColumnFormat WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ColumnFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ColumnFormat, Interfaces.IColumnFormat>();
		}

		/// <summary>
		/// Wrapper interface for _ColumnFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_ColumnFormat WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._ColumnFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._ColumnFormat, Interfaces.I_ColumnFormat>();
		}

		/// <summary>
		/// Wrapper interface for _ViewFields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_ViewFields WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._ViewFields resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._ViewFields, Interfaces.I_ViewFields>();
		}

		/// <summary>
		/// Wrapper interface for ViewField which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IViewField WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ViewField resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ViewField, Interfaces.IViewField>();
		}

		/// <summary>
		/// Wrapper interface for _IconView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_IconView WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._IconView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._IconView, Interfaces.I_IconView>();
		}

		/// <summary>
		/// Wrapper interface for OrderFields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOrderFields WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OrderFields resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OrderFields, Interfaces.IOrderFields>();
		}

		/// <summary>
		/// Wrapper interface for _OrderFields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OrderFields WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OrderFields resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OrderFields, Interfaces.I_OrderFields>();
		}

		/// <summary>
		/// Wrapper interface for _OrderField which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OrderField WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._OrderField resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._OrderField, Interfaces.I_OrderField>();
		}

		/// <summary>
		/// Wrapper interface for OrderField which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOrderField WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OrderField resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OrderField, Interfaces.IOrderField>();
		}

		/// <summary>
		/// Wrapper interface for _CardView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CardView WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._CardView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._CardView, Interfaces.I_CardView>();
		}

		/// <summary>
		/// Wrapper interface for ViewFields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IViewFields WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ViewFields resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ViewFields, Interfaces.IViewFields>();
		}

		/// <summary>
		/// Wrapper interface for ViewFont which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IViewFont WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ViewFont resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ViewFont, Interfaces.IViewFont>();
		}

		/// <summary>
		/// Wrapper interface for _ViewFont which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_ViewFont WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._ViewFont resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._ViewFont, Interfaces.I_ViewFont>();
		}

		/// <summary>
		/// Wrapper interface for AutoFormatRules which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAutoFormatRules WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.AutoFormatRules resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.AutoFormatRules, Interfaces.IAutoFormatRules>();
		}

		/// <summary>
		/// Wrapper interface for _AutoFormatRules which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_AutoFormatRules WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._AutoFormatRules resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._AutoFormatRules, Interfaces.I_AutoFormatRules>();
		}

		/// <summary>
		/// Wrapper interface for AutoFormatRule which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAutoFormatRule WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.AutoFormatRule resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.AutoFormatRule, Interfaces.IAutoFormatRule>();
		}

		/// <summary>
		/// Wrapper interface for _AutoFormatRule which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_AutoFormatRule WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._AutoFormatRule resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._AutoFormatRule, Interfaces.I_AutoFormatRule>();
		}

		/// <summary>
		/// Wrapper interface for _TimelineView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_TimelineView WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._TimelineView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._TimelineView, Interfaces.I_TimelineView>();
		}

		/// <summary>
		/// Wrapper interface for _MailModule which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_MailModule WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._MailModule resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._MailModule, Interfaces.I_MailModule>();
		}

		/// <summary>
		/// Wrapper interface for _NavigationGroups which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_NavigationGroups WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._NavigationGroups resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._NavigationGroups, Interfaces.I_NavigationGroups>();
		}

		/// <summary>
		/// Wrapper interface for _NavigationGroup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_NavigationGroup WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._NavigationGroup resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._NavigationGroup, Interfaces.I_NavigationGroup>();
		}

		/// <summary>
		/// Wrapper interface for NavigationFolders which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INavigationFolders WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.NavigationFolders resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.NavigationFolders, Interfaces.INavigationFolders>();
		}

		/// <summary>
		/// Wrapper interface for _NavigationFolders which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_NavigationFolders WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._NavigationFolders resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._NavigationFolders, Interfaces.I_NavigationFolders>();
		}

		/// <summary>
		/// Wrapper interface for _NavigationFolder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_NavigationFolder WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._NavigationFolder resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._NavigationFolder, Interfaces.I_NavigationFolder>();
		}

		/// <summary>
		/// Wrapper interface for NavigationFolder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INavigationFolder WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.NavigationFolder resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.NavigationFolder, Interfaces.INavigationFolder>();
		}

		/// <summary>
		/// Wrapper interface for NavigationGroup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INavigationGroup WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.NavigationGroup resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.NavigationGroup, Interfaces.INavigationGroup>();
		}

		/// <summary>
		/// Wrapper interface for _CalendarModule which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CalendarModule WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._CalendarModule resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._CalendarModule, Interfaces.I_CalendarModule>();
		}

		/// <summary>
		/// Wrapper interface for _ContactsModule which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_ContactsModule WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._ContactsModule resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._ContactsModule, Interfaces.I_ContactsModule>();
		}

		/// <summary>
		/// Wrapper interface for _TasksModule which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_TasksModule WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._TasksModule resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._TasksModule, Interfaces.I_TasksModule>();
		}

		/// <summary>
		/// Wrapper interface for _JournalModule which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_JournalModule WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._JournalModule resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._JournalModule, Interfaces.I_JournalModule>();
		}

		/// <summary>
		/// Wrapper interface for _NotesModule which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_NotesModule WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._NotesModule resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._NotesModule, Interfaces.I_NotesModule>();
		}

		/// <summary>
		/// Wrapper interface for NavigationPaneEvents_12 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INavigationPaneEvents_12 WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.NavigationPaneEvents_12 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.NavigationPaneEvents_12, Interfaces.INavigationPaneEvents_12>();
		}

		/// <summary>
		/// Wrapper interface for NavigationGroupsEvents_12 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INavigationGroupsEvents_12 WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.NavigationGroupsEvents_12 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.NavigationGroupsEvents_12, Interfaces.INavigationGroupsEvents_12>();
		}

		/// <summary>
		/// Wrapper interface for _BusinessCardView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_BusinessCardView WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._BusinessCardView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._BusinessCardView, Interfaces.I_BusinessCardView>();
		}

		/// <summary>
		/// Wrapper interface for _FormRegionStartup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_FormRegionStartup WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._FormRegionStartup resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._FormRegionStartup, Interfaces.I_FormRegionStartup>();
		}

		/// <summary>
		/// Wrapper interface for FormRegionEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFormRegionEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.FormRegionEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.FormRegionEvents_Event, Interfaces.IFormRegionEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for FormRegion which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFormRegion WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.FormRegion resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.FormRegion, Interfaces.IFormRegion>();
		}

		/// <summary>
		/// Wrapper interface for _FormRegion which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_FormRegion WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._FormRegion resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._FormRegion, Interfaces.I_FormRegion>();
		}

		/// <summary>
		/// Wrapper interface for _SolutionsModule which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_SolutionsModule WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._SolutionsModule resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._SolutionsModule, Interfaces.I_SolutionsModule>();
		}

		/// <summary>
		/// Wrapper interface for _CalendarView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CalendarView WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._CalendarView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._CalendarView, Interfaces.I_CalendarView>();
		}

		/// <summary>
		/// Wrapper interface for _TableView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_TableView WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._TableView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._TableView, Interfaces.I_TableView>();
		}

		/// <summary>
		/// Wrapper interface for _MobileItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_MobileItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._MobileItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._MobileItem, Interfaces.I_MobileItem>();
		}

		/// <summary>
		/// Wrapper interface for MobileItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMobileItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.MobileItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.MobileItem, Interfaces.IMobileItem>();
		}

		/// <summary>
		/// Wrapper interface for _JournalItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_JournalItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._JournalItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._JournalItem, Interfaces.I_JournalItem>();
		}

		/// <summary>
		/// Wrapper interface for _PostItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_PostItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._PostItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._PostItem, Interfaces.I_PostItem>();
		}

		/// <summary>
		/// Wrapper interface for _TaskItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_TaskItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._TaskItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._TaskItem, Interfaces.I_TaskItem>();
		}

		/// <summary>
		/// Wrapper interface for TaskItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITaskItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.TaskItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.TaskItem, Interfaces.ITaskItem>();
		}

		/// <summary>
		/// Wrapper interface for AccountSelectorEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAccountSelectorEvents WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.AccountSelectorEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.AccountSelectorEvents, Interfaces.IAccountSelectorEvents>();
		}

		/// <summary>
		/// Wrapper interface for _DistListItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_DistListItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._DistListItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._DistListItem, Interfaces.I_DistListItem>();
		}

		/// <summary>
		/// Wrapper interface for _ReportItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_ReportItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._ReportItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._ReportItem, Interfaces.I_ReportItem>();
		}

		/// <summary>
		/// Wrapper interface for _RemoteItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_RemoteItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._RemoteItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._RemoteItem, Interfaces.I_RemoteItem>();
		}

		/// <summary>
		/// Wrapper interface for _TaskRequestItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_TaskRequestItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._TaskRequestItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._TaskRequestItem, Interfaces.I_TaskRequestItem>();
		}

		/// <summary>
		/// Wrapper interface for _TaskRequestAcceptItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_TaskRequestAcceptItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._TaskRequestAcceptItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._TaskRequestAcceptItem, Interfaces.I_TaskRequestAcceptItem>();
		}

		/// <summary>
		/// Wrapper interface for _TaskRequestDeclineItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_TaskRequestDeclineItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._TaskRequestDeclineItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._TaskRequestDeclineItem, Interfaces.I_TaskRequestDeclineItem>();
		}

		/// <summary>
		/// Wrapper interface for _TaskRequestUpdateItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_TaskRequestUpdateItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._TaskRequestUpdateItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._TaskRequestUpdateItem, Interfaces.I_TaskRequestUpdateItem>();
		}

		/// <summary>
		/// Wrapper interface for _ConversationHeader which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_ConversationHeader WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._ConversationHeader resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._ConversationHeader, Interfaces.I_ConversationHeader>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ApplicationEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ApplicationEvents_Event, Interfaces.IApplicationEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents_10_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents_10_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ApplicationEvents_10_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ApplicationEvents_10_Event, Interfaces.IApplicationEvents_10_Event>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents_11_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents_11_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ApplicationEvents_11_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ApplicationEvents_11_Event, Interfaces.IApplicationEvents_11_Event>();
		}

		/// <summary>
		/// Wrapper interface for Application which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplication WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Application resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Application, Interfaces.IApplication>();
		}

		/// <summary>
		/// Wrapper interface for DistListItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDistListItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.DistListItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.DistListItem, Interfaces.IDistListItem>();
		}

		/// <summary>
		/// Wrapper interface for DocumentItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocumentItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.DocumentItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.DocumentItem, Interfaces.IDocumentItem>();
		}

		/// <summary>
		/// Wrapper interface for ExplorersEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IExplorersEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ExplorersEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ExplorersEvents_Event, Interfaces.IExplorersEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Explorers which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IExplorers WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Explorers resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Explorers, Interfaces.IExplorers>();
		}

		/// <summary>
		/// Wrapper interface for InspectorsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IInspectorsEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.InspectorsEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.InspectorsEvents_Event, Interfaces.IInspectorsEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Inspectors which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IInspectors WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Inspectors resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Inspectors, Interfaces.IInspectors>();
		}

		/// <summary>
		/// Wrapper interface for FoldersEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFoldersEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.FoldersEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.FoldersEvents_Event, Interfaces.IFoldersEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Folders which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFolders WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Folders resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Folders, Interfaces.IFolders>();
		}

		/// <summary>
		/// Wrapper interface for ItemsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IItemsEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ItemsEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ItemsEvents_Event, Interfaces.IItemsEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Items which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IItems WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Items resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Items, Interfaces.IItems>();
		}

		/// <summary>
		/// Wrapper interface for JournalItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IJournalItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.JournalItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.JournalItem, Interfaces.IJournalItem>();
		}

		/// <summary>
		/// Wrapper interface for NameSpaceEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INameSpaceEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.NameSpaceEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.NameSpaceEvents_Event, Interfaces.INameSpaceEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for NameSpace which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INameSpace WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.NameSpace resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.NameSpace, Interfaces.INameSpace>();
		}

		/// <summary>
		/// Wrapper interface for NoteItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INoteItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.NoteItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.NoteItem, Interfaces.INoteItem>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarGroupsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOutlookBarGroupsEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OutlookBarGroupsEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OutlookBarGroupsEvents_Event, Interfaces.IOutlookBarGroupsEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarGroups which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOutlookBarGroups WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OutlookBarGroups resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OutlookBarGroups, Interfaces.IOutlookBarGroups>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarPaneEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOutlookBarPaneEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OutlookBarPaneEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OutlookBarPaneEvents_Event, Interfaces.IOutlookBarPaneEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarPane which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOutlookBarPane WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OutlookBarPane resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OutlookBarPane, Interfaces.IOutlookBarPane>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarShortcutsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOutlookBarShortcutsEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OutlookBarShortcutsEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OutlookBarShortcutsEvents_Event, Interfaces.IOutlookBarShortcutsEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OutlookBarShortcuts which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOutlookBarShortcuts WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.OutlookBarShortcuts resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.OutlookBarShortcuts, Interfaces.IOutlookBarShortcuts>();
		}

		/// <summary>
		/// Wrapper interface for PostItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPostItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.PostItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.PostItem, Interfaces.IPostItem>();
		}

		/// <summary>
		/// Wrapper interface for RemoteItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRemoteItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.RemoteItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.RemoteItem, Interfaces.IRemoteItem>();
		}

		/// <summary>
		/// Wrapper interface for ReportItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IReportItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ReportItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ReportItem, Interfaces.IReportItem>();
		}

		/// <summary>
		/// Wrapper interface for TaskRequestAcceptItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITaskRequestAcceptItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.TaskRequestAcceptItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.TaskRequestAcceptItem, Interfaces.ITaskRequestAcceptItem>();
		}

		/// <summary>
		/// Wrapper interface for TaskRequestDeclineItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITaskRequestDeclineItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.TaskRequestDeclineItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.TaskRequestDeclineItem, Interfaces.ITaskRequestDeclineItem>();
		}

		/// <summary>
		/// Wrapper interface for TaskRequestItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITaskRequestItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.TaskRequestItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.TaskRequestItem, Interfaces.ITaskRequestItem>();
		}

		/// <summary>
		/// Wrapper interface for TaskRequestUpdateItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITaskRequestUpdateItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.TaskRequestUpdateItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.TaskRequestUpdateItem, Interfaces.ITaskRequestUpdateItem>();
		}

		/// <summary>
		/// Wrapper interface for ResultsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IResultsEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ResultsEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ResultsEvents_Event, Interfaces.IResultsEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Results which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IResults WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Results resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Results, Interfaces.IResults>();
		}

		/// <summary>
		/// Wrapper interface for _ViewsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_ViewsEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook._ViewsEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook._ViewsEvents_Event, Interfaces.I_ViewsEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Views which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IViews WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Views resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Views, Interfaces.IViews>();
		}

		/// <summary>
		/// Wrapper interface for Reminder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IReminder WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Reminder resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Reminder, Interfaces.IReminder>();
		}

		/// <summary>
		/// Wrapper interface for ReminderCollectionEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IReminderCollectionEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ReminderCollectionEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ReminderCollectionEvents_Event, Interfaces.IReminderCollectionEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Reminders which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IReminders WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Reminders resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Reminders, Interfaces.IReminders>();
		}

		/// <summary>
		/// Wrapper interface for StorageItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IStorageItem WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.StorageItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.StorageItem, Interfaces.IStorageItem>();
		}

		/// <summary>
		/// Wrapper interface for NavigationPaneEvents_12_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INavigationPaneEvents_12_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.NavigationPaneEvents_12_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.NavigationPaneEvents_12_Event, Interfaces.INavigationPaneEvents_12_Event>();
		}

		/// <summary>
		/// Wrapper interface for NavigationPane which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INavigationPane WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.NavigationPane resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.NavigationPane, Interfaces.INavigationPane>();
		}

		/// <summary>
		/// Wrapper interface for NavigationGroupsEvents_12_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INavigationGroupsEvents_12_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.NavigationGroupsEvents_12_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.NavigationGroupsEvents_12_Event, Interfaces.INavigationGroupsEvents_12_Event>();
		}

		/// <summary>
		/// Wrapper interface for NavigationGroups which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INavigationGroups WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.NavigationGroups resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.NavigationGroups, Interfaces.INavigationGroups>();
		}

		/// <summary>
		/// Wrapper interface for DoNotUseMeFolder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDoNotUseMeFolder WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.DoNotUseMeFolder resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.DoNotUseMeFolder, Interfaces.IDoNotUseMeFolder>();
		}

		/// <summary>
		/// Wrapper interface for TimelineView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITimelineView WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.TimelineView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.TimelineView, Interfaces.ITimelineView>();
		}

		/// <summary>
		/// Wrapper interface for MailModule which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailModule WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.MailModule resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.MailModule, Interfaces.IMailModule>();
		}

		/// <summary>
		/// Wrapper interface for CalendarModule which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICalendarModule WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.CalendarModule resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.CalendarModule, Interfaces.ICalendarModule>();
		}

		/// <summary>
		/// Wrapper interface for ContactsModule which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IContactsModule WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ContactsModule resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ContactsModule, Interfaces.IContactsModule>();
		}

		/// <summary>
		/// Wrapper interface for TasksModule which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITasksModule WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.TasksModule resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.TasksModule, Interfaces.ITasksModule>();
		}

		/// <summary>
		/// Wrapper interface for JournalModule which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IJournalModule WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.JournalModule resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.JournalModule, Interfaces.IJournalModule>();
		}

		/// <summary>
		/// Wrapper interface for NotesModule which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INotesModule WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.NotesModule resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.NotesModule, Interfaces.INotesModule>();
		}

		/// <summary>
		/// Wrapper interface for TableView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITableView WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.TableView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.TableView, Interfaces.ITableView>();
		}

		/// <summary>
		/// Wrapper interface for IconView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIconView WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.IconView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.IconView, Interfaces.IIconView>();
		}

		/// <summary>
		/// Wrapper interface for CardView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICardView WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.CardView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.CardView, Interfaces.ICardView>();
		}

		/// <summary>
		/// Wrapper interface for CalendarView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICalendarView WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.CalendarView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.CalendarView, Interfaces.ICalendarView>();
		}

		/// <summary>
		/// Wrapper interface for BusinessCardView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBusinessCardView WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.BusinessCardView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.BusinessCardView, Interfaces.IBusinessCardView>();
		}

		/// <summary>
		/// Wrapper interface for FormRegionStartup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFormRegionStartup WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.FormRegionStartup resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.FormRegionStartup, Interfaces.IFormRegionStartup>();
		}

		/// <summary>
		/// Wrapper interface for TimeZone which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITimeZone WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.TimeZone resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.TimeZone, Interfaces.ITimeZone>();
		}

		/// <summary>
		/// Wrapper interface for SolutionsModule which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISolutionsModule WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.SolutionsModule resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.SolutionsModule, Interfaces.ISolutionsModule>();
		}

		/// <summary>
		/// Wrapper interface for Conversation which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IConversation WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.Conversation resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.Conversation, Interfaces.IConversation>();
		}

		/// <summary>
		/// Wrapper interface for AccountSelectorEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAccountSelectorEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.AccountSelectorEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.AccountSelectorEvents_Event, Interfaces.IAccountSelectorEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for AccountSelector which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAccountSelector WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.AccountSelector resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.AccountSelector, Interfaces.IAccountSelector>();
		}

		/// <summary>
		/// Wrapper interface for ConversationHeader which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IConversationHeader WithComCleanupProxy(this Microsoft.Office.Interop.Outlook.ConversationHeader resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Outlook.ConversationHeader, Interfaces.IConversationHeader>();
		}

	}
}