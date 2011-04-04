//Microsoft.Office.Interop.Outlook, Version=12.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c
namespace Office.Contrib.Interfaces
{
	/// <summary>
	/// Wrapper interface for _IRecipientControl which adds IDispose to the interface
	/// </summary>
	public interface I_IRecipientControl : Microsoft.Office.Interop.Outlook._IRecipientControl, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._IRecipientControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _DRecipientControl which adds IDispose to the interface
	/// </summary>
	public interface I_DRecipientControl : Microsoft.Office.Interop.Outlook._DRecipientControl, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._DRecipientControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _DRecipientControlEvents which adds IDispose to the interface
	/// </summary>
	public interface I_DRecipientControlEvents : Microsoft.Office.Interop.Outlook._DRecipientControlEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._DRecipientControlEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _DRecipientControlEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface I_DRecipientControlEvents_Event : Microsoft.Office.Interop.Outlook._DRecipientControlEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._DRecipientControlEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _RecipientControl which adds IDispose to the interface
	/// </summary>
	public interface I_RecipientControl : Microsoft.Office.Interop.Outlook._RecipientControl, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._RecipientControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _IDocSiteControl which adds IDispose to the interface
	/// </summary>
	public interface I_IDocSiteControl : Microsoft.Office.Interop.Outlook._IDocSiteControl, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._IDocSiteControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _DDocSiteControl which adds IDispose to the interface
	/// </summary>
	public interface I_DDocSiteControl : Microsoft.Office.Interop.Outlook._DDocSiteControl, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._DDocSiteControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _DDocSiteControlEvents which adds IDispose to the interface
	/// </summary>
	public interface I_DDocSiteControlEvents : Microsoft.Office.Interop.Outlook._DDocSiteControlEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._DDocSiteControlEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _DDocSiteControlEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface I_DDocSiteControlEvents_Event : Microsoft.Office.Interop.Outlook._DDocSiteControlEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._DDocSiteControlEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _DocSiteControl which adds IDispose to the interface
	/// </summary>
	public interface I_DocSiteControl : Microsoft.Office.Interop.Outlook._DocSiteControl, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._DocSiteControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkControl which adds IDispose to the interface
	/// </summary>
	public interface IOlkControl : Microsoft.Office.Interop.Outlook.OlkControl, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OlkTextBox which adds IDispose to the interface
	/// </summary>
	public interface I_OlkTextBox : Microsoft.Office.Interop.Outlook._OlkTextBox, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OlkTextBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkTextBoxEvents which adds IDispose to the interface
	/// </summary>
	public interface IOlkTextBoxEvents : Microsoft.Office.Interop.Outlook.OlkTextBoxEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkTextBoxEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkTextBoxEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOlkTextBoxEvents_Event : Microsoft.Office.Interop.Outlook.OlkTextBoxEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkTextBoxEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkTextBox which adds IDispose to the interface
	/// </summary>
	public interface IOlkTextBox : Microsoft.Office.Interop.Outlook.OlkTextBox, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkTextBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OlkLabel which adds IDispose to the interface
	/// </summary>
	public interface I_OlkLabel : Microsoft.Office.Interop.Outlook._OlkLabel, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OlkLabel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkLabelEvents which adds IDispose to the interface
	/// </summary>
	public interface IOlkLabelEvents : Microsoft.Office.Interop.Outlook.OlkLabelEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkLabelEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkLabelEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOlkLabelEvents_Event : Microsoft.Office.Interop.Outlook.OlkLabelEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkLabelEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkLabel which adds IDispose to the interface
	/// </summary>
	public interface IOlkLabel : Microsoft.Office.Interop.Outlook.OlkLabel, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkLabel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OlkCommandButton which adds IDispose to the interface
	/// </summary>
	public interface I_OlkCommandButton : Microsoft.Office.Interop.Outlook._OlkCommandButton, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OlkCommandButton Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkCommandButtonEvents which adds IDispose to the interface
	/// </summary>
	public interface IOlkCommandButtonEvents : Microsoft.Office.Interop.Outlook.OlkCommandButtonEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkCommandButtonEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkCommandButtonEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOlkCommandButtonEvents_Event : Microsoft.Office.Interop.Outlook.OlkCommandButtonEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkCommandButtonEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkCommandButton which adds IDispose to the interface
	/// </summary>
	public interface IOlkCommandButton : Microsoft.Office.Interop.Outlook.OlkCommandButton, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkCommandButton Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OlkCheckBox which adds IDispose to the interface
	/// </summary>
	public interface I_OlkCheckBox : Microsoft.Office.Interop.Outlook._OlkCheckBox, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OlkCheckBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkCheckBoxEvents which adds IDispose to the interface
	/// </summary>
	public interface IOlkCheckBoxEvents : Microsoft.Office.Interop.Outlook.OlkCheckBoxEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkCheckBoxEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkCheckBoxEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOlkCheckBoxEvents_Event : Microsoft.Office.Interop.Outlook.OlkCheckBoxEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkCheckBoxEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkCheckBox which adds IDispose to the interface
	/// </summary>
	public interface IOlkCheckBox : Microsoft.Office.Interop.Outlook.OlkCheckBox, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkCheckBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OlkOptionButton which adds IDispose to the interface
	/// </summary>
	public interface I_OlkOptionButton : Microsoft.Office.Interop.Outlook._OlkOptionButton, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OlkOptionButton Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkOptionButtonEvents which adds IDispose to the interface
	/// </summary>
	public interface IOlkOptionButtonEvents : Microsoft.Office.Interop.Outlook.OlkOptionButtonEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkOptionButtonEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkOptionButtonEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOlkOptionButtonEvents_Event : Microsoft.Office.Interop.Outlook.OlkOptionButtonEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkOptionButtonEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkOptionButton which adds IDispose to the interface
	/// </summary>
	public interface IOlkOptionButton : Microsoft.Office.Interop.Outlook.OlkOptionButton, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkOptionButton Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OlkComboBox which adds IDispose to the interface
	/// </summary>
	public interface I_OlkComboBox : Microsoft.Office.Interop.Outlook._OlkComboBox, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OlkComboBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkComboBoxEvents which adds IDispose to the interface
	/// </summary>
	public interface IOlkComboBoxEvents : Microsoft.Office.Interop.Outlook.OlkComboBoxEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkComboBoxEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkComboBoxEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOlkComboBoxEvents_Event : Microsoft.Office.Interop.Outlook.OlkComboBoxEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkComboBoxEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkComboBox which adds IDispose to the interface
	/// </summary>
	public interface IOlkComboBox : Microsoft.Office.Interop.Outlook.OlkComboBox, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkComboBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OlkListBox which adds IDispose to the interface
	/// </summary>
	public interface I_OlkListBox : Microsoft.Office.Interop.Outlook._OlkListBox, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OlkListBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkListBoxEvents which adds IDispose to the interface
	/// </summary>
	public interface IOlkListBoxEvents : Microsoft.Office.Interop.Outlook.OlkListBoxEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkListBoxEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkListBoxEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOlkListBoxEvents_Event : Microsoft.Office.Interop.Outlook.OlkListBoxEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkListBoxEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkListBox which adds IDispose to the interface
	/// </summary>
	public interface IOlkListBox : Microsoft.Office.Interop.Outlook.OlkListBox, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkListBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OlkInfoBar which adds IDispose to the interface
	/// </summary>
	public interface I_OlkInfoBar : Microsoft.Office.Interop.Outlook._OlkInfoBar, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OlkInfoBar Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkInfoBarEvents which adds IDispose to the interface
	/// </summary>
	public interface IOlkInfoBarEvents : Microsoft.Office.Interop.Outlook.OlkInfoBarEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkInfoBarEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkInfoBarEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOlkInfoBarEvents_Event : Microsoft.Office.Interop.Outlook.OlkInfoBarEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkInfoBarEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkInfoBar which adds IDispose to the interface
	/// </summary>
	public interface IOlkInfoBar : Microsoft.Office.Interop.Outlook.OlkInfoBar, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkInfoBar Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OlkContactPhoto which adds IDispose to the interface
	/// </summary>
	public interface I_OlkContactPhoto : Microsoft.Office.Interop.Outlook._OlkContactPhoto, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OlkContactPhoto Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkContactPhotoEvents which adds IDispose to the interface
	/// </summary>
	public interface IOlkContactPhotoEvents : Microsoft.Office.Interop.Outlook.OlkContactPhotoEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkContactPhotoEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkContactPhotoEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOlkContactPhotoEvents_Event : Microsoft.Office.Interop.Outlook.OlkContactPhotoEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkContactPhotoEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkContactPhoto which adds IDispose to the interface
	/// </summary>
	public interface IOlkContactPhoto : Microsoft.Office.Interop.Outlook.OlkContactPhoto, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkContactPhoto Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OlkBusinessCardControl which adds IDispose to the interface
	/// </summary>
	public interface I_OlkBusinessCardControl : Microsoft.Office.Interop.Outlook._OlkBusinessCardControl, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OlkBusinessCardControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkBusinessCardControlEvents which adds IDispose to the interface
	/// </summary>
	public interface IOlkBusinessCardControlEvents : Microsoft.Office.Interop.Outlook.OlkBusinessCardControlEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkBusinessCardControlEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkBusinessCardControlEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOlkBusinessCardControlEvents_Event : Microsoft.Office.Interop.Outlook.OlkBusinessCardControlEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkBusinessCardControlEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkBusinessCardControl which adds IDispose to the interface
	/// </summary>
	public interface IOlkBusinessCardControl : Microsoft.Office.Interop.Outlook.OlkBusinessCardControl, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkBusinessCardControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OlkPageControl which adds IDispose to the interface
	/// </summary>
	public interface I_OlkPageControl : Microsoft.Office.Interop.Outlook._OlkPageControl, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OlkPageControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkPageControlEvents which adds IDispose to the interface
	/// </summary>
	public interface IOlkPageControlEvents : Microsoft.Office.Interop.Outlook.OlkPageControlEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkPageControlEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkPageControlEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOlkPageControlEvents_Event : Microsoft.Office.Interop.Outlook.OlkPageControlEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkPageControlEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkPageControl which adds IDispose to the interface
	/// </summary>
	public interface IOlkPageControl : Microsoft.Office.Interop.Outlook.OlkPageControl, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkPageControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OlkDateControl which adds IDispose to the interface
	/// </summary>
	public interface I_OlkDateControl : Microsoft.Office.Interop.Outlook._OlkDateControl, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OlkDateControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkDateControlEvents which adds IDispose to the interface
	/// </summary>
	public interface IOlkDateControlEvents : Microsoft.Office.Interop.Outlook.OlkDateControlEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkDateControlEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkDateControlEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOlkDateControlEvents_Event : Microsoft.Office.Interop.Outlook.OlkDateControlEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkDateControlEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkDateControl which adds IDispose to the interface
	/// </summary>
	public interface IOlkDateControl : Microsoft.Office.Interop.Outlook.OlkDateControl, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkDateControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OlkTimeControl which adds IDispose to the interface
	/// </summary>
	public interface I_OlkTimeControl : Microsoft.Office.Interop.Outlook._OlkTimeControl, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OlkTimeControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkTimeControlEvents which adds IDispose to the interface
	/// </summary>
	public interface IOlkTimeControlEvents : Microsoft.Office.Interop.Outlook.OlkTimeControlEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkTimeControlEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkTimeControlEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOlkTimeControlEvents_Event : Microsoft.Office.Interop.Outlook.OlkTimeControlEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkTimeControlEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkTimeControl which adds IDispose to the interface
	/// </summary>
	public interface IOlkTimeControl : Microsoft.Office.Interop.Outlook.OlkTimeControl, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkTimeControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OlkCategory which adds IDispose to the interface
	/// </summary>
	public interface I_OlkCategory : Microsoft.Office.Interop.Outlook._OlkCategory, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OlkCategory Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkCategoryEvents which adds IDispose to the interface
	/// </summary>
	public interface IOlkCategoryEvents : Microsoft.Office.Interop.Outlook.OlkCategoryEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkCategoryEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkCategoryEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOlkCategoryEvents_Event : Microsoft.Office.Interop.Outlook.OlkCategoryEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkCategoryEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkCategory which adds IDispose to the interface
	/// </summary>
	public interface IOlkCategory : Microsoft.Office.Interop.Outlook.OlkCategory, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkCategory Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OlkFrameHeader which adds IDispose to the interface
	/// </summary>
	public interface I_OlkFrameHeader : Microsoft.Office.Interop.Outlook._OlkFrameHeader, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OlkFrameHeader Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkFrameHeaderEvents which adds IDispose to the interface
	/// </summary>
	public interface IOlkFrameHeaderEvents : Microsoft.Office.Interop.Outlook.OlkFrameHeaderEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkFrameHeaderEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkFrameHeaderEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOlkFrameHeaderEvents_Event : Microsoft.Office.Interop.Outlook.OlkFrameHeaderEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkFrameHeaderEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkFrameHeader which adds IDispose to the interface
	/// </summary>
	public interface IOlkFrameHeader : Microsoft.Office.Interop.Outlook.OlkFrameHeader, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkFrameHeader Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OlkSenderPhoto which adds IDispose to the interface
	/// </summary>
	public interface I_OlkSenderPhoto : Microsoft.Office.Interop.Outlook._OlkSenderPhoto, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OlkSenderPhoto Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkSenderPhotoEvents which adds IDispose to the interface
	/// </summary>
	public interface IOlkSenderPhotoEvents : Microsoft.Office.Interop.Outlook.OlkSenderPhotoEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkSenderPhotoEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkSenderPhotoEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOlkSenderPhotoEvents_Event : Microsoft.Office.Interop.Outlook.OlkSenderPhotoEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkSenderPhotoEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkSenderPhoto which adds IDispose to the interface
	/// </summary>
	public interface IOlkSenderPhoto : Microsoft.Office.Interop.Outlook.OlkSenderPhoto, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkSenderPhoto Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _TimeZone which adds IDispose to the interface
	/// </summary>
	public interface I_TimeZone : Microsoft.Office.Interop.Outlook._TimeZone, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._TimeZone Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Application which adds IDispose to the interface
	/// </summary>
	public interface I_Application : Microsoft.Office.Interop.Outlook._Application, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Application Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _NameSpace which adds IDispose to the interface
	/// </summary>
	public interface I_NameSpace : Microsoft.Office.Interop.Outlook._NameSpace, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._NameSpace Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Recipient which adds IDispose to the interface
	/// </summary>
	public interface IRecipient : Microsoft.Office.Interop.Outlook.Recipient, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Recipient Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AddressEntry which adds IDispose to the interface
	/// </summary>
	public interface IAddressEntry : Microsoft.Office.Interop.Outlook.AddressEntry, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.AddressEntry Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AddressEntries which adds IDispose to the interface
	/// </summary>
	public interface IAddressEntries : Microsoft.Office.Interop.Outlook.AddressEntries, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.AddressEntries Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _ContactItem which adds IDispose to the interface
	/// </summary>
	public interface I_ContactItem : Microsoft.Office.Interop.Outlook._ContactItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._ContactItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Actions which adds IDispose to the interface
	/// </summary>
	public interface IActions : Microsoft.Office.Interop.Outlook.Actions, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Actions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Action which adds IDispose to the interface
	/// </summary>
	public interface IAction : Microsoft.Office.Interop.Outlook.Action, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Action Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Attachments which adds IDispose to the interface
	/// </summary>
	public interface IAttachments : Microsoft.Office.Interop.Outlook.Attachments, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Attachments Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Attachment which adds IDispose to the interface
	/// </summary>
	public interface IAttachment : Microsoft.Office.Interop.Outlook.Attachment, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Attachment Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PropertyAccessor which adds IDispose to the interface
	/// </summary>
	public interface IPropertyAccessor : Microsoft.Office.Interop.Outlook.PropertyAccessor, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.PropertyAccessor Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _PropertyAccessor which adds IDispose to the interface
	/// </summary>
	public interface I_PropertyAccessor : Microsoft.Office.Interop.Outlook._PropertyAccessor, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._PropertyAccessor Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FormDescription which adds IDispose to the interface
	/// </summary>
	public interface IFormDescription : Microsoft.Office.Interop.Outlook.FormDescription, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.FormDescription Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Inspector which adds IDispose to the interface
	/// </summary>
	public interface I_Inspector : Microsoft.Office.Interop.Outlook._Inspector, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Inspector Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for UserProperties which adds IDispose to the interface
	/// </summary>
	public interface IUserProperties : Microsoft.Office.Interop.Outlook.UserProperties, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.UserProperties Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for UserProperty which adds IDispose to the interface
	/// </summary>
	public interface IUserProperty : Microsoft.Office.Interop.Outlook.UserProperty, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.UserProperty Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MAPIFolder which adds IDispose to the interface
	/// </summary>
	public interface IMAPIFolder : Microsoft.Office.Interop.Outlook.MAPIFolder, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.MAPIFolder Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Folders which adds IDispose to the interface
	/// </summary>
	public interface I_Folders : Microsoft.Office.Interop.Outlook._Folders, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Folders Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Items which adds IDispose to the interface
	/// </summary>
	public interface I_Items : Microsoft.Office.Interop.Outlook._Items, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Items Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Explorer which adds IDispose to the interface
	/// </summary>
	public interface I_Explorer : Microsoft.Office.Interop.Outlook._Explorer, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Explorer Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Panes which adds IDispose to the interface
	/// </summary>
	public interface IPanes : Microsoft.Office.Interop.Outlook.Panes, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Panes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Selection which adds IDispose to the interface
	/// </summary>
	public interface ISelection : Microsoft.Office.Interop.Outlook.Selection, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Selection Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _NavigationPane which adds IDispose to the interface
	/// </summary>
	public interface I_NavigationPane : Microsoft.Office.Interop.Outlook._NavigationPane, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._NavigationPane Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for NavigationModule which adds IDispose to the interface
	/// </summary>
	public interface INavigationModule : Microsoft.Office.Interop.Outlook.NavigationModule, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.NavigationModule Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _NavigationModule which adds IDispose to the interface
	/// </summary>
	public interface I_NavigationModule : Microsoft.Office.Interop.Outlook._NavigationModule, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._NavigationModule Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for NavigationModules which adds IDispose to the interface
	/// </summary>
	public interface INavigationModules : Microsoft.Office.Interop.Outlook.NavigationModules, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.NavigationModules Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _NavigationModules which adds IDispose to the interface
	/// </summary>
	public interface I_NavigationModules : Microsoft.Office.Interop.Outlook._NavigationModules, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._NavigationModules Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for View which adds IDispose to the interface
	/// </summary>
	public interface IView : Microsoft.Office.Interop.Outlook.View, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.View Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Views which adds IDispose to the interface
	/// </summary>
	public interface I_Views : Microsoft.Office.Interop.Outlook._Views, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Views Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Store which adds IDispose to the interface
	/// </summary>
	public interface IStore : Microsoft.Office.Interop.Outlook.Store, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Store Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Store which adds IDispose to the interface
	/// </summary>
	public interface I_Store : Microsoft.Office.Interop.Outlook._Store, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Store Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Rules which adds IDispose to the interface
	/// </summary>
	public interface IRules : Microsoft.Office.Interop.Outlook.Rules, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Rules Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Rules which adds IDispose to the interface
	/// </summary>
	public interface I_Rules : Microsoft.Office.Interop.Outlook._Rules, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Rules Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Rule which adds IDispose to the interface
	/// </summary>
	public interface I_Rule : Microsoft.Office.Interop.Outlook._Rule, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Rule Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RuleActions which adds IDispose to the interface
	/// </summary>
	public interface IRuleActions : Microsoft.Office.Interop.Outlook.RuleActions, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.RuleActions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _RuleActions which adds IDispose to the interface
	/// </summary>
	public interface I_RuleActions : Microsoft.Office.Interop.Outlook._RuleActions, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._RuleActions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _RuleAction which adds IDispose to the interface
	/// </summary>
	public interface I_RuleAction : Microsoft.Office.Interop.Outlook._RuleAction, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._RuleAction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MoveOrCopyRuleAction which adds IDispose to the interface
	/// </summary>
	public interface IMoveOrCopyRuleAction : Microsoft.Office.Interop.Outlook.MoveOrCopyRuleAction, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.MoveOrCopyRuleAction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _MoveOrCopyRuleAction which adds IDispose to the interface
	/// </summary>
	public interface I_MoveOrCopyRuleAction : Microsoft.Office.Interop.Outlook._MoveOrCopyRuleAction, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._MoveOrCopyRuleAction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RuleAction which adds IDispose to the interface
	/// </summary>
	public interface IRuleAction : Microsoft.Office.Interop.Outlook.RuleAction, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.RuleAction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SendRuleAction which adds IDispose to the interface
	/// </summary>
	public interface ISendRuleAction : Microsoft.Office.Interop.Outlook.SendRuleAction, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.SendRuleAction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _SendRuleAction which adds IDispose to the interface
	/// </summary>
	public interface I_SendRuleAction : Microsoft.Office.Interop.Outlook._SendRuleAction, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._SendRuleAction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Recipients which adds IDispose to the interface
	/// </summary>
	public interface IRecipients : Microsoft.Office.Interop.Outlook.Recipients, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Recipients Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AssignToCategoryRuleAction which adds IDispose to the interface
	/// </summary>
	public interface IAssignToCategoryRuleAction : Microsoft.Office.Interop.Outlook.AssignToCategoryRuleAction, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.AssignToCategoryRuleAction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _AssignToCategoryRuleAction which adds IDispose to the interface
	/// </summary>
	public interface I_AssignToCategoryRuleAction : Microsoft.Office.Interop.Outlook._AssignToCategoryRuleAction, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._AssignToCategoryRuleAction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PlaySoundRuleAction which adds IDispose to the interface
	/// </summary>
	public interface IPlaySoundRuleAction : Microsoft.Office.Interop.Outlook.PlaySoundRuleAction, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.PlaySoundRuleAction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _PlaySoundRuleAction which adds IDispose to the interface
	/// </summary>
	public interface I_PlaySoundRuleAction : Microsoft.Office.Interop.Outlook._PlaySoundRuleAction, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._PlaySoundRuleAction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MarkAsTaskRuleAction which adds IDispose to the interface
	/// </summary>
	public interface IMarkAsTaskRuleAction : Microsoft.Office.Interop.Outlook.MarkAsTaskRuleAction, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.MarkAsTaskRuleAction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _MarkAsTaskRuleAction which adds IDispose to the interface
	/// </summary>
	public interface I_MarkAsTaskRuleAction : Microsoft.Office.Interop.Outlook._MarkAsTaskRuleAction, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._MarkAsTaskRuleAction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for NewItemAlertRuleAction which adds IDispose to the interface
	/// </summary>
	public interface INewItemAlertRuleAction : Microsoft.Office.Interop.Outlook.NewItemAlertRuleAction, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.NewItemAlertRuleAction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _NewItemAlertRuleAction which adds IDispose to the interface
	/// </summary>
	public interface I_NewItemAlertRuleAction : Microsoft.Office.Interop.Outlook._NewItemAlertRuleAction, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._NewItemAlertRuleAction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RuleConditions which adds IDispose to the interface
	/// </summary>
	public interface IRuleConditions : Microsoft.Office.Interop.Outlook.RuleConditions, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.RuleConditions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _RuleConditions which adds IDispose to the interface
	/// </summary>
	public interface I_RuleConditions : Microsoft.Office.Interop.Outlook._RuleConditions, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._RuleConditions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _RuleCondition which adds IDispose to the interface
	/// </summary>
	public interface I_RuleCondition : Microsoft.Office.Interop.Outlook._RuleCondition, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._RuleCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RuleCondition which adds IDispose to the interface
	/// </summary>
	public interface IRuleCondition : Microsoft.Office.Interop.Outlook.RuleCondition, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.RuleCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ImportanceRuleCondition which adds IDispose to the interface
	/// </summary>
	public interface IImportanceRuleCondition : Microsoft.Office.Interop.Outlook.ImportanceRuleCondition, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ImportanceRuleCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _ImportanceRuleCondition which adds IDispose to the interface
	/// </summary>
	public interface I_ImportanceRuleCondition : Microsoft.Office.Interop.Outlook._ImportanceRuleCondition, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._ImportanceRuleCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AccountRuleCondition which adds IDispose to the interface
	/// </summary>
	public interface IAccountRuleCondition : Microsoft.Office.Interop.Outlook.AccountRuleCondition, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.AccountRuleCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _AccountRuleCondition which adds IDispose to the interface
	/// </summary>
	public interface I_AccountRuleCondition : Microsoft.Office.Interop.Outlook._AccountRuleCondition, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._AccountRuleCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Account which adds IDispose to the interface
	/// </summary>
	public interface IAccount : Microsoft.Office.Interop.Outlook.Account, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Account Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Account which adds IDispose to the interface
	/// </summary>
	public interface I_Account : Microsoft.Office.Interop.Outlook._Account, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Account Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextRuleCondition which adds IDispose to the interface
	/// </summary>
	public interface ITextRuleCondition : Microsoft.Office.Interop.Outlook.TextRuleCondition, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.TextRuleCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _TextRuleCondition which adds IDispose to the interface
	/// </summary>
	public interface I_TextRuleCondition : Microsoft.Office.Interop.Outlook._TextRuleCondition, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._TextRuleCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CategoryRuleCondition which adds IDispose to the interface
	/// </summary>
	public interface ICategoryRuleCondition : Microsoft.Office.Interop.Outlook.CategoryRuleCondition, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.CategoryRuleCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CategoryRuleCondition which adds IDispose to the interface
	/// </summary>
	public interface I_CategoryRuleCondition : Microsoft.Office.Interop.Outlook._CategoryRuleCondition, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._CategoryRuleCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FormNameRuleCondition which adds IDispose to the interface
	/// </summary>
	public interface IFormNameRuleCondition : Microsoft.Office.Interop.Outlook.FormNameRuleCondition, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.FormNameRuleCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _FormNameRuleCondition which adds IDispose to the interface
	/// </summary>
	public interface I_FormNameRuleCondition : Microsoft.Office.Interop.Outlook._FormNameRuleCondition, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._FormNameRuleCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ToOrFromRuleCondition which adds IDispose to the interface
	/// </summary>
	public interface IToOrFromRuleCondition : Microsoft.Office.Interop.Outlook.ToOrFromRuleCondition, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ToOrFromRuleCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _ToOrFromRuleCondition which adds IDispose to the interface
	/// </summary>
	public interface I_ToOrFromRuleCondition : Microsoft.Office.Interop.Outlook._ToOrFromRuleCondition, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._ToOrFromRuleCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AddressRuleCondition which adds IDispose to the interface
	/// </summary>
	public interface IAddressRuleCondition : Microsoft.Office.Interop.Outlook.AddressRuleCondition, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.AddressRuleCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _AddressRuleCondition which adds IDispose to the interface
	/// </summary>
	public interface I_AddressRuleCondition : Microsoft.Office.Interop.Outlook._AddressRuleCondition, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._AddressRuleCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SenderInAddressListRuleCondition which adds IDispose to the interface
	/// </summary>
	public interface ISenderInAddressListRuleCondition : Microsoft.Office.Interop.Outlook.SenderInAddressListRuleCondition, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.SenderInAddressListRuleCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _SenderInAddressListRuleCondition which adds IDispose to the interface
	/// </summary>
	public interface I_SenderInAddressListRuleCondition : Microsoft.Office.Interop.Outlook._SenderInAddressListRuleCondition, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._SenderInAddressListRuleCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AddressList which adds IDispose to the interface
	/// </summary>
	public interface IAddressList : Microsoft.Office.Interop.Outlook.AddressList, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.AddressList Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FromRssFeedRuleCondition which adds IDispose to the interface
	/// </summary>
	public interface IFromRssFeedRuleCondition : Microsoft.Office.Interop.Outlook.FromRssFeedRuleCondition, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.FromRssFeedRuleCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _FromRssFeedRuleCondition which adds IDispose to the interface
	/// </summary>
	public interface I_FromRssFeedRuleCondition : Microsoft.Office.Interop.Outlook._FromRssFeedRuleCondition, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._FromRssFeedRuleCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Rule which adds IDispose to the interface
	/// </summary>
	public interface IRule : Microsoft.Office.Interop.Outlook.Rule, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Rule Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _StorageItem which adds IDispose to the interface
	/// </summary>
	public interface I_StorageItem : Microsoft.Office.Interop.Outlook._StorageItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._StorageItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Table which adds IDispose to the interface
	/// </summary>
	public interface ITable : Microsoft.Office.Interop.Outlook.Table, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Table Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Table which adds IDispose to the interface
	/// </summary>
	public interface I_Table : Microsoft.Office.Interop.Outlook._Table, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Table Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Row which adds IDispose to the interface
	/// </summary>
	public interface IRow : Microsoft.Office.Interop.Outlook.Row, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Row Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Row which adds IDispose to the interface
	/// </summary>
	public interface I_Row : Microsoft.Office.Interop.Outlook._Row, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Row Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Columns which adds IDispose to the interface
	/// </summary>
	public interface IColumns : Microsoft.Office.Interop.Outlook.Columns, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Columns Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Columns which adds IDispose to the interface
	/// </summary>
	public interface I_Columns : Microsoft.Office.Interop.Outlook._Columns, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Columns Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Column which adds IDispose to the interface
	/// </summary>
	public interface I_Column : Microsoft.Office.Interop.Outlook._Column, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Column Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Column which adds IDispose to the interface
	/// </summary>
	public interface IColumn : Microsoft.Office.Interop.Outlook.Column, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Column Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CalendarSharing which adds IDispose to the interface
	/// </summary>
	public interface ICalendarSharing : Microsoft.Office.Interop.Outlook.CalendarSharing, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.CalendarSharing Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CalendarSharing which adds IDispose to the interface
	/// </summary>
	public interface I_CalendarSharing : Microsoft.Office.Interop.Outlook._CalendarSharing, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._CalendarSharing Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ItemEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IItemEvents_Event : Microsoft.Office.Interop.Outlook.ItemEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ItemEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ItemEvents_10_Event which adds IDispose to the interface
	/// </summary>
	public interface IItemEvents_10_Event : Microsoft.Office.Interop.Outlook.ItemEvents_10_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ItemEvents_10_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MailItem which adds IDispose to the interface
	/// </summary>
	public interface IMailItem : Microsoft.Office.Interop.Outlook.MailItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.MailItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _MailItem which adds IDispose to the interface
	/// </summary>
	public interface I_MailItem : Microsoft.Office.Interop.Outlook._MailItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._MailItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Links which adds IDispose to the interface
	/// </summary>
	public interface ILinks : Microsoft.Office.Interop.Outlook.Links, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Links Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Link which adds IDispose to the interface
	/// </summary>
	public interface ILink : Microsoft.Office.Interop.Outlook.Link, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Link Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ItemProperties which adds IDispose to the interface
	/// </summary>
	public interface IItemProperties : Microsoft.Office.Interop.Outlook.ItemProperties, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ItemProperties Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ItemProperty which adds IDispose to the interface
	/// </summary>
	public interface IItemProperty : Microsoft.Office.Interop.Outlook.ItemProperty, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ItemProperty Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Conflicts which adds IDispose to the interface
	/// </summary>
	public interface IConflicts : Microsoft.Office.Interop.Outlook.Conflicts, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Conflicts Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Conflict which adds IDispose to the interface
	/// </summary>
	public interface IConflict : Microsoft.Office.Interop.Outlook.Conflict, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Conflict Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ContactItem which adds IDispose to the interface
	/// </summary>
	public interface IContactItem : Microsoft.Office.Interop.Outlook.ContactItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ContactItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ItemEvents which adds IDispose to the interface
	/// </summary>
	public interface IItemEvents : Microsoft.Office.Interop.Outlook.ItemEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ItemEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ItemEvents_10 which adds IDispose to the interface
	/// </summary>
	public interface IItemEvents_10 : Microsoft.Office.Interop.Outlook.ItemEvents_10, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ItemEvents_10 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for UserDefinedProperties which adds IDispose to the interface
	/// </summary>
	public interface IUserDefinedProperties : Microsoft.Office.Interop.Outlook.UserDefinedProperties, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.UserDefinedProperties Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _UserDefinedProperties which adds IDispose to the interface
	/// </summary>
	public interface I_UserDefinedProperties : Microsoft.Office.Interop.Outlook._UserDefinedProperties, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._UserDefinedProperties Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _UserDefinedProperty which adds IDispose to the interface
	/// </summary>
	public interface I_UserDefinedProperty : Microsoft.Office.Interop.Outlook._UserDefinedProperty, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._UserDefinedProperty Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for UserDefinedProperty which adds IDispose to the interface
	/// </summary>
	public interface IUserDefinedProperty : Microsoft.Office.Interop.Outlook.UserDefinedProperty, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.UserDefinedProperty Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ExchangeUser which adds IDispose to the interface
	/// </summary>
	public interface IExchangeUser : Microsoft.Office.Interop.Outlook.ExchangeUser, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ExchangeUser Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _ExchangeUser which adds IDispose to the interface
	/// </summary>
	public interface I_ExchangeUser : Microsoft.Office.Interop.Outlook._ExchangeUser, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._ExchangeUser Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ExchangeDistributionList which adds IDispose to the interface
	/// </summary>
	public interface IExchangeDistributionList : Microsoft.Office.Interop.Outlook.ExchangeDistributionList, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ExchangeDistributionList Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _ExchangeDistributionList which adds IDispose to the interface
	/// </summary>
	public interface I_ExchangeDistributionList : Microsoft.Office.Interop.Outlook._ExchangeDistributionList, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._ExchangeDistributionList Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AddressLists which adds IDispose to the interface
	/// </summary>
	public interface IAddressLists : Microsoft.Office.Interop.Outlook.AddressLists, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.AddressLists Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SyncObjects which adds IDispose to the interface
	/// </summary>
	public interface ISyncObjects : Microsoft.Office.Interop.Outlook.SyncObjects, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.SyncObjects Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SyncObjectEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface ISyncObjectEvents_Event : Microsoft.Office.Interop.Outlook.SyncObjectEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.SyncObjectEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SyncObject which adds IDispose to the interface
	/// </summary>
	public interface ISyncObject : Microsoft.Office.Interop.Outlook.SyncObject, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.SyncObject Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _SyncObject which adds IDispose to the interface
	/// </summary>
	public interface I_SyncObject : Microsoft.Office.Interop.Outlook._SyncObject, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._SyncObject Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SyncObjectEvents which adds IDispose to the interface
	/// </summary>
	public interface ISyncObjectEvents : Microsoft.Office.Interop.Outlook.SyncObjectEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.SyncObjectEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Accounts which adds IDispose to the interface
	/// </summary>
	public interface IAccounts : Microsoft.Office.Interop.Outlook.Accounts, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Accounts Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Accounts which adds IDispose to the interface
	/// </summary>
	public interface I_Accounts : Microsoft.Office.Interop.Outlook._Accounts, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Accounts Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for StoresEvents_12_Event which adds IDispose to the interface
	/// </summary>
	public interface IStoresEvents_12_Event : Microsoft.Office.Interop.Outlook.StoresEvents_12_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.StoresEvents_12_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Stores which adds IDispose to the interface
	/// </summary>
	public interface IStores : Microsoft.Office.Interop.Outlook.Stores, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Stores Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Stores which adds IDispose to the interface
	/// </summary>
	public interface I_Stores : Microsoft.Office.Interop.Outlook._Stores, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Stores Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for StoresEvents_12 which adds IDispose to the interface
	/// </summary>
	public interface IStoresEvents_12 : Microsoft.Office.Interop.Outlook.StoresEvents_12, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.StoresEvents_12 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SelectNamesDialog which adds IDispose to the interface
	/// </summary>
	public interface ISelectNamesDialog : Microsoft.Office.Interop.Outlook.SelectNamesDialog, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.SelectNamesDialog Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _SelectNamesDialog which adds IDispose to the interface
	/// </summary>
	public interface I_SelectNamesDialog : Microsoft.Office.Interop.Outlook._SelectNamesDialog, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._SelectNamesDialog Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Categories which adds IDispose to the interface
	/// </summary>
	public interface ICategories : Microsoft.Office.Interop.Outlook.Categories, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Categories Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Categories which adds IDispose to the interface
	/// </summary>
	public interface I_Categories : Microsoft.Office.Interop.Outlook._Categories, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Categories Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Category which adds IDispose to the interface
	/// </summary>
	public interface I_Category : Microsoft.Office.Interop.Outlook._Category, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Category Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Category which adds IDispose to the interface
	/// </summary>
	public interface ICategory : Microsoft.Office.Interop.Outlook.Category, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Category Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SharingItem which adds IDispose to the interface
	/// </summary>
	public interface ISharingItem : Microsoft.Office.Interop.Outlook.SharingItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.SharingItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _SharingItem which adds IDispose to the interface
	/// </summary>
	public interface I_SharingItem : Microsoft.Office.Interop.Outlook._SharingItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._SharingItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Explorers which adds IDispose to the interface
	/// </summary>
	public interface I_Explorers : Microsoft.Office.Interop.Outlook._Explorers, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Explorers Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ExplorerEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IExplorerEvents_Event : Microsoft.Office.Interop.Outlook.ExplorerEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ExplorerEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ExplorerEvents_10_Event which adds IDispose to the interface
	/// </summary>
	public interface IExplorerEvents_10_Event : Microsoft.Office.Interop.Outlook.ExplorerEvents_10_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ExplorerEvents_10_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Explorer which adds IDispose to the interface
	/// </summary>
	public interface IExplorer : Microsoft.Office.Interop.Outlook.Explorer, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Explorer Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ExplorerEvents which adds IDispose to the interface
	/// </summary>
	public interface IExplorerEvents : Microsoft.Office.Interop.Outlook.ExplorerEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ExplorerEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ExplorerEvents_10 which adds IDispose to the interface
	/// </summary>
	public interface IExplorerEvents_10 : Microsoft.Office.Interop.Outlook.ExplorerEvents_10, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ExplorerEvents_10 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Inspectors which adds IDispose to the interface
	/// </summary>
	public interface I_Inspectors : Microsoft.Office.Interop.Outlook._Inspectors, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Inspectors Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for InspectorEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IInspectorEvents_Event : Microsoft.Office.Interop.Outlook.InspectorEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.InspectorEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for InspectorEvents_10_Event which adds IDispose to the interface
	/// </summary>
	public interface IInspectorEvents_10_Event : Microsoft.Office.Interop.Outlook.InspectorEvents_10_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.InspectorEvents_10_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Inspector which adds IDispose to the interface
	/// </summary>
	public interface IInspector : Microsoft.Office.Interop.Outlook.Inspector, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Inspector Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for InspectorEvents which adds IDispose to the interface
	/// </summary>
	public interface IInspectorEvents : Microsoft.Office.Interop.Outlook.InspectorEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.InspectorEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for InspectorEvents_10 which adds IDispose to the interface
	/// </summary>
	public interface IInspectorEvents_10 : Microsoft.Office.Interop.Outlook.InspectorEvents_10, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.InspectorEvents_10 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Search which adds IDispose to the interface
	/// </summary>
	public interface ISearch : Microsoft.Office.Interop.Outlook.Search, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Search Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Results which adds IDispose to the interface
	/// </summary>
	public interface I_Results : Microsoft.Office.Interop.Outlook._Results, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Results Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Reminders which adds IDispose to the interface
	/// </summary>
	public interface I_Reminders : Microsoft.Office.Interop.Outlook._Reminders, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Reminders Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Reminder which adds IDispose to the interface
	/// </summary>
	public interface I_Reminder : Microsoft.Office.Interop.Outlook._Reminder, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._Reminder Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TimeZones which adds IDispose to the interface
	/// </summary>
	public interface ITimeZones : Microsoft.Office.Interop.Outlook.TimeZones, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.TimeZones Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _TimeZones which adds IDispose to the interface
	/// </summary>
	public interface I_TimeZones : Microsoft.Office.Interop.Outlook._TimeZones, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._TimeZones Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OlkTimeZoneControl which adds IDispose to the interface
	/// </summary>
	public interface I_OlkTimeZoneControl : Microsoft.Office.Interop.Outlook._OlkTimeZoneControl, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OlkTimeZoneControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkTimeZoneControlEvents which adds IDispose to the interface
	/// </summary>
	public interface IOlkTimeZoneControlEvents : Microsoft.Office.Interop.Outlook.OlkTimeZoneControlEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkTimeZoneControlEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkTimeZoneControlEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOlkTimeZoneControlEvents_Event : Microsoft.Office.Interop.Outlook.OlkTimeZoneControlEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkTimeZoneControlEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OlkTimeZoneControl which adds IDispose to the interface
	/// </summary>
	public interface IOlkTimeZoneControl : Microsoft.Office.Interop.Outlook.OlkTimeZoneControl, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OlkTimeZoneControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ApplicationEvents which adds IDispose to the interface
	/// </summary>
	public interface IApplicationEvents : Microsoft.Office.Interop.Outlook.ApplicationEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ApplicationEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PropertyPages which adds IDispose to the interface
	/// </summary>
	public interface IPropertyPages : Microsoft.Office.Interop.Outlook.PropertyPages, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.PropertyPages Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RecurrencePattern which adds IDispose to the interface
	/// </summary>
	public interface IRecurrencePattern : Microsoft.Office.Interop.Outlook.RecurrencePattern, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.RecurrencePattern Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Exceptions which adds IDispose to the interface
	/// </summary>
	public interface IExceptions : Microsoft.Office.Interop.Outlook.Exceptions, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Exceptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Exception which adds IDispose to the interface
	/// </summary>
	public interface IException : Microsoft.Office.Interop.Outlook.Exception, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Exception Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AppointmentItem which adds IDispose to the interface
	/// </summary>
	public interface IAppointmentItem : Microsoft.Office.Interop.Outlook.AppointmentItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.AppointmentItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _AppointmentItem which adds IDispose to the interface
	/// </summary>
	public interface I_AppointmentItem : Microsoft.Office.Interop.Outlook._AppointmentItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._AppointmentItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MeetingItem which adds IDispose to the interface
	/// </summary>
	public interface IMeetingItem : Microsoft.Office.Interop.Outlook.MeetingItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.MeetingItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _MeetingItem which adds IDispose to the interface
	/// </summary>
	public interface I_MeetingItem : Microsoft.Office.Interop.Outlook._MeetingItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._MeetingItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ExplorersEvents which adds IDispose to the interface
	/// </summary>
	public interface IExplorersEvents : Microsoft.Office.Interop.Outlook.ExplorersEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ExplorersEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FoldersEvents which adds IDispose to the interface
	/// </summary>
	public interface IFoldersEvents : Microsoft.Office.Interop.Outlook.FoldersEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.FoldersEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for InspectorsEvents which adds IDispose to the interface
	/// </summary>
	public interface IInspectorsEvents : Microsoft.Office.Interop.Outlook.InspectorsEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.InspectorsEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ItemsEvents which adds IDispose to the interface
	/// </summary>
	public interface IItemsEvents : Microsoft.Office.Interop.Outlook.ItemsEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ItemsEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for NameSpaceEvents which adds IDispose to the interface
	/// </summary>
	public interface INameSpaceEvents : Microsoft.Office.Interop.Outlook.NameSpaceEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.NameSpaceEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OutlookBarGroup which adds IDispose to the interface
	/// </summary>
	public interface IOutlookBarGroup : Microsoft.Office.Interop.Outlook.OutlookBarGroup, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OutlookBarGroup Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OutlookBarShortcuts which adds IDispose to the interface
	/// </summary>
	public interface I_OutlookBarShortcuts : Microsoft.Office.Interop.Outlook._OutlookBarShortcuts, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OutlookBarShortcuts Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OutlookBarShortcut which adds IDispose to the interface
	/// </summary>
	public interface IOutlookBarShortcut : Microsoft.Office.Interop.Outlook.OutlookBarShortcut, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OutlookBarShortcut Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OutlookBarGroups which adds IDispose to the interface
	/// </summary>
	public interface I_OutlookBarGroups : Microsoft.Office.Interop.Outlook._OutlookBarGroups, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OutlookBarGroups Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OutlookBarGroupsEvents which adds IDispose to the interface
	/// </summary>
	public interface IOutlookBarGroupsEvents : Microsoft.Office.Interop.Outlook.OutlookBarGroupsEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OutlookBarGroupsEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OutlookBarPane which adds IDispose to the interface
	/// </summary>
	public interface I_OutlookBarPane : Microsoft.Office.Interop.Outlook._OutlookBarPane, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OutlookBarPane Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OutlookBarStorage which adds IDispose to the interface
	/// </summary>
	public interface IOutlookBarStorage : Microsoft.Office.Interop.Outlook.OutlookBarStorage, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OutlookBarStorage Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OutlookBarPaneEvents which adds IDispose to the interface
	/// </summary>
	public interface IOutlookBarPaneEvents : Microsoft.Office.Interop.Outlook.OutlookBarPaneEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OutlookBarPaneEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OutlookBarShortcutsEvents which adds IDispose to the interface
	/// </summary>
	public interface IOutlookBarShortcutsEvents : Microsoft.Office.Interop.Outlook.OutlookBarShortcutsEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OutlookBarShortcutsEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PropertyPage which adds IDispose to the interface
	/// </summary>
	public interface IPropertyPage : Microsoft.Office.Interop.Outlook.PropertyPage, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.PropertyPage Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PropertyPageSite which adds IDispose to the interface
	/// </summary>
	public interface IPropertyPageSite : Microsoft.Office.Interop.Outlook.PropertyPageSite, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.PropertyPageSite Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Pages which adds IDispose to the interface
	/// </summary>
	public interface IPages : Microsoft.Office.Interop.Outlook.Pages, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Pages Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ApplicationEvents_10 which adds IDispose to the interface
	/// </summary>
	public interface IApplicationEvents_10 : Microsoft.Office.Interop.Outlook.ApplicationEvents_10, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ApplicationEvents_10 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ApplicationEvents_11 which adds IDispose to the interface
	/// </summary>
	public interface IApplicationEvents_11 : Microsoft.Office.Interop.Outlook.ApplicationEvents_11, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ApplicationEvents_11 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AttachmentSelection which adds IDispose to the interface
	/// </summary>
	public interface IAttachmentSelection : Microsoft.Office.Interop.Outlook.AttachmentSelection, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.AttachmentSelection Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _AttachmentSelection which adds IDispose to the interface
	/// </summary>
	public interface I_AttachmentSelection : Microsoft.Office.Interop.Outlook._AttachmentSelection, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._AttachmentSelection Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MAPIFolderEvents_12_Event which adds IDispose to the interface
	/// </summary>
	public interface IMAPIFolderEvents_12_Event : Microsoft.Office.Interop.Outlook.MAPIFolderEvents_12_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.MAPIFolderEvents_12_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Folder which adds IDispose to the interface
	/// </summary>
	public interface IFolder : Microsoft.Office.Interop.Outlook.Folder, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Folder Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MAPIFolderEvents_12 which adds IDispose to the interface
	/// </summary>
	public interface IMAPIFolderEvents_12 : Microsoft.Office.Interop.Outlook.MAPIFolderEvents_12, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.MAPIFolderEvents_12 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ResultsEvents which adds IDispose to the interface
	/// </summary>
	public interface IResultsEvents : Microsoft.Office.Interop.Outlook.ResultsEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ResultsEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _ViewsEvents which adds IDispose to the interface
	/// </summary>
	public interface I_ViewsEvents : Microsoft.Office.Interop.Outlook._ViewsEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._ViewsEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ReminderCollectionEvents which adds IDispose to the interface
	/// </summary>
	public interface IReminderCollectionEvents : Microsoft.Office.Interop.Outlook.ReminderCollectionEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ReminderCollectionEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _DistListItem which adds IDispose to the interface
	/// </summary>
	public interface I_DistListItem : Microsoft.Office.Interop.Outlook._DistListItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._DistListItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _DocumentItem which adds IDispose to the interface
	/// </summary>
	public interface I_DocumentItem : Microsoft.Office.Interop.Outlook._DocumentItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._DocumentItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _JournalItem which adds IDispose to the interface
	/// </summary>
	public interface I_JournalItem : Microsoft.Office.Interop.Outlook._JournalItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._JournalItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _NoteItem which adds IDispose to the interface
	/// </summary>
	public interface I_NoteItem : Microsoft.Office.Interop.Outlook._NoteItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._NoteItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _PostItem which adds IDispose to the interface
	/// </summary>
	public interface I_PostItem : Microsoft.Office.Interop.Outlook._PostItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._PostItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _RemoteItem which adds IDispose to the interface
	/// </summary>
	public interface I_RemoteItem : Microsoft.Office.Interop.Outlook._RemoteItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._RemoteItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _ReportItem which adds IDispose to the interface
	/// </summary>
	public interface I_ReportItem : Microsoft.Office.Interop.Outlook._ReportItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._ReportItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _TaskItem which adds IDispose to the interface
	/// </summary>
	public interface I_TaskItem : Microsoft.Office.Interop.Outlook._TaskItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._TaskItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TaskItem which adds IDispose to the interface
	/// </summary>
	public interface ITaskItem : Microsoft.Office.Interop.Outlook.TaskItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.TaskItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _TaskRequestAcceptItem which adds IDispose to the interface
	/// </summary>
	public interface I_TaskRequestAcceptItem : Microsoft.Office.Interop.Outlook._TaskRequestAcceptItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._TaskRequestAcceptItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _TaskRequestDeclineItem which adds IDispose to the interface
	/// </summary>
	public interface I_TaskRequestDeclineItem : Microsoft.Office.Interop.Outlook._TaskRequestDeclineItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._TaskRequestDeclineItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _TaskRequestItem which adds IDispose to the interface
	/// </summary>
	public interface I_TaskRequestItem : Microsoft.Office.Interop.Outlook._TaskRequestItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._TaskRequestItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _TaskRequestUpdateItem which adds IDispose to the interface
	/// </summary>
	public interface I_TaskRequestUpdateItem : Microsoft.Office.Interop.Outlook._TaskRequestUpdateItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._TaskRequestUpdateItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _FormRegion which adds IDispose to the interface
	/// </summary>
	public interface I_FormRegion : Microsoft.Office.Interop.Outlook._FormRegion, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._FormRegion Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FormRegionEvents which adds IDispose to the interface
	/// </summary>
	public interface IFormRegionEvents : Microsoft.Office.Interop.Outlook.FormRegionEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.FormRegionEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _TableView which adds IDispose to the interface
	/// </summary>
	public interface I_TableView : Microsoft.Office.Interop.Outlook._TableView, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._TableView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ViewFields which adds IDispose to the interface
	/// </summary>
	public interface IViewFields : Microsoft.Office.Interop.Outlook.ViewFields, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ViewFields Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _ViewFields which adds IDispose to the interface
	/// </summary>
	public interface I_ViewFields : Microsoft.Office.Interop.Outlook._ViewFields, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._ViewFields Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _ViewField which adds IDispose to the interface
	/// </summary>
	public interface I_ViewField : Microsoft.Office.Interop.Outlook._ViewField, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._ViewField Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ColumnFormat which adds IDispose to the interface
	/// </summary>
	public interface IColumnFormat : Microsoft.Office.Interop.Outlook.ColumnFormat, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ColumnFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _ColumnFormat which adds IDispose to the interface
	/// </summary>
	public interface I_ColumnFormat : Microsoft.Office.Interop.Outlook._ColumnFormat, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._ColumnFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ViewField which adds IDispose to the interface
	/// </summary>
	public interface IViewField : Microsoft.Office.Interop.Outlook.ViewField, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ViewField Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OrderFields which adds IDispose to the interface
	/// </summary>
	public interface IOrderFields : Microsoft.Office.Interop.Outlook.OrderFields, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OrderFields Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OrderFields which adds IDispose to the interface
	/// </summary>
	public interface I_OrderFields : Microsoft.Office.Interop.Outlook._OrderFields, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OrderFields Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OrderField which adds IDispose to the interface
	/// </summary>
	public interface I_OrderField : Microsoft.Office.Interop.Outlook._OrderField, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._OrderField Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OrderField which adds IDispose to the interface
	/// </summary>
	public interface IOrderField : Microsoft.Office.Interop.Outlook.OrderField, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OrderField Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ViewFont which adds IDispose to the interface
	/// </summary>
	public interface IViewFont : Microsoft.Office.Interop.Outlook.ViewFont, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ViewFont Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _ViewFont which adds IDispose to the interface
	/// </summary>
	public interface I_ViewFont : Microsoft.Office.Interop.Outlook._ViewFont, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._ViewFont Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AutoFormatRules which adds IDispose to the interface
	/// </summary>
	public interface IAutoFormatRules : Microsoft.Office.Interop.Outlook.AutoFormatRules, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.AutoFormatRules Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _AutoFormatRules which adds IDispose to the interface
	/// </summary>
	public interface I_AutoFormatRules : Microsoft.Office.Interop.Outlook._AutoFormatRules, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._AutoFormatRules Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AutoFormatRule which adds IDispose to the interface
	/// </summary>
	public interface IAutoFormatRule : Microsoft.Office.Interop.Outlook.AutoFormatRule, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.AutoFormatRule Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _AutoFormatRule which adds IDispose to the interface
	/// </summary>
	public interface I_AutoFormatRule : Microsoft.Office.Interop.Outlook._AutoFormatRule, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._AutoFormatRule Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _IconView which adds IDispose to the interface
	/// </summary>
	public interface I_IconView : Microsoft.Office.Interop.Outlook._IconView, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._IconView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CardView which adds IDispose to the interface
	/// </summary>
	public interface I_CardView : Microsoft.Office.Interop.Outlook._CardView, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._CardView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CalendarView which adds IDispose to the interface
	/// </summary>
	public interface I_CalendarView : Microsoft.Office.Interop.Outlook._CalendarView, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._CalendarView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _TimelineView which adds IDispose to the interface
	/// </summary>
	public interface I_TimelineView : Microsoft.Office.Interop.Outlook._TimelineView, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._TimelineView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _MailModule which adds IDispose to the interface
	/// </summary>
	public interface I_MailModule : Microsoft.Office.Interop.Outlook._MailModule, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._MailModule Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _NavigationGroups which adds IDispose to the interface
	/// </summary>
	public interface I_NavigationGroups : Microsoft.Office.Interop.Outlook._NavigationGroups, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._NavigationGroups Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _NavigationGroup which adds IDispose to the interface
	/// </summary>
	public interface I_NavigationGroup : Microsoft.Office.Interop.Outlook._NavigationGroup, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._NavigationGroup Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for NavigationFolders which adds IDispose to the interface
	/// </summary>
	public interface INavigationFolders : Microsoft.Office.Interop.Outlook.NavigationFolders, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.NavigationFolders Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _NavigationFolders which adds IDispose to the interface
	/// </summary>
	public interface I_NavigationFolders : Microsoft.Office.Interop.Outlook._NavigationFolders, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._NavigationFolders Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _NavigationFolder which adds IDispose to the interface
	/// </summary>
	public interface I_NavigationFolder : Microsoft.Office.Interop.Outlook._NavigationFolder, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._NavigationFolder Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for NavigationFolder which adds IDispose to the interface
	/// </summary>
	public interface INavigationFolder : Microsoft.Office.Interop.Outlook.NavigationFolder, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.NavigationFolder Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for NavigationGroup which adds IDispose to the interface
	/// </summary>
	public interface INavigationGroup : Microsoft.Office.Interop.Outlook.NavigationGroup, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.NavigationGroup Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CalendarModule which adds IDispose to the interface
	/// </summary>
	public interface I_CalendarModule : Microsoft.Office.Interop.Outlook._CalendarModule, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._CalendarModule Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _ContactsModule which adds IDispose to the interface
	/// </summary>
	public interface I_ContactsModule : Microsoft.Office.Interop.Outlook._ContactsModule, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._ContactsModule Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _TasksModule which adds IDispose to the interface
	/// </summary>
	public interface I_TasksModule : Microsoft.Office.Interop.Outlook._TasksModule, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._TasksModule Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _JournalModule which adds IDispose to the interface
	/// </summary>
	public interface I_JournalModule : Microsoft.Office.Interop.Outlook._JournalModule, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._JournalModule Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _NotesModule which adds IDispose to the interface
	/// </summary>
	public interface I_NotesModule : Microsoft.Office.Interop.Outlook._NotesModule, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._NotesModule Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for NavigationPaneEvents_12 which adds IDispose to the interface
	/// </summary>
	public interface INavigationPaneEvents_12 : Microsoft.Office.Interop.Outlook.NavigationPaneEvents_12, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.NavigationPaneEvents_12 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for NavigationGroupsEvents_12 which adds IDispose to the interface
	/// </summary>
	public interface INavigationGroupsEvents_12 : Microsoft.Office.Interop.Outlook.NavigationGroupsEvents_12, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.NavigationGroupsEvents_12 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _BusinessCardView which adds IDispose to the interface
	/// </summary>
	public interface I_BusinessCardView : Microsoft.Office.Interop.Outlook._BusinessCardView, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._BusinessCardView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _FormRegionStartup which adds IDispose to the interface
	/// </summary>
	public interface I_FormRegionStartup : Microsoft.Office.Interop.Outlook._FormRegionStartup, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._FormRegionStartup Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FormRegionEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IFormRegionEvents_Event : Microsoft.Office.Interop.Outlook.FormRegionEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.FormRegionEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FormRegion which adds IDispose to the interface
	/// </summary>
	public interface IFormRegion : Microsoft.Office.Interop.Outlook.FormRegion, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.FormRegion Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ApplicationEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IApplicationEvents_Event : Microsoft.Office.Interop.Outlook.ApplicationEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ApplicationEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ApplicationEvents_10_Event which adds IDispose to the interface
	/// </summary>
	public interface IApplicationEvents_10_Event : Microsoft.Office.Interop.Outlook.ApplicationEvents_10_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ApplicationEvents_10_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ApplicationEvents_11_Event which adds IDispose to the interface
	/// </summary>
	public interface IApplicationEvents_11_Event : Microsoft.Office.Interop.Outlook.ApplicationEvents_11_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ApplicationEvents_11_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Application which adds IDispose to the interface
	/// </summary>
	public interface IApplication : Microsoft.Office.Interop.Outlook.Application, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Application Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DistListItem which adds IDispose to the interface
	/// </summary>
	public interface IDistListItem : Microsoft.Office.Interop.Outlook.DistListItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.DistListItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DocumentItem which adds IDispose to the interface
	/// </summary>
	public interface IDocumentItem : Microsoft.Office.Interop.Outlook.DocumentItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.DocumentItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ExplorersEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IExplorersEvents_Event : Microsoft.Office.Interop.Outlook.ExplorersEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ExplorersEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Explorers which adds IDispose to the interface
	/// </summary>
	public interface IExplorers : Microsoft.Office.Interop.Outlook.Explorers, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Explorers Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for InspectorsEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IInspectorsEvents_Event : Microsoft.Office.Interop.Outlook.InspectorsEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.InspectorsEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Inspectors which adds IDispose to the interface
	/// </summary>
	public interface IInspectors : Microsoft.Office.Interop.Outlook.Inspectors, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Inspectors Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FoldersEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IFoldersEvents_Event : Microsoft.Office.Interop.Outlook.FoldersEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.FoldersEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Folders which adds IDispose to the interface
	/// </summary>
	public interface IFolders : Microsoft.Office.Interop.Outlook.Folders, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Folders Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ItemsEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IItemsEvents_Event : Microsoft.Office.Interop.Outlook.ItemsEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ItemsEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Items which adds IDispose to the interface
	/// </summary>
	public interface IItems : Microsoft.Office.Interop.Outlook.Items, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Items Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for JournalItem which adds IDispose to the interface
	/// </summary>
	public interface IJournalItem : Microsoft.Office.Interop.Outlook.JournalItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.JournalItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for NameSpaceEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface INameSpaceEvents_Event : Microsoft.Office.Interop.Outlook.NameSpaceEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.NameSpaceEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for NameSpace which adds IDispose to the interface
	/// </summary>
	public interface INameSpace : Microsoft.Office.Interop.Outlook.NameSpace, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.NameSpace Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for NoteItem which adds IDispose to the interface
	/// </summary>
	public interface INoteItem : Microsoft.Office.Interop.Outlook.NoteItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.NoteItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OutlookBarGroupsEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOutlookBarGroupsEvents_Event : Microsoft.Office.Interop.Outlook.OutlookBarGroupsEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OutlookBarGroupsEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OutlookBarGroups which adds IDispose to the interface
	/// </summary>
	public interface IOutlookBarGroups : Microsoft.Office.Interop.Outlook.OutlookBarGroups, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OutlookBarGroups Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OutlookBarPaneEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOutlookBarPaneEvents_Event : Microsoft.Office.Interop.Outlook.OutlookBarPaneEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OutlookBarPaneEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OutlookBarPane which adds IDispose to the interface
	/// </summary>
	public interface IOutlookBarPane : Microsoft.Office.Interop.Outlook.OutlookBarPane, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OutlookBarPane Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OutlookBarShortcutsEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOutlookBarShortcutsEvents_Event : Microsoft.Office.Interop.Outlook.OutlookBarShortcutsEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OutlookBarShortcutsEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OutlookBarShortcuts which adds IDispose to the interface
	/// </summary>
	public interface IOutlookBarShortcuts : Microsoft.Office.Interop.Outlook.OutlookBarShortcuts, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.OutlookBarShortcuts Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PostItem which adds IDispose to the interface
	/// </summary>
	public interface IPostItem : Microsoft.Office.Interop.Outlook.PostItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.PostItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RemoteItem which adds IDispose to the interface
	/// </summary>
	public interface IRemoteItem : Microsoft.Office.Interop.Outlook.RemoteItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.RemoteItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ReportItem which adds IDispose to the interface
	/// </summary>
	public interface IReportItem : Microsoft.Office.Interop.Outlook.ReportItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ReportItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TaskRequestAcceptItem which adds IDispose to the interface
	/// </summary>
	public interface ITaskRequestAcceptItem : Microsoft.Office.Interop.Outlook.TaskRequestAcceptItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.TaskRequestAcceptItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TaskRequestDeclineItem which adds IDispose to the interface
	/// </summary>
	public interface ITaskRequestDeclineItem : Microsoft.Office.Interop.Outlook.TaskRequestDeclineItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.TaskRequestDeclineItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TaskRequestItem which adds IDispose to the interface
	/// </summary>
	public interface ITaskRequestItem : Microsoft.Office.Interop.Outlook.TaskRequestItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.TaskRequestItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TaskRequestUpdateItem which adds IDispose to the interface
	/// </summary>
	public interface ITaskRequestUpdateItem : Microsoft.Office.Interop.Outlook.TaskRequestUpdateItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.TaskRequestUpdateItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ResultsEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IResultsEvents_Event : Microsoft.Office.Interop.Outlook.ResultsEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ResultsEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Results which adds IDispose to the interface
	/// </summary>
	public interface IResults : Microsoft.Office.Interop.Outlook.Results, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Results Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _ViewsEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface I_ViewsEvents_Event : Microsoft.Office.Interop.Outlook._ViewsEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook._ViewsEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Views which adds IDispose to the interface
	/// </summary>
	public interface IViews : Microsoft.Office.Interop.Outlook.Views, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Views Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Reminder which adds IDispose to the interface
	/// </summary>
	public interface IReminder : Microsoft.Office.Interop.Outlook.Reminder, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Reminder Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ReminderCollectionEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IReminderCollectionEvents_Event : Microsoft.Office.Interop.Outlook.ReminderCollectionEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ReminderCollectionEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Reminders which adds IDispose to the interface
	/// </summary>
	public interface IReminders : Microsoft.Office.Interop.Outlook.Reminders, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.Reminders Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for StorageItem which adds IDispose to the interface
	/// </summary>
	public interface IStorageItem : Microsoft.Office.Interop.Outlook.StorageItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.StorageItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for NavigationPaneEvents_12_Event which adds IDispose to the interface
	/// </summary>
	public interface INavigationPaneEvents_12_Event : Microsoft.Office.Interop.Outlook.NavigationPaneEvents_12_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.NavigationPaneEvents_12_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for NavigationPane which adds IDispose to the interface
	/// </summary>
	public interface INavigationPane : Microsoft.Office.Interop.Outlook.NavigationPane, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.NavigationPane Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for NavigationGroupsEvents_12_Event which adds IDispose to the interface
	/// </summary>
	public interface INavigationGroupsEvents_12_Event : Microsoft.Office.Interop.Outlook.NavigationGroupsEvents_12_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.NavigationGroupsEvents_12_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for NavigationGroups which adds IDispose to the interface
	/// </summary>
	public interface INavigationGroups : Microsoft.Office.Interop.Outlook.NavigationGroups, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.NavigationGroups Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DoNotUseMeFolder which adds IDispose to the interface
	/// </summary>
	public interface IDoNotUseMeFolder : Microsoft.Office.Interop.Outlook.DoNotUseMeFolder, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.DoNotUseMeFolder Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TimelineView which adds IDispose to the interface
	/// </summary>
	public interface ITimelineView : Microsoft.Office.Interop.Outlook.TimelineView, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.TimelineView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MailModule which adds IDispose to the interface
	/// </summary>
	public interface IMailModule : Microsoft.Office.Interop.Outlook.MailModule, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.MailModule Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CalendarModule which adds IDispose to the interface
	/// </summary>
	public interface ICalendarModule : Microsoft.Office.Interop.Outlook.CalendarModule, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.CalendarModule Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ContactsModule which adds IDispose to the interface
	/// </summary>
	public interface IContactsModule : Microsoft.Office.Interop.Outlook.ContactsModule, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.ContactsModule Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TasksModule which adds IDispose to the interface
	/// </summary>
	public interface ITasksModule : Microsoft.Office.Interop.Outlook.TasksModule, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.TasksModule Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for JournalModule which adds IDispose to the interface
	/// </summary>
	public interface IJournalModule : Microsoft.Office.Interop.Outlook.JournalModule, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.JournalModule Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for NotesModule which adds IDispose to the interface
	/// </summary>
	public interface INotesModule : Microsoft.Office.Interop.Outlook.NotesModule, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.NotesModule Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TableView which adds IDispose to the interface
	/// </summary>
	public interface ITableView : Microsoft.Office.Interop.Outlook.TableView, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.TableView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IconView which adds IDispose to the interface
	/// </summary>
	public interface IIconView : Microsoft.Office.Interop.Outlook.IconView, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.IconView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CardView which adds IDispose to the interface
	/// </summary>
	public interface ICardView : Microsoft.Office.Interop.Outlook.CardView, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.CardView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CalendarView which adds IDispose to the interface
	/// </summary>
	public interface ICalendarView : Microsoft.Office.Interop.Outlook.CalendarView, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.CalendarView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for BusinessCardView which adds IDispose to the interface
	/// </summary>
	public interface IBusinessCardView : Microsoft.Office.Interop.Outlook.BusinessCardView, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.BusinessCardView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FormRegionStartup which adds IDispose to the interface
	/// </summary>
	public interface IFormRegionStartup : Microsoft.Office.Interop.Outlook.FormRegionStartup, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.FormRegionStartup Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TimeZone which adds IDispose to the interface
	/// </summary>
	public interface ITimeZone : Microsoft.Office.Interop.Outlook.TimeZone, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Outlook.TimeZone Resource { get; }
	}

}