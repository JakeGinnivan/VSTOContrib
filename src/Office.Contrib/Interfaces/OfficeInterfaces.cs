
//office, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c
namespace Office.Contrib.Interfaces
{
	/// <summary>
	/// Wrapper interface for IAccessible which adds IDispose to the interface
	/// </summary>
	public interface IIAccessible : Microsoft.Office.Core.IAccessible, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IAccessible Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _IMsoDispObj which adds IDispose to the interface
	/// </summary>
	public interface I_IMsoDispObj : Microsoft.Office.Core._IMsoDispObj, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._IMsoDispObj Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _IMsoOleAccDispObj which adds IDispose to the interface
	/// </summary>
	public interface I_IMsoOleAccDispObj : Microsoft.Office.Core._IMsoOleAccDispObj, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._IMsoOleAccDispObj Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CommandBars which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBars : Microsoft.Office.Core._CommandBars, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._CommandBars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CommandBar which adds IDispose to the interface
	/// </summary>
	public interface ICommandBar : Microsoft.Office.Core.CommandBar, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.CommandBar Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CommandBarControls which adds IDispose to the interface
	/// </summary>
	public interface ICommandBarControls : Microsoft.Office.Core.CommandBarControls, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.CommandBarControls Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CommandBarControl which adds IDispose to the interface
	/// </summary>
	public interface ICommandBarControl : Microsoft.Office.Core.CommandBarControl, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.CommandBarControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CommandBarButton which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarButton : Microsoft.Office.Core._CommandBarButton, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._CommandBarButton Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CommandBarPopup which adds IDispose to the interface
	/// </summary>
	public interface ICommandBarPopup : Microsoft.Office.Core.CommandBarPopup, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.CommandBarPopup Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CommandBarComboBox which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarComboBox : Microsoft.Office.Core._CommandBarComboBox, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._CommandBarComboBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CommandBarActiveX which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarActiveX : Microsoft.Office.Core._CommandBarActiveX, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._CommandBarActiveX Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Adjustments which adds IDispose to the interface
	/// </summary>
	public interface IAdjustments : Microsoft.Office.Core.Adjustments, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.Adjustments Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CalloutFormat which adds IDispose to the interface
	/// </summary>
	public interface ICalloutFormat : Microsoft.Office.Core.CalloutFormat, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.CalloutFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ColorFormat which adds IDispose to the interface
	/// </summary>
	public interface IColorFormat : Microsoft.Office.Core.ColorFormat, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ColorFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ConnectorFormat which adds IDispose to the interface
	/// </summary>
	public interface IConnectorFormat : Microsoft.Office.Core.ConnectorFormat, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ConnectorFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FillFormat which adds IDispose to the interface
	/// </summary>
	public interface IFillFormat : Microsoft.Office.Core.FillFormat, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.FillFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FreeformBuilder which adds IDispose to the interface
	/// </summary>
	public interface IFreeformBuilder : Microsoft.Office.Core.FreeformBuilder, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.FreeformBuilder Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for GroupShapes which adds IDispose to the interface
	/// </summary>
	public interface IGroupShapes : Microsoft.Office.Core.GroupShapes, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.GroupShapes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LineFormat which adds IDispose to the interface
	/// </summary>
	public interface ILineFormat : Microsoft.Office.Core.LineFormat, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.LineFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ShapeNode which adds IDispose to the interface
	/// </summary>
	public interface IShapeNode : Microsoft.Office.Core.ShapeNode, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ShapeNode Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ShapeNodes which adds IDispose to the interface
	/// </summary>
	public interface IShapeNodes : Microsoft.Office.Core.ShapeNodes, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ShapeNodes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PictureFormat which adds IDispose to the interface
	/// </summary>
	public interface IPictureFormat : Microsoft.Office.Core.PictureFormat, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.PictureFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ShadowFormat which adds IDispose to the interface
	/// </summary>
	public interface IShadowFormat : Microsoft.Office.Core.ShadowFormat, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ShadowFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Script which adds IDispose to the interface
	/// </summary>
	public interface IScript : Microsoft.Office.Core.Script, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.Script Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Scripts which adds IDispose to the interface
	/// </summary>
	public interface IScripts : Microsoft.Office.Core.Scripts, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.Scripts Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Shape which adds IDispose to the interface
	/// </summary>
	public interface IShape : Microsoft.Office.Core.Shape, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.Shape Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ShapeRange which adds IDispose to the interface
	/// </summary>
	public interface IShapeRange : Microsoft.Office.Core.ShapeRange, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ShapeRange Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Shapes which adds IDispose to the interface
	/// </summary>
	public interface IShapes : Microsoft.Office.Core.Shapes, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.Shapes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextEffectFormat which adds IDispose to the interface
	/// </summary>
	public interface ITextEffectFormat : Microsoft.Office.Core.TextEffectFormat, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.TextEffectFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextFrame which adds IDispose to the interface
	/// </summary>
	public interface ITextFrame : Microsoft.Office.Core.TextFrame, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.TextFrame Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ThreeDFormat which adds IDispose to the interface
	/// </summary>
	public interface IThreeDFormat : Microsoft.Office.Core.ThreeDFormat, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ThreeDFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoDispCagNotifySink which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDispCagNotifySink : Microsoft.Office.Core.IMsoDispCagNotifySink, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoDispCagNotifySink Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Balloon which adds IDispose to the interface
	/// </summary>
	public interface IBalloon : Microsoft.Office.Core.Balloon, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.Balloon Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for BalloonCheckboxes which adds IDispose to the interface
	/// </summary>
	public interface IBalloonCheckboxes : Microsoft.Office.Core.BalloonCheckboxes, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.BalloonCheckboxes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for BalloonCheckbox which adds IDispose to the interface
	/// </summary>
	public interface IBalloonCheckbox : Microsoft.Office.Core.BalloonCheckbox, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.BalloonCheckbox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for BalloonLabels which adds IDispose to the interface
	/// </summary>
	public interface IBalloonLabels : Microsoft.Office.Core.BalloonLabels, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.BalloonLabels Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for BalloonLabel which adds IDispose to the interface
	/// </summary>
	public interface IBalloonLabel : Microsoft.Office.Core.BalloonLabel, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.BalloonLabel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AnswerWizardFiles which adds IDispose to the interface
	/// </summary>
	public interface IAnswerWizardFiles : Microsoft.Office.Core.AnswerWizardFiles, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.AnswerWizardFiles Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AnswerWizard which adds IDispose to the interface
	/// </summary>
	public interface IAnswerWizard : Microsoft.Office.Core.AnswerWizard, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.AnswerWizard Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Assistant which adds IDispose to the interface
	/// </summary>
	public interface IAssistant : Microsoft.Office.Core.Assistant, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.Assistant Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DocumentProperty which adds IDispose to the interface
	/// </summary>
	public interface IDocumentProperty : Microsoft.Office.Core.DocumentProperty, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.DocumentProperty Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DocumentProperties which adds IDispose to the interface
	/// </summary>
	public interface IDocumentProperties : Microsoft.Office.Core.DocumentProperties, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.DocumentProperties Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IFoundFiles which adds IDispose to the interface
	/// </summary>
	public interface IIFoundFiles : Microsoft.Office.Core.IFoundFiles, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IFoundFiles Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IFind which adds IDispose to the interface
	/// </summary>
	public interface IIFind : Microsoft.Office.Core.IFind, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IFind Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FoundFiles which adds IDispose to the interface
	/// </summary>
	public interface IFoundFiles : Microsoft.Office.Core.FoundFiles, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.FoundFiles Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PropertyTest which adds IDispose to the interface
	/// </summary>
	public interface IPropertyTest : Microsoft.Office.Core.PropertyTest, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.PropertyTest Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PropertyTests which adds IDispose to the interface
	/// </summary>
	public interface IPropertyTests : Microsoft.Office.Core.PropertyTests, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.PropertyTests Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FileSearch which adds IDispose to the interface
	/// </summary>
	public interface IFileSearch : Microsoft.Office.Core.FileSearch, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.FileSearch Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for COMAddIn which adds IDispose to the interface
	/// </summary>
	public interface ICOMAddIn : Microsoft.Office.Core.COMAddIn, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.COMAddIn Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for COMAddIns which adds IDispose to the interface
	/// </summary>
	public interface ICOMAddIns : Microsoft.Office.Core.COMAddIns, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.COMAddIns Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LanguageSettings which adds IDispose to the interface
	/// </summary>
	public interface ILanguageSettings : Microsoft.Office.Core.LanguageSettings, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.LanguageSettings Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICommandBarsEvents which adds IDispose to the interface
	/// </summary>
	public interface IICommandBarsEvents : Microsoft.Office.Core.ICommandBarsEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ICommandBarsEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CommandBarsEvents which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarsEvents : Microsoft.Office.Core._CommandBarsEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._CommandBarsEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CommandBarsEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarsEvents_Event : Microsoft.Office.Core._CommandBarsEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._CommandBarsEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CommandBars which adds IDispose to the interface
	/// </summary>
	public interface ICommandBars : Microsoft.Office.Core.CommandBars, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.CommandBars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICommandBarComboBoxEvents which adds IDispose to the interface
	/// </summary>
	public interface IICommandBarComboBoxEvents : Microsoft.Office.Core.ICommandBarComboBoxEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ICommandBarComboBoxEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CommandBarComboBoxEvents which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarComboBoxEvents : Microsoft.Office.Core._CommandBarComboBoxEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._CommandBarComboBoxEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CommandBarComboBoxEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarComboBoxEvents_Event : Microsoft.Office.Core._CommandBarComboBoxEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._CommandBarComboBoxEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CommandBarComboBox which adds IDispose to the interface
	/// </summary>
	public interface ICommandBarComboBox : Microsoft.Office.Core.CommandBarComboBox, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.CommandBarComboBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICommandBarButtonEvents which adds IDispose to the interface
	/// </summary>
	public interface IICommandBarButtonEvents : Microsoft.Office.Core.ICommandBarButtonEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ICommandBarButtonEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CommandBarButtonEvents which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarButtonEvents : Microsoft.Office.Core._CommandBarButtonEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._CommandBarButtonEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CommandBarButtonEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarButtonEvents_Event : Microsoft.Office.Core._CommandBarButtonEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._CommandBarButtonEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CommandBarButton which adds IDispose to the interface
	/// </summary>
	public interface ICommandBarButton : Microsoft.Office.Core.CommandBarButton, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.CommandBarButton Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for WebPageFont which adds IDispose to the interface
	/// </summary>
	public interface IWebPageFont : Microsoft.Office.Core.WebPageFont, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.WebPageFont Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for WebPageFonts which adds IDispose to the interface
	/// </summary>
	public interface IWebPageFonts : Microsoft.Office.Core.WebPageFonts, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.WebPageFonts Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for HTMLProjectItem which adds IDispose to the interface
	/// </summary>
	public interface IHTMLProjectItem : Microsoft.Office.Core.HTMLProjectItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.HTMLProjectItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for HTMLProjectItems which adds IDispose to the interface
	/// </summary>
	public interface IHTMLProjectItems : Microsoft.Office.Core.HTMLProjectItems, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.HTMLProjectItems Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for HTMLProject which adds IDispose to the interface
	/// </summary>
	public interface IHTMLProject : Microsoft.Office.Core.HTMLProject, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.HTMLProject Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MsoDebugOptions which adds IDispose to the interface
	/// </summary>
	public interface IMsoDebugOptions : Microsoft.Office.Core.MsoDebugOptions, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.MsoDebugOptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FileDialogSelectedItems which adds IDispose to the interface
	/// </summary>
	public interface IFileDialogSelectedItems : Microsoft.Office.Core.FileDialogSelectedItems, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.FileDialogSelectedItems Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FileDialogFilter which adds IDispose to the interface
	/// </summary>
	public interface IFileDialogFilter : Microsoft.Office.Core.FileDialogFilter, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.FileDialogFilter Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FileDialogFilters which adds IDispose to the interface
	/// </summary>
	public interface IFileDialogFilters : Microsoft.Office.Core.FileDialogFilters, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.FileDialogFilters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FileDialog which adds IDispose to the interface
	/// </summary>
	public interface IFileDialog : Microsoft.Office.Core.FileDialog, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.FileDialog Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SignatureSet which adds IDispose to the interface
	/// </summary>
	public interface ISignatureSet : Microsoft.Office.Core.SignatureSet, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SignatureSet Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Signature which adds IDispose to the interface
	/// </summary>
	public interface ISignature : Microsoft.Office.Core.Signature, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.Signature Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoEnvelopeVB which adds IDispose to the interface
	/// </summary>
	public interface IIMsoEnvelopeVB : Microsoft.Office.Core.IMsoEnvelopeVB, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoEnvelopeVB Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoEnvelopeVBEvents which adds IDispose to the interface
	/// </summary>
	public interface IIMsoEnvelopeVBEvents : Microsoft.Office.Core.IMsoEnvelopeVBEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoEnvelopeVBEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoEnvelopeVBEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IIMsoEnvelopeVBEvents_Event : Microsoft.Office.Core.IMsoEnvelopeVBEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoEnvelopeVBEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MsoEnvelope which adds IDispose to the interface
	/// </summary>
	public interface IMsoEnvelope : Microsoft.Office.Core.MsoEnvelope, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.MsoEnvelope Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FileTypes which adds IDispose to the interface
	/// </summary>
	public interface IFileTypes : Microsoft.Office.Core.FileTypes, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.FileTypes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SearchFolders which adds IDispose to the interface
	/// </summary>
	public interface ISearchFolders : Microsoft.Office.Core.SearchFolders, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SearchFolders Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ScopeFolders which adds IDispose to the interface
	/// </summary>
	public interface IScopeFolders : Microsoft.Office.Core.ScopeFolders, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ScopeFolders Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ScopeFolder which adds IDispose to the interface
	/// </summary>
	public interface IScopeFolder : Microsoft.Office.Core.ScopeFolder, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ScopeFolder Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SearchScope which adds IDispose to the interface
	/// </summary>
	public interface ISearchScope : Microsoft.Office.Core.SearchScope, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SearchScope Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SearchScopes which adds IDispose to the interface
	/// </summary>
	public interface ISearchScopes : Microsoft.Office.Core.SearchScopes, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SearchScopes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoDiagram which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDiagram : Microsoft.Office.Core.IMsoDiagram, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoDiagram Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DiagramNodes which adds IDispose to the interface
	/// </summary>
	public interface IDiagramNodes : Microsoft.Office.Core.DiagramNodes, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.DiagramNodes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DiagramNodeChildren which adds IDispose to the interface
	/// </summary>
	public interface IDiagramNodeChildren : Microsoft.Office.Core.DiagramNodeChildren, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.DiagramNodeChildren Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DiagramNode which adds IDispose to the interface
	/// </summary>
	public interface IDiagramNode : Microsoft.Office.Core.DiagramNode, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.DiagramNode Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CanvasShapes which adds IDispose to the interface
	/// </summary>
	public interface ICanvasShapes : Microsoft.Office.Core.CanvasShapes, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.CanvasShapes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OfficeDataSourceObject which adds IDispose to the interface
	/// </summary>
	public interface IOfficeDataSourceObject : Microsoft.Office.Core.OfficeDataSourceObject, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.OfficeDataSourceObject Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ODSOColumn which adds IDispose to the interface
	/// </summary>
	public interface IODSOColumn : Microsoft.Office.Core.ODSOColumn, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ODSOColumn Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ODSOColumns which adds IDispose to the interface
	/// </summary>
	public interface IODSOColumns : Microsoft.Office.Core.ODSOColumns, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ODSOColumns Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ODSOFilter which adds IDispose to the interface
	/// </summary>
	public interface IODSOFilter : Microsoft.Office.Core.ODSOFilter, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ODSOFilter Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ODSOFilters which adds IDispose to the interface
	/// </summary>
	public interface IODSOFilters : Microsoft.Office.Core.ODSOFilters, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ODSOFilters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for NewFile which adds IDispose to the interface
	/// </summary>
	public interface INewFile : Microsoft.Office.Core.NewFile, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.NewFile Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for WebComponent which adds IDispose to the interface
	/// </summary>
	public interface IWebComponent : Microsoft.Office.Core.WebComponent, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.WebComponent Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for WebComponentWindowExternal which adds IDispose to the interface
	/// </summary>
	public interface IWebComponentWindowExternal : Microsoft.Office.Core.WebComponentWindowExternal, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.WebComponentWindowExternal Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for WebComponentFormat which adds IDispose to the interface
	/// </summary>
	public interface IWebComponentFormat : Microsoft.Office.Core.WebComponentFormat, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.WebComponentFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ILicWizExternal which adds IDispose to the interface
	/// </summary>
	public interface IILicWizExternal : Microsoft.Office.Core.ILicWizExternal, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ILicWizExternal Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ILicValidator which adds IDispose to the interface
	/// </summary>
	public interface IILicValidator : Microsoft.Office.Core.ILicValidator, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ILicValidator Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ILicAgent which adds IDispose to the interface
	/// </summary>
	public interface IILicAgent : Microsoft.Office.Core.ILicAgent, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ILicAgent Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoEServicesDialog which adds IDispose to the interface
	/// </summary>
	public interface IIMsoEServicesDialog : Microsoft.Office.Core.IMsoEServicesDialog, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoEServicesDialog Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for WebComponentProperties which adds IDispose to the interface
	/// </summary>
	public interface IWebComponentProperties : Microsoft.Office.Core.WebComponentProperties, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.WebComponentProperties Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartDocument which adds IDispose to the interface
	/// </summary>
	public interface ISmartDocument : Microsoft.Office.Core.SmartDocument, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SmartDocument Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SharedWorkspaceMember which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceMember : Microsoft.Office.Core.SharedWorkspaceMember, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SharedWorkspaceMember Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SharedWorkspaceMembers which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceMembers : Microsoft.Office.Core.SharedWorkspaceMembers, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SharedWorkspaceMembers Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SharedWorkspaceTask which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceTask : Microsoft.Office.Core.SharedWorkspaceTask, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SharedWorkspaceTask Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SharedWorkspaceTasks which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceTasks : Microsoft.Office.Core.SharedWorkspaceTasks, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SharedWorkspaceTasks Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SharedWorkspaceFile which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceFile : Microsoft.Office.Core.SharedWorkspaceFile, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SharedWorkspaceFile Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SharedWorkspaceFiles which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceFiles : Microsoft.Office.Core.SharedWorkspaceFiles, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SharedWorkspaceFiles Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SharedWorkspaceFolder which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceFolder : Microsoft.Office.Core.SharedWorkspaceFolder, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SharedWorkspaceFolder Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SharedWorkspaceFolders which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceFolders : Microsoft.Office.Core.SharedWorkspaceFolders, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SharedWorkspaceFolders Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SharedWorkspaceLink which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceLink : Microsoft.Office.Core.SharedWorkspaceLink, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SharedWorkspaceLink Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SharedWorkspaceLinks which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceLinks : Microsoft.Office.Core.SharedWorkspaceLinks, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SharedWorkspaceLinks Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SharedWorkspace which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspace : Microsoft.Office.Core.SharedWorkspace, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SharedWorkspace Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Sync which adds IDispose to the interface
	/// </summary>
	public interface ISync : Microsoft.Office.Core.Sync, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.Sync Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DocumentLibraryVersion which adds IDispose to the interface
	/// </summary>
	public interface IDocumentLibraryVersion : Microsoft.Office.Core.DocumentLibraryVersion, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.DocumentLibraryVersion Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DocumentLibraryVersions which adds IDispose to the interface
	/// </summary>
	public interface IDocumentLibraryVersions : Microsoft.Office.Core.DocumentLibraryVersions, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.DocumentLibraryVersions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for UserPermission which adds IDispose to the interface
	/// </summary>
	public interface IUserPermission : Microsoft.Office.Core.UserPermission, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.UserPermission Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Permission which adds IDispose to the interface
	/// </summary>
	public interface IPermission : Microsoft.Office.Core.Permission, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.Permission Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MsoDebugOptions_UTRunResult which adds IDispose to the interface
	/// </summary>
	public interface IMsoDebugOptions_UTRunResult : Microsoft.Office.Core.MsoDebugOptions_UTRunResult, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.MsoDebugOptions_UTRunResult Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MsoDebugOptions_UT which adds IDispose to the interface
	/// </summary>
	public interface IMsoDebugOptions_UT : Microsoft.Office.Core.MsoDebugOptions_UT, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.MsoDebugOptions_UT Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MsoDebugOptions_UTs which adds IDispose to the interface
	/// </summary>
	public interface IMsoDebugOptions_UTs : Microsoft.Office.Core.MsoDebugOptions_UTs, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.MsoDebugOptions_UTs Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MsoDebugOptions_UTManager which adds IDispose to the interface
	/// </summary>
	public interface IMsoDebugOptions_UTManager : Microsoft.Office.Core.MsoDebugOptions_UTManager, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.MsoDebugOptions_UTManager Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MetaProperty which adds IDispose to the interface
	/// </summary>
	public interface IMetaProperty : Microsoft.Office.Core.MetaProperty, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.MetaProperty Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MetaProperties which adds IDispose to the interface
	/// </summary>
	public interface IMetaProperties : Microsoft.Office.Core.MetaProperties, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.MetaProperties Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PolicyItem which adds IDispose to the interface
	/// </summary>
	public interface IPolicyItem : Microsoft.Office.Core.PolicyItem, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.PolicyItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ServerPolicy which adds IDispose to the interface
	/// </summary>
	public interface IServerPolicy : Microsoft.Office.Core.ServerPolicy, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ServerPolicy Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DocumentInspector which adds IDispose to the interface
	/// </summary>
	public interface IDocumentInspector : Microsoft.Office.Core.DocumentInspector, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.DocumentInspector Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DocumentInspectors which adds IDispose to the interface
	/// </summary>
	public interface IDocumentInspectors : Microsoft.Office.Core.DocumentInspectors, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.DocumentInspectors Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for WorkflowTask which adds IDispose to the interface
	/// </summary>
	public interface IWorkflowTask : Microsoft.Office.Core.WorkflowTask, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.WorkflowTask Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for WorkflowTasks which adds IDispose to the interface
	/// </summary>
	public interface IWorkflowTasks : Microsoft.Office.Core.WorkflowTasks, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.WorkflowTasks Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for WorkflowTemplate which adds IDispose to the interface
	/// </summary>
	public interface IWorkflowTemplate : Microsoft.Office.Core.WorkflowTemplate, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.WorkflowTemplate Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for WorkflowTemplates which adds IDispose to the interface
	/// </summary>
	public interface IWorkflowTemplates : Microsoft.Office.Core.WorkflowTemplates, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.WorkflowTemplates Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDocumentInspector which adds IDispose to the interface
	/// </summary>
	public interface IIDocumentInspector : Microsoft.Office.Core.IDocumentInspector, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IDocumentInspector Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SignatureSetup which adds IDispose to the interface
	/// </summary>
	public interface ISignatureSetup : Microsoft.Office.Core.SignatureSetup, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SignatureSetup Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SignatureInfo which adds IDispose to the interface
	/// </summary>
	public interface ISignatureInfo : Microsoft.Office.Core.SignatureInfo, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SignatureInfo Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SignatureProvider which adds IDispose to the interface
	/// </summary>
	public interface ISignatureProvider : Microsoft.Office.Core.SignatureProvider, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SignatureProvider Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomXMLPrefixMapping which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLPrefixMapping : Microsoft.Office.Core.CustomXMLPrefixMapping, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.CustomXMLPrefixMapping Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomXMLPrefixMappings which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLPrefixMappings : Microsoft.Office.Core.CustomXMLPrefixMappings, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.CustomXMLPrefixMappings Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomXMLSchema which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLSchema : Microsoft.Office.Core.CustomXMLSchema, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.CustomXMLSchema Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CustomXMLSchemaCollection which adds IDispose to the interface
	/// </summary>
	public interface I_CustomXMLSchemaCollection : Microsoft.Office.Core._CustomXMLSchemaCollection, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._CustomXMLSchemaCollection Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomXMLSchemaCollection which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLSchemaCollection : Microsoft.Office.Core.CustomXMLSchemaCollection, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.CustomXMLSchemaCollection Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomXMLNodes which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLNodes : Microsoft.Office.Core.CustomXMLNodes, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.CustomXMLNodes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomXMLNode which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLNode : Microsoft.Office.Core.CustomXMLNode, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.CustomXMLNode Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomXMLValidationError which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLValidationError : Microsoft.Office.Core.CustomXMLValidationError, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.CustomXMLValidationError Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomXMLValidationErrors which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLValidationErrors : Microsoft.Office.Core.CustomXMLValidationErrors, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.CustomXMLValidationErrors Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CustomXMLPart which adds IDispose to the interface
	/// </summary>
	public interface I_CustomXMLPart : Microsoft.Office.Core._CustomXMLPart, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._CustomXMLPart Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICustomXMLPartEvents which adds IDispose to the interface
	/// </summary>
	public interface IICustomXMLPartEvents : Microsoft.Office.Core.ICustomXMLPartEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ICustomXMLPartEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CustomXMLPartEvents which adds IDispose to the interface
	/// </summary>
	public interface I_CustomXMLPartEvents : Microsoft.Office.Core._CustomXMLPartEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._CustomXMLPartEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CustomXMLPartEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface I_CustomXMLPartEvents_Event : Microsoft.Office.Core._CustomXMLPartEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._CustomXMLPartEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomXMLPart which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLPart : Microsoft.Office.Core.CustomXMLPart, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.CustomXMLPart Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CustomXMLParts which adds IDispose to the interface
	/// </summary>
	public interface I_CustomXMLParts : Microsoft.Office.Core._CustomXMLParts, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._CustomXMLParts Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICustomXMLPartsEvents which adds IDispose to the interface
	/// </summary>
	public interface IICustomXMLPartsEvents : Microsoft.Office.Core.ICustomXMLPartsEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ICustomXMLPartsEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CustomXMLPartsEvents which adds IDispose to the interface
	/// </summary>
	public interface I_CustomXMLPartsEvents : Microsoft.Office.Core._CustomXMLPartsEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._CustomXMLPartsEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CustomXMLPartsEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface I_CustomXMLPartsEvents_Event : Microsoft.Office.Core._CustomXMLPartsEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._CustomXMLPartsEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomXMLParts which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLParts : Microsoft.Office.Core.CustomXMLParts, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.CustomXMLParts Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for GradientStop which adds IDispose to the interface
	/// </summary>
	public interface IGradientStop : Microsoft.Office.Core.GradientStop, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.GradientStop Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for GradientStops which adds IDispose to the interface
	/// </summary>
	public interface IGradientStops : Microsoft.Office.Core.GradientStops, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.GradientStops Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SoftEdgeFormat which adds IDispose to the interface
	/// </summary>
	public interface ISoftEdgeFormat : Microsoft.Office.Core.SoftEdgeFormat, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SoftEdgeFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for GlowFormat which adds IDispose to the interface
	/// </summary>
	public interface IGlowFormat : Microsoft.Office.Core.GlowFormat, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.GlowFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ReflectionFormat which adds IDispose to the interface
	/// </summary>
	public interface IReflectionFormat : Microsoft.Office.Core.ReflectionFormat, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ReflectionFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ParagraphFormat2 which adds IDispose to the interface
	/// </summary>
	public interface IParagraphFormat2 : Microsoft.Office.Core.ParagraphFormat2, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ParagraphFormat2 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Font2 which adds IDispose to the interface
	/// </summary>
	public interface IFont2 : Microsoft.Office.Core.Font2, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.Font2 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextColumn2 which adds IDispose to the interface
	/// </summary>
	public interface ITextColumn2 : Microsoft.Office.Core.TextColumn2, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.TextColumn2 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextRange2 which adds IDispose to the interface
	/// </summary>
	public interface ITextRange2 : Microsoft.Office.Core.TextRange2, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.TextRange2 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextFrame2 which adds IDispose to the interface
	/// </summary>
	public interface ITextFrame2 : Microsoft.Office.Core.TextFrame2, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.TextFrame2 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ThemeColor which adds IDispose to the interface
	/// </summary>
	public interface IThemeColor : Microsoft.Office.Core.ThemeColor, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ThemeColor Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ThemeColorScheme which adds IDispose to the interface
	/// </summary>
	public interface IThemeColorScheme : Microsoft.Office.Core.ThemeColorScheme, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ThemeColorScheme Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ThemeFont which adds IDispose to the interface
	/// </summary>
	public interface IThemeFont : Microsoft.Office.Core.ThemeFont, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ThemeFont Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ThemeFonts which adds IDispose to the interface
	/// </summary>
	public interface IThemeFonts : Microsoft.Office.Core.ThemeFonts, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ThemeFonts Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ThemeFontScheme which adds IDispose to the interface
	/// </summary>
	public interface IThemeFontScheme : Microsoft.Office.Core.ThemeFontScheme, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ThemeFontScheme Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ThemeEffectScheme which adds IDispose to the interface
	/// </summary>
	public interface IThemeEffectScheme : Microsoft.Office.Core.ThemeEffectScheme, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ThemeEffectScheme Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OfficeTheme which adds IDispose to the interface
	/// </summary>
	public interface IOfficeTheme : Microsoft.Office.Core.OfficeTheme, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.OfficeTheme Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CustomTaskPane which adds IDispose to the interface
	/// </summary>
	public interface I_CustomTaskPane : Microsoft.Office.Core._CustomTaskPane, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._CustomTaskPane Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomTaskPaneEvents which adds IDispose to the interface
	/// </summary>
	public interface ICustomTaskPaneEvents : Microsoft.Office.Core.CustomTaskPaneEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.CustomTaskPaneEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CustomTaskPaneEvents which adds IDispose to the interface
	/// </summary>
	public interface I_CustomTaskPaneEvents : Microsoft.Office.Core._CustomTaskPaneEvents, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._CustomTaskPaneEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _CustomTaskPaneEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface I_CustomTaskPaneEvents_Event : Microsoft.Office.Core._CustomTaskPaneEvents_Event, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core._CustomTaskPaneEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomTaskPane which adds IDispose to the interface
	/// </summary>
	public interface ICustomTaskPane : Microsoft.Office.Core.CustomTaskPane, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.CustomTaskPane Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICTPFactory which adds IDispose to the interface
	/// </summary>
	public interface IICTPFactory : Microsoft.Office.Core.ICTPFactory, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ICTPFactory Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICustomTaskPaneConsumer which adds IDispose to the interface
	/// </summary>
	public interface IICustomTaskPaneConsumer : Microsoft.Office.Core.ICustomTaskPaneConsumer, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ICustomTaskPaneConsumer Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IRibbonUI which adds IDispose to the interface
	/// </summary>
	public interface IIRibbonUI : Microsoft.Office.Core.IRibbonUI, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IRibbonUI Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IRibbonControl which adds IDispose to the interface
	/// </summary>
	public interface IIRibbonControl : Microsoft.Office.Core.IRibbonControl, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IRibbonControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IRibbonExtensibility which adds IDispose to the interface
	/// </summary>
	public interface IIRibbonExtensibility : Microsoft.Office.Core.IRibbonExtensibility, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IRibbonExtensibility Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IAssistance which adds IDispose to the interface
	/// </summary>
	public interface IIAssistance : Microsoft.Office.Core.IAssistance, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IAssistance Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoChartData which adds IDispose to the interface
	/// </summary>
	public interface IIMsoChartData : Microsoft.Office.Core.IMsoChartData, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoChartData Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoChart which adds IDispose to the interface
	/// </summary>
	public interface IIMsoChart : Microsoft.Office.Core.IMsoChart, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoChart Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoCorners which adds IDispose to the interface
	/// </summary>
	public interface IIMsoCorners : Microsoft.Office.Core.IMsoCorners, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoCorners Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoLegend which adds IDispose to the interface
	/// </summary>
	public interface IIMsoLegend : Microsoft.Office.Core.IMsoLegend, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoLegend Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoBorder which adds IDispose to the interface
	/// </summary>
	public interface IIMsoBorder : Microsoft.Office.Core.IMsoBorder, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoBorder Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoWalls which adds IDispose to the interface
	/// </summary>
	public interface IIMsoWalls : Microsoft.Office.Core.IMsoWalls, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoWalls Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoFloor which adds IDispose to the interface
	/// </summary>
	public interface IIMsoFloor : Microsoft.Office.Core.IMsoFloor, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoFloor Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoPlotArea which adds IDispose to the interface
	/// </summary>
	public interface IIMsoPlotArea : Microsoft.Office.Core.IMsoPlotArea, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoPlotArea Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoChartArea which adds IDispose to the interface
	/// </summary>
	public interface IIMsoChartArea : Microsoft.Office.Core.IMsoChartArea, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoChartArea Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoSeriesLines which adds IDispose to the interface
	/// </summary>
	public interface IIMsoSeriesLines : Microsoft.Office.Core.IMsoSeriesLines, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoSeriesLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoLeaderLines which adds IDispose to the interface
	/// </summary>
	public interface IIMsoLeaderLines : Microsoft.Office.Core.IMsoLeaderLines, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoLeaderLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for GridLines which adds IDispose to the interface
	/// </summary>
	public interface IGridLines : Microsoft.Office.Core.GridLines, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.GridLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoUpBars which adds IDispose to the interface
	/// </summary>
	public interface IIMsoUpBars : Microsoft.Office.Core.IMsoUpBars, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoUpBars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoDownBars which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDownBars : Microsoft.Office.Core.IMsoDownBars, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoDownBars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoInterior which adds IDispose to the interface
	/// </summary>
	public interface IIMsoInterior : Microsoft.Office.Core.IMsoInterior, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoInterior Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartFillFormat which adds IDispose to the interface
	/// </summary>
	public interface IChartFillFormat : Microsoft.Office.Core.ChartFillFormat, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ChartFillFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LegendEntries which adds IDispose to the interface
	/// </summary>
	public interface ILegendEntries : Microsoft.Office.Core.LegendEntries, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.LegendEntries Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartFont which adds IDispose to the interface
	/// </summary>
	public interface IChartFont : Microsoft.Office.Core.ChartFont, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ChartFont Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartColorFormat which adds IDispose to the interface
	/// </summary>
	public interface IChartColorFormat : Microsoft.Office.Core.ChartColorFormat, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ChartColorFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LegendEntry which adds IDispose to the interface
	/// </summary>
	public interface ILegendEntry : Microsoft.Office.Core.LegendEntry, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.LegendEntry Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoLegendKey which adds IDispose to the interface
	/// </summary>
	public interface IIMsoLegendKey : Microsoft.Office.Core.IMsoLegendKey, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoLegendKey Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SeriesCollection which adds IDispose to the interface
	/// </summary>
	public interface ISeriesCollection : Microsoft.Office.Core.SeriesCollection, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SeriesCollection Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoSeries which adds IDispose to the interface
	/// </summary>
	public interface IIMsoSeries : Microsoft.Office.Core.IMsoSeries, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoSeries Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoErrorBars which adds IDispose to the interface
	/// </summary>
	public interface IIMsoErrorBars : Microsoft.Office.Core.IMsoErrorBars, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoErrorBars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoTrendline which adds IDispose to the interface
	/// </summary>
	public interface IIMsoTrendline : Microsoft.Office.Core.IMsoTrendline, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoTrendline Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Trendlines which adds IDispose to the interface
	/// </summary>
	public interface ITrendlines : Microsoft.Office.Core.Trendlines, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.Trendlines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoDataLabels which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDataLabels : Microsoft.Office.Core.IMsoDataLabels, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoDataLabels Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoDataLabel which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDataLabel : Microsoft.Office.Core.IMsoDataLabel, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoDataLabel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Points which adds IDispose to the interface
	/// </summary>
	public interface IPoints : Microsoft.Office.Core.Points, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.Points Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartPoint which adds IDispose to the interface
	/// </summary>
	public interface IChartPoint : Microsoft.Office.Core.ChartPoint, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ChartPoint Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Axes which adds IDispose to the interface
	/// </summary>
	public interface IAxes : Microsoft.Office.Core.Axes, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.Axes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoAxis which adds IDispose to the interface
	/// </summary>
	public interface IIMsoAxis : Microsoft.Office.Core.IMsoAxis, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoAxis Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoDataTable which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDataTable : Microsoft.Office.Core.IMsoDataTable, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoDataTable Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoChartTitle which adds IDispose to the interface
	/// </summary>
	public interface IIMsoChartTitle : Microsoft.Office.Core.IMsoChartTitle, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoChartTitle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoAxisTitle which adds IDispose to the interface
	/// </summary>
	public interface IIMsoAxisTitle : Microsoft.Office.Core.IMsoAxisTitle, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoAxisTitle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoDisplayUnitLabel which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDisplayUnitLabel : Microsoft.Office.Core.IMsoDisplayUnitLabel, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoDisplayUnitLabel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoTickLabels which adds IDispose to the interface
	/// </summary>
	public interface IIMsoTickLabels : Microsoft.Office.Core.IMsoTickLabels, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoTickLabels Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoHyperlinks which adds IDispose to the interface
	/// </summary>
	public interface IIMsoHyperlinks : Microsoft.Office.Core.IMsoHyperlinks, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoHyperlinks Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoDropLines which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDropLines : Microsoft.Office.Core.IMsoDropLines, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoDropLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoHiLoLines which adds IDispose to the interface
	/// </summary>
	public interface IIMsoHiLoLines : Microsoft.Office.Core.IMsoHiLoLines, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoHiLoLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoChartGroup which adds IDispose to the interface
	/// </summary>
	public interface IIMsoChartGroup : Microsoft.Office.Core.IMsoChartGroup, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoChartGroup Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartGroups which adds IDispose to the interface
	/// </summary>
	public interface IChartGroups : Microsoft.Office.Core.ChartGroups, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ChartGroups Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoCharacters which adds IDispose to the interface
	/// </summary>
	public interface IIMsoCharacters : Microsoft.Office.Core.IMsoCharacters, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoCharacters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoChartFormat which adds IDispose to the interface
	/// </summary>
	public interface IIMsoChartFormat : Microsoft.Office.Core.IMsoChartFormat, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoChartFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for BulletFormat2 which adds IDispose to the interface
	/// </summary>
	public interface IBulletFormat2 : Microsoft.Office.Core.BulletFormat2, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.BulletFormat2 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TabStops2 which adds IDispose to the interface
	/// </summary>
	public interface ITabStops2 : Microsoft.Office.Core.TabStops2, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.TabStops2 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TabStop2 which adds IDispose to the interface
	/// </summary>
	public interface ITabStop2 : Microsoft.Office.Core.TabStop2, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.TabStop2 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Ruler2 which adds IDispose to the interface
	/// </summary>
	public interface IRuler2 : Microsoft.Office.Core.Ruler2, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.Ruler2 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RulerLevels2 which adds IDispose to the interface
	/// </summary>
	public interface IRulerLevels2 : Microsoft.Office.Core.RulerLevels2, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.RulerLevels2 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RulerLevel2 which adds IDispose to the interface
	/// </summary>
	public interface IRulerLevel2 : Microsoft.Office.Core.RulerLevel2, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.RulerLevel2 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for EncryptionProvider which adds IDispose to the interface
	/// </summary>
	public interface IEncryptionProvider : Microsoft.Office.Core.EncryptionProvider, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.EncryptionProvider Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IBlogExtensibility which adds IDispose to the interface
	/// </summary>
	public interface IIBlogExtensibility : Microsoft.Office.Core.IBlogExtensibility, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IBlogExtensibility Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IBlogPictureExtensibility which adds IDispose to the interface
	/// </summary>
	public interface IIBlogPictureExtensibility : Microsoft.Office.Core.IBlogPictureExtensibility, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IBlogPictureExtensibility Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IConverterPreferences which adds IDispose to the interface
	/// </summary>
	public interface IIConverterPreferences : Microsoft.Office.Core.IConverterPreferences, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IConverterPreferences Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IConverterApplicationPreferences which adds IDispose to the interface
	/// </summary>
	public interface IIConverterApplicationPreferences : Microsoft.Office.Core.IConverterApplicationPreferences, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IConverterApplicationPreferences Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IConverterUICallback which adds IDispose to the interface
	/// </summary>
	public interface IIConverterUICallback : Microsoft.Office.Core.IConverterUICallback, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IConverterUICallback Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IConverter which adds IDispose to the interface
	/// </summary>
	public interface IIConverter : Microsoft.Office.Core.IConverter, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IConverter Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartArt which adds IDispose to the interface
	/// </summary>
	public interface ISmartArt : Microsoft.Office.Core.SmartArt, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SmartArt Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartArtNodes which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtNodes : Microsoft.Office.Core.SmartArtNodes, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SmartArtNodes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartArtNode which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtNode : Microsoft.Office.Core.SmartArtNode, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SmartArtNode Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartArtLayouts which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtLayouts : Microsoft.Office.Core.SmartArtLayouts, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SmartArtLayouts Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartArtLayout which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtLayout : Microsoft.Office.Core.SmartArtLayout, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SmartArtLayout Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartArtQuickStyles which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtQuickStyles : Microsoft.Office.Core.SmartArtQuickStyles, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SmartArtQuickStyles Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartArtQuickStyle which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtQuickStyle : Microsoft.Office.Core.SmartArtQuickStyle, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SmartArtQuickStyle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartArtColors which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtColors : Microsoft.Office.Core.SmartArtColors, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SmartArtColors Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartArtColor which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtColor : Microsoft.Office.Core.SmartArtColor, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.SmartArtColor Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PickerField which adds IDispose to the interface
	/// </summary>
	public interface IPickerField : Microsoft.Office.Core.PickerField, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.PickerField Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PickerFields which adds IDispose to the interface
	/// </summary>
	public interface IPickerFields : Microsoft.Office.Core.PickerFields, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.PickerFields Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PickerProperty which adds IDispose to the interface
	/// </summary>
	public interface IPickerProperty : Microsoft.Office.Core.PickerProperty, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.PickerProperty Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PickerProperties which adds IDispose to the interface
	/// </summary>
	public interface IPickerProperties : Microsoft.Office.Core.PickerProperties, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.PickerProperties Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PickerResult which adds IDispose to the interface
	/// </summary>
	public interface IPickerResult : Microsoft.Office.Core.PickerResult, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.PickerResult Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PickerResults which adds IDispose to the interface
	/// </summary>
	public interface IPickerResults : Microsoft.Office.Core.PickerResults, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.PickerResults Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PickerDialog which adds IDispose to the interface
	/// </summary>
	public interface IPickerDialog : Microsoft.Office.Core.PickerDialog, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.PickerDialog Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMsoContactCard which adds IDispose to the interface
	/// </summary>
	public interface IIMsoContactCard : Microsoft.Office.Core.IMsoContactCard, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.IMsoContactCard Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for EffectParameter which adds IDispose to the interface
	/// </summary>
	public interface IEffectParameter : Microsoft.Office.Core.EffectParameter, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.EffectParameter Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for EffectParameters which adds IDispose to the interface
	/// </summary>
	public interface IEffectParameters : Microsoft.Office.Core.EffectParameters, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.EffectParameters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PictureEffect which adds IDispose to the interface
	/// </summary>
	public interface IPictureEffect : Microsoft.Office.Core.PictureEffect, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.PictureEffect Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PictureEffects which adds IDispose to the interface
	/// </summary>
	public interface IPictureEffects : Microsoft.Office.Core.PictureEffects, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.PictureEffects Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Crop which adds IDispose to the interface
	/// </summary>
	public interface ICrop : Microsoft.Office.Core.Crop, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.Crop Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ContactCard which adds IDispose to the interface
	/// </summary>
	public interface IContactCard : Microsoft.Office.Core.ContactCard, System.IDisposable 
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Core.ContactCard Resource { get; }
	}

}