//office, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c
namespace VSTOContrib.Extensions.Proxies
{
	/// <summary>
	/// Provides cleanup extension methods for interfaces exposed by Office.dll
	/// </summary>
	public static class OfficeCleanupExtensions
	{
		/// <summary>
		/// Wrapper interface for IAccessible which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIAccessible WithComCleanupProxy(this Microsoft.Office.Core.IAccessible resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IAccessible, Interfaces.IIAccessible>();
		}

		/// <summary>
		/// Wrapper interface for _IMsoDispObj which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_IMsoDispObj WithComCleanupProxy(this Microsoft.Office.Core._IMsoDispObj resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._IMsoDispObj, Interfaces.I_IMsoDispObj>();
		}

		/// <summary>
		/// Wrapper interface for _IMsoOleAccDispObj which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_IMsoOleAccDispObj WithComCleanupProxy(this Microsoft.Office.Core._IMsoOleAccDispObj resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._IMsoOleAccDispObj, Interfaces.I_IMsoOleAccDispObj>();
		}

		/// <summary>
		/// Wrapper interface for _CommandBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CommandBars WithComCleanupProxy(this Microsoft.Office.Core._CommandBars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._CommandBars, Interfaces.I_CommandBars>();
		}

		/// <summary>
		/// Wrapper interface for CommandBar which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICommandBar WithComCleanupProxy(this Microsoft.Office.Core.CommandBar resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.CommandBar, Interfaces.ICommandBar>();
		}

		/// <summary>
		/// Wrapper interface for CommandBarControls which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICommandBarControls WithComCleanupProxy(this Microsoft.Office.Core.CommandBarControls resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.CommandBarControls, Interfaces.ICommandBarControls>();
		}

		/// <summary>
		/// Wrapper interface for CommandBarControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICommandBarControl WithComCleanupProxy(this Microsoft.Office.Core.CommandBarControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.CommandBarControl, Interfaces.ICommandBarControl>();
		}

		/// <summary>
		/// Wrapper interface for _CommandBarButton which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CommandBarButton WithComCleanupProxy(this Microsoft.Office.Core._CommandBarButton resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._CommandBarButton, Interfaces.I_CommandBarButton>();
		}

		/// <summary>
		/// Wrapper interface for CommandBarPopup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICommandBarPopup WithComCleanupProxy(this Microsoft.Office.Core.CommandBarPopup resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.CommandBarPopup, Interfaces.ICommandBarPopup>();
		}

		/// <summary>
		/// Wrapper interface for _CommandBarComboBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CommandBarComboBox WithComCleanupProxy(this Microsoft.Office.Core._CommandBarComboBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._CommandBarComboBox, Interfaces.I_CommandBarComboBox>();
		}

		/// <summary>
		/// Wrapper interface for _CommandBarActiveX which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CommandBarActiveX WithComCleanupProxy(this Microsoft.Office.Core._CommandBarActiveX resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._CommandBarActiveX, Interfaces.I_CommandBarActiveX>();
		}

		/// <summary>
		/// Wrapper interface for Adjustments which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAdjustments WithComCleanupProxy(this Microsoft.Office.Core.Adjustments resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.Adjustments, Interfaces.IAdjustments>();
		}

		/// <summary>
		/// Wrapper interface for CalloutFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICalloutFormat WithComCleanupProxy(this Microsoft.Office.Core.CalloutFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.CalloutFormat, Interfaces.ICalloutFormat>();
		}

		/// <summary>
		/// Wrapper interface for ColorFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColorFormat WithComCleanupProxy(this Microsoft.Office.Core.ColorFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ColorFormat, Interfaces.IColorFormat>();
		}

		/// <summary>
		/// Wrapper interface for ConnectorFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IConnectorFormat WithComCleanupProxy(this Microsoft.Office.Core.ConnectorFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ConnectorFormat, Interfaces.IConnectorFormat>();
		}

		/// <summary>
		/// Wrapper interface for FillFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFillFormat WithComCleanupProxy(this Microsoft.Office.Core.FillFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.FillFormat, Interfaces.IFillFormat>();
		}

		/// <summary>
		/// Wrapper interface for FreeformBuilder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFreeformBuilder WithComCleanupProxy(this Microsoft.Office.Core.FreeformBuilder resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.FreeformBuilder, Interfaces.IFreeformBuilder>();
		}

		/// <summary>
		/// Wrapper interface for GroupShapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGroupShapes WithComCleanupProxy(this Microsoft.Office.Core.GroupShapes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.GroupShapes, Interfaces.IGroupShapes>();
		}

		/// <summary>
		/// Wrapper interface for LineFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILineFormat WithComCleanupProxy(this Microsoft.Office.Core.LineFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.LineFormat, Interfaces.ILineFormat>();
		}

		/// <summary>
		/// Wrapper interface for ShapeNode which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapeNode WithComCleanupProxy(this Microsoft.Office.Core.ShapeNode resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ShapeNode, Interfaces.IShapeNode>();
		}

		/// <summary>
		/// Wrapper interface for ShapeNodes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapeNodes WithComCleanupProxy(this Microsoft.Office.Core.ShapeNodes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ShapeNodes, Interfaces.IShapeNodes>();
		}

		/// <summary>
		/// Wrapper interface for PictureFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPictureFormat WithComCleanupProxy(this Microsoft.Office.Core.PictureFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.PictureFormat, Interfaces.IPictureFormat>();
		}

		/// <summary>
		/// Wrapper interface for ShadowFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShadowFormat WithComCleanupProxy(this Microsoft.Office.Core.ShadowFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ShadowFormat, Interfaces.IShadowFormat>();
		}

		/// <summary>
		/// Wrapper interface for Script which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IScript WithComCleanupProxy(this Microsoft.Office.Core.Script resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.Script, Interfaces.IScript>();
		}

		/// <summary>
		/// Wrapper interface for Scripts which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IScripts WithComCleanupProxy(this Microsoft.Office.Core.Scripts resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.Scripts, Interfaces.IScripts>();
		}

		/// <summary>
		/// Wrapper interface for Shape which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShape WithComCleanupProxy(this Microsoft.Office.Core.Shape resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.Shape, Interfaces.IShape>();
		}

		/// <summary>
		/// Wrapper interface for ShapeRange which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapeRange WithComCleanupProxy(this Microsoft.Office.Core.ShapeRange resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ShapeRange, Interfaces.IShapeRange>();
		}

		/// <summary>
		/// Wrapper interface for Shapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapes WithComCleanupProxy(this Microsoft.Office.Core.Shapes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.Shapes, Interfaces.IShapes>();
		}

		/// <summary>
		/// Wrapper interface for TextEffectFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextEffectFormat WithComCleanupProxy(this Microsoft.Office.Core.TextEffectFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.TextEffectFormat, Interfaces.ITextEffectFormat>();
		}

		/// <summary>
		/// Wrapper interface for TextFrame which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextFrame WithComCleanupProxy(this Microsoft.Office.Core.TextFrame resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.TextFrame, Interfaces.ITextFrame>();
		}

		/// <summary>
		/// Wrapper interface for ThreeDFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IThreeDFormat WithComCleanupProxy(this Microsoft.Office.Core.ThreeDFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ThreeDFormat, Interfaces.IThreeDFormat>();
		}

		/// <summary>
		/// Wrapper interface for IMsoDispCagNotifySink which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoDispCagNotifySink WithComCleanupProxy(this Microsoft.Office.Core.IMsoDispCagNotifySink resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoDispCagNotifySink, Interfaces.IIMsoDispCagNotifySink>();
		}

		/// <summary>
		/// Wrapper interface for Balloon which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBalloon WithComCleanupProxy(this Microsoft.Office.Core.Balloon resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.Balloon, Interfaces.IBalloon>();
		}

		/// <summary>
		/// Wrapper interface for BalloonCheckboxes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBalloonCheckboxes WithComCleanupProxy(this Microsoft.Office.Core.BalloonCheckboxes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.BalloonCheckboxes, Interfaces.IBalloonCheckboxes>();
		}

		/// <summary>
		/// Wrapper interface for BalloonCheckbox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBalloonCheckbox WithComCleanupProxy(this Microsoft.Office.Core.BalloonCheckbox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.BalloonCheckbox, Interfaces.IBalloonCheckbox>();
		}

		/// <summary>
		/// Wrapper interface for BalloonLabels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBalloonLabels WithComCleanupProxy(this Microsoft.Office.Core.BalloonLabels resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.BalloonLabels, Interfaces.IBalloonLabels>();
		}

		/// <summary>
		/// Wrapper interface for BalloonLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBalloonLabel WithComCleanupProxy(this Microsoft.Office.Core.BalloonLabel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.BalloonLabel, Interfaces.IBalloonLabel>();
		}

		/// <summary>
		/// Wrapper interface for AnswerWizardFiles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAnswerWizardFiles WithComCleanupProxy(this Microsoft.Office.Core.AnswerWizardFiles resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.AnswerWizardFiles, Interfaces.IAnswerWizardFiles>();
		}

		/// <summary>
		/// Wrapper interface for AnswerWizard which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAnswerWizard WithComCleanupProxy(this Microsoft.Office.Core.AnswerWizard resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.AnswerWizard, Interfaces.IAnswerWizard>();
		}

		/// <summary>
		/// Wrapper interface for Assistant which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAssistant WithComCleanupProxy(this Microsoft.Office.Core.Assistant resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.Assistant, Interfaces.IAssistant>();
		}

		/// <summary>
		/// Wrapper interface for DocumentProperty which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocumentProperty WithComCleanupProxy(this Microsoft.Office.Core.DocumentProperty resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.DocumentProperty, Interfaces.IDocumentProperty>();
		}

		/// <summary>
		/// Wrapper interface for DocumentProperties which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocumentProperties WithComCleanupProxy(this Microsoft.Office.Core.DocumentProperties resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.DocumentProperties, Interfaces.IDocumentProperties>();
		}

		/// <summary>
		/// Wrapper interface for IFoundFiles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIFoundFiles WithComCleanupProxy(this Microsoft.Office.Core.IFoundFiles resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IFoundFiles, Interfaces.IIFoundFiles>();
		}

		/// <summary>
		/// Wrapper interface for IFind which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIFind WithComCleanupProxy(this Microsoft.Office.Core.IFind resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IFind, Interfaces.IIFind>();
		}

		/// <summary>
		/// Wrapper interface for FoundFiles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFoundFiles WithComCleanupProxy(this Microsoft.Office.Core.FoundFiles resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.FoundFiles, Interfaces.IFoundFiles>();
		}

		/// <summary>
		/// Wrapper interface for PropertyTest which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPropertyTest WithComCleanupProxy(this Microsoft.Office.Core.PropertyTest resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.PropertyTest, Interfaces.IPropertyTest>();
		}

		/// <summary>
		/// Wrapper interface for PropertyTests which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPropertyTests WithComCleanupProxy(this Microsoft.Office.Core.PropertyTests resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.PropertyTests, Interfaces.IPropertyTests>();
		}

		/// <summary>
		/// Wrapper interface for FileSearch which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFileSearch WithComCleanupProxy(this Microsoft.Office.Core.FileSearch resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.FileSearch, Interfaces.IFileSearch>();
		}

		/// <summary>
		/// Wrapper interface for COMAddIn which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICOMAddIn WithComCleanupProxy(this Microsoft.Office.Core.COMAddIn resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.COMAddIn, Interfaces.ICOMAddIn>();
		}

		/// <summary>
		/// Wrapper interface for COMAddIns which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICOMAddIns WithComCleanupProxy(this Microsoft.Office.Core.COMAddIns resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.COMAddIns, Interfaces.ICOMAddIns>();
		}

		/// <summary>
		/// Wrapper interface for LanguageSettings which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILanguageSettings WithComCleanupProxy(this Microsoft.Office.Core.LanguageSettings resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.LanguageSettings, Interfaces.ILanguageSettings>();
		}

		/// <summary>
		/// Wrapper interface for ICommandBarsEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICommandBarsEvents WithComCleanupProxy(this Microsoft.Office.Core.ICommandBarsEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ICommandBarsEvents, Interfaces.IICommandBarsEvents>();
		}

		/// <summary>
		/// Wrapper interface for _CommandBarsEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CommandBarsEvents WithComCleanupProxy(this Microsoft.Office.Core._CommandBarsEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._CommandBarsEvents, Interfaces.I_CommandBarsEvents>();
		}

		/// <summary>
		/// Wrapper interface for _CommandBarsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CommandBarsEvents_Event WithComCleanupProxy(this Microsoft.Office.Core._CommandBarsEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._CommandBarsEvents_Event, Interfaces.I_CommandBarsEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for CommandBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICommandBars WithComCleanupProxy(this Microsoft.Office.Core.CommandBars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.CommandBars, Interfaces.ICommandBars>();
		}

		/// <summary>
		/// Wrapper interface for ICommandBarComboBoxEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICommandBarComboBoxEvents WithComCleanupProxy(this Microsoft.Office.Core.ICommandBarComboBoxEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ICommandBarComboBoxEvents, Interfaces.IICommandBarComboBoxEvents>();
		}

		/// <summary>
		/// Wrapper interface for _CommandBarComboBoxEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CommandBarComboBoxEvents WithComCleanupProxy(this Microsoft.Office.Core._CommandBarComboBoxEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._CommandBarComboBoxEvents, Interfaces.I_CommandBarComboBoxEvents>();
		}

		/// <summary>
		/// Wrapper interface for _CommandBarComboBoxEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CommandBarComboBoxEvents_Event WithComCleanupProxy(this Microsoft.Office.Core._CommandBarComboBoxEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._CommandBarComboBoxEvents_Event, Interfaces.I_CommandBarComboBoxEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for CommandBarComboBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICommandBarComboBox WithComCleanupProxy(this Microsoft.Office.Core.CommandBarComboBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.CommandBarComboBox, Interfaces.ICommandBarComboBox>();
		}

		/// <summary>
		/// Wrapper interface for ICommandBarButtonEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICommandBarButtonEvents WithComCleanupProxy(this Microsoft.Office.Core.ICommandBarButtonEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ICommandBarButtonEvents, Interfaces.IICommandBarButtonEvents>();
		}

		/// <summary>
		/// Wrapper interface for _CommandBarButtonEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CommandBarButtonEvents WithComCleanupProxy(this Microsoft.Office.Core._CommandBarButtonEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._CommandBarButtonEvents, Interfaces.I_CommandBarButtonEvents>();
		}

		/// <summary>
		/// Wrapper interface for _CommandBarButtonEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CommandBarButtonEvents_Event WithComCleanupProxy(this Microsoft.Office.Core._CommandBarButtonEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._CommandBarButtonEvents_Event, Interfaces.I_CommandBarButtonEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for CommandBarButton which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICommandBarButton WithComCleanupProxy(this Microsoft.Office.Core.CommandBarButton resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.CommandBarButton, Interfaces.ICommandBarButton>();
		}

		/// <summary>
		/// Wrapper interface for WebPageFont which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWebPageFont WithComCleanupProxy(this Microsoft.Office.Core.WebPageFont resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.WebPageFont, Interfaces.IWebPageFont>();
		}

		/// <summary>
		/// Wrapper interface for WebPageFonts which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWebPageFonts WithComCleanupProxy(this Microsoft.Office.Core.WebPageFonts resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.WebPageFonts, Interfaces.IWebPageFonts>();
		}

		/// <summary>
		/// Wrapper interface for HTMLProjectItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHTMLProjectItem WithComCleanupProxy(this Microsoft.Office.Core.HTMLProjectItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.HTMLProjectItem, Interfaces.IHTMLProjectItem>();
		}

		/// <summary>
		/// Wrapper interface for HTMLProjectItems which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHTMLProjectItems WithComCleanupProxy(this Microsoft.Office.Core.HTMLProjectItems resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.HTMLProjectItems, Interfaces.IHTMLProjectItems>();
		}

		/// <summary>
		/// Wrapper interface for HTMLProject which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHTMLProject WithComCleanupProxy(this Microsoft.Office.Core.HTMLProject resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.HTMLProject, Interfaces.IHTMLProject>();
		}

		/// <summary>
		/// Wrapper interface for MsoDebugOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMsoDebugOptions WithComCleanupProxy(this Microsoft.Office.Core.MsoDebugOptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.MsoDebugOptions, Interfaces.IMsoDebugOptions>();
		}

		/// <summary>
		/// Wrapper interface for FileDialogSelectedItems which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFileDialogSelectedItems WithComCleanupProxy(this Microsoft.Office.Core.FileDialogSelectedItems resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.FileDialogSelectedItems, Interfaces.IFileDialogSelectedItems>();
		}

		/// <summary>
		/// Wrapper interface for FileDialogFilter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFileDialogFilter WithComCleanupProxy(this Microsoft.Office.Core.FileDialogFilter resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.FileDialogFilter, Interfaces.IFileDialogFilter>();
		}

		/// <summary>
		/// Wrapper interface for FileDialogFilters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFileDialogFilters WithComCleanupProxy(this Microsoft.Office.Core.FileDialogFilters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.FileDialogFilters, Interfaces.IFileDialogFilters>();
		}

		/// <summary>
		/// Wrapper interface for FileDialog which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFileDialog WithComCleanupProxy(this Microsoft.Office.Core.FileDialog resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.FileDialog, Interfaces.IFileDialog>();
		}

		/// <summary>
		/// Wrapper interface for SignatureSet which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISignatureSet WithComCleanupProxy(this Microsoft.Office.Core.SignatureSet resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SignatureSet, Interfaces.ISignatureSet>();
		}

		/// <summary>
		/// Wrapper interface for Signature which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISignature WithComCleanupProxy(this Microsoft.Office.Core.Signature resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.Signature, Interfaces.ISignature>();
		}

		/// <summary>
		/// Wrapper interface for IMsoEnvelopeVB which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoEnvelopeVB WithComCleanupProxy(this Microsoft.Office.Core.IMsoEnvelopeVB resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoEnvelopeVB, Interfaces.IIMsoEnvelopeVB>();
		}

		/// <summary>
		/// Wrapper interface for IMsoEnvelopeVBEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoEnvelopeVBEvents WithComCleanupProxy(this Microsoft.Office.Core.IMsoEnvelopeVBEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoEnvelopeVBEvents, Interfaces.IIMsoEnvelopeVBEvents>();
		}

		/// <summary>
		/// Wrapper interface for IMsoEnvelopeVBEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoEnvelopeVBEvents_Event WithComCleanupProxy(this Microsoft.Office.Core.IMsoEnvelopeVBEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoEnvelopeVBEvents_Event, Interfaces.IIMsoEnvelopeVBEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for MsoEnvelope which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMsoEnvelope WithComCleanupProxy(this Microsoft.Office.Core.MsoEnvelope resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.MsoEnvelope, Interfaces.IMsoEnvelope>();
		}

		/// <summary>
		/// Wrapper interface for FileTypes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFileTypes WithComCleanupProxy(this Microsoft.Office.Core.FileTypes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.FileTypes, Interfaces.IFileTypes>();
		}

		/// <summary>
		/// Wrapper interface for SearchFolders which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISearchFolders WithComCleanupProxy(this Microsoft.Office.Core.SearchFolders resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SearchFolders, Interfaces.ISearchFolders>();
		}

		/// <summary>
		/// Wrapper interface for ScopeFolders which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IScopeFolders WithComCleanupProxy(this Microsoft.Office.Core.ScopeFolders resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ScopeFolders, Interfaces.IScopeFolders>();
		}

		/// <summary>
		/// Wrapper interface for ScopeFolder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IScopeFolder WithComCleanupProxy(this Microsoft.Office.Core.ScopeFolder resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ScopeFolder, Interfaces.IScopeFolder>();
		}

		/// <summary>
		/// Wrapper interface for SearchScope which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISearchScope WithComCleanupProxy(this Microsoft.Office.Core.SearchScope resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SearchScope, Interfaces.ISearchScope>();
		}

		/// <summary>
		/// Wrapper interface for SearchScopes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISearchScopes WithComCleanupProxy(this Microsoft.Office.Core.SearchScopes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SearchScopes, Interfaces.ISearchScopes>();
		}

		/// <summary>
		/// Wrapper interface for IMsoDiagram which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoDiagram WithComCleanupProxy(this Microsoft.Office.Core.IMsoDiagram resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoDiagram, Interfaces.IIMsoDiagram>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNodes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagramNodes WithComCleanupProxy(this Microsoft.Office.Core.DiagramNodes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.DiagramNodes, Interfaces.IDiagramNodes>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNodeChildren which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagramNodeChildren WithComCleanupProxy(this Microsoft.Office.Core.DiagramNodeChildren resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.DiagramNodeChildren, Interfaces.IDiagramNodeChildren>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNode which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagramNode WithComCleanupProxy(this Microsoft.Office.Core.DiagramNode resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.DiagramNode, Interfaces.IDiagramNode>();
		}

		/// <summary>
		/// Wrapper interface for CanvasShapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICanvasShapes WithComCleanupProxy(this Microsoft.Office.Core.CanvasShapes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.CanvasShapes, Interfaces.ICanvasShapes>();
		}

		/// <summary>
		/// Wrapper interface for OfficeDataSourceObject which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOfficeDataSourceObject WithComCleanupProxy(this Microsoft.Office.Core.OfficeDataSourceObject resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.OfficeDataSourceObject, Interfaces.IOfficeDataSourceObject>();
		}

		/// <summary>
		/// Wrapper interface for ODSOColumn which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IODSOColumn WithComCleanupProxy(this Microsoft.Office.Core.ODSOColumn resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ODSOColumn, Interfaces.IODSOColumn>();
		}

		/// <summary>
		/// Wrapper interface for ODSOColumns which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IODSOColumns WithComCleanupProxy(this Microsoft.Office.Core.ODSOColumns resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ODSOColumns, Interfaces.IODSOColumns>();
		}

		/// <summary>
		/// Wrapper interface for ODSOFilter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IODSOFilter WithComCleanupProxy(this Microsoft.Office.Core.ODSOFilter resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ODSOFilter, Interfaces.IODSOFilter>();
		}

		/// <summary>
		/// Wrapper interface for ODSOFilters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IODSOFilters WithComCleanupProxy(this Microsoft.Office.Core.ODSOFilters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ODSOFilters, Interfaces.IODSOFilters>();
		}

		/// <summary>
		/// Wrapper interface for NewFile which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INewFile WithComCleanupProxy(this Microsoft.Office.Core.NewFile resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.NewFile, Interfaces.INewFile>();
		}

		/// <summary>
		/// Wrapper interface for WebComponent which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWebComponent WithComCleanupProxy(this Microsoft.Office.Core.WebComponent resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.WebComponent, Interfaces.IWebComponent>();
		}

		/// <summary>
		/// Wrapper interface for WebComponentWindowExternal which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWebComponentWindowExternal WithComCleanupProxy(this Microsoft.Office.Core.WebComponentWindowExternal resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.WebComponentWindowExternal, Interfaces.IWebComponentWindowExternal>();
		}

		/// <summary>
		/// Wrapper interface for WebComponentFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWebComponentFormat WithComCleanupProxy(this Microsoft.Office.Core.WebComponentFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.WebComponentFormat, Interfaces.IWebComponentFormat>();
		}

		/// <summary>
		/// Wrapper interface for ILicWizExternal which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IILicWizExternal WithComCleanupProxy(this Microsoft.Office.Core.ILicWizExternal resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ILicWizExternal, Interfaces.IILicWizExternal>();
		}

		/// <summary>
		/// Wrapper interface for ILicValidator which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IILicValidator WithComCleanupProxy(this Microsoft.Office.Core.ILicValidator resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ILicValidator, Interfaces.IILicValidator>();
		}

		/// <summary>
		/// Wrapper interface for ILicAgent which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IILicAgent WithComCleanupProxy(this Microsoft.Office.Core.ILicAgent resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ILicAgent, Interfaces.IILicAgent>();
		}

		/// <summary>
		/// Wrapper interface for IMsoEServicesDialog which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoEServicesDialog WithComCleanupProxy(this Microsoft.Office.Core.IMsoEServicesDialog resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoEServicesDialog, Interfaces.IIMsoEServicesDialog>();
		}

		/// <summary>
		/// Wrapper interface for WebComponentProperties which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWebComponentProperties WithComCleanupProxy(this Microsoft.Office.Core.WebComponentProperties resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.WebComponentProperties, Interfaces.IWebComponentProperties>();
		}

		/// <summary>
		/// Wrapper interface for SmartDocument which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartDocument WithComCleanupProxy(this Microsoft.Office.Core.SmartDocument resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SmartDocument, Interfaces.ISmartDocument>();
		}

		/// <summary>
		/// Wrapper interface for SharedWorkspaceMember which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISharedWorkspaceMember WithComCleanupProxy(this Microsoft.Office.Core.SharedWorkspaceMember resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SharedWorkspaceMember, Interfaces.ISharedWorkspaceMember>();
		}

		/// <summary>
		/// Wrapper interface for SharedWorkspaceMembers which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISharedWorkspaceMembers WithComCleanupProxy(this Microsoft.Office.Core.SharedWorkspaceMembers resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SharedWorkspaceMembers, Interfaces.ISharedWorkspaceMembers>();
		}

		/// <summary>
		/// Wrapper interface for SharedWorkspaceTask which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISharedWorkspaceTask WithComCleanupProxy(this Microsoft.Office.Core.SharedWorkspaceTask resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SharedWorkspaceTask, Interfaces.ISharedWorkspaceTask>();
		}

		/// <summary>
		/// Wrapper interface for SharedWorkspaceTasks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISharedWorkspaceTasks WithComCleanupProxy(this Microsoft.Office.Core.SharedWorkspaceTasks resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SharedWorkspaceTasks, Interfaces.ISharedWorkspaceTasks>();
		}

		/// <summary>
		/// Wrapper interface for SharedWorkspaceFile which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISharedWorkspaceFile WithComCleanupProxy(this Microsoft.Office.Core.SharedWorkspaceFile resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SharedWorkspaceFile, Interfaces.ISharedWorkspaceFile>();
		}

		/// <summary>
		/// Wrapper interface for SharedWorkspaceFiles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISharedWorkspaceFiles WithComCleanupProxy(this Microsoft.Office.Core.SharedWorkspaceFiles resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SharedWorkspaceFiles, Interfaces.ISharedWorkspaceFiles>();
		}

		/// <summary>
		/// Wrapper interface for SharedWorkspaceFolder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISharedWorkspaceFolder WithComCleanupProxy(this Microsoft.Office.Core.SharedWorkspaceFolder resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SharedWorkspaceFolder, Interfaces.ISharedWorkspaceFolder>();
		}

		/// <summary>
		/// Wrapper interface for SharedWorkspaceFolders which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISharedWorkspaceFolders WithComCleanupProxy(this Microsoft.Office.Core.SharedWorkspaceFolders resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SharedWorkspaceFolders, Interfaces.ISharedWorkspaceFolders>();
		}

		/// <summary>
		/// Wrapper interface for SharedWorkspaceLink which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISharedWorkspaceLink WithComCleanupProxy(this Microsoft.Office.Core.SharedWorkspaceLink resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SharedWorkspaceLink, Interfaces.ISharedWorkspaceLink>();
		}

		/// <summary>
		/// Wrapper interface for SharedWorkspaceLinks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISharedWorkspaceLinks WithComCleanupProxy(this Microsoft.Office.Core.SharedWorkspaceLinks resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SharedWorkspaceLinks, Interfaces.ISharedWorkspaceLinks>();
		}

		/// <summary>
		/// Wrapper interface for SharedWorkspace which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISharedWorkspace WithComCleanupProxy(this Microsoft.Office.Core.SharedWorkspace resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SharedWorkspace, Interfaces.ISharedWorkspace>();
		}

		/// <summary>
		/// Wrapper interface for Sync which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISync WithComCleanupProxy(this Microsoft.Office.Core.Sync resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.Sync, Interfaces.ISync>();
		}

		/// <summary>
		/// Wrapper interface for DocumentLibraryVersion which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocumentLibraryVersion WithComCleanupProxy(this Microsoft.Office.Core.DocumentLibraryVersion resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.DocumentLibraryVersion, Interfaces.IDocumentLibraryVersion>();
		}

		/// <summary>
		/// Wrapper interface for DocumentLibraryVersions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocumentLibraryVersions WithComCleanupProxy(this Microsoft.Office.Core.DocumentLibraryVersions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.DocumentLibraryVersions, Interfaces.IDocumentLibraryVersions>();
		}

		/// <summary>
		/// Wrapper interface for UserPermission which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IUserPermission WithComCleanupProxy(this Microsoft.Office.Core.UserPermission resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.UserPermission, Interfaces.IUserPermission>();
		}

		/// <summary>
		/// Wrapper interface for Permission which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPermission WithComCleanupProxy(this Microsoft.Office.Core.Permission resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.Permission, Interfaces.IPermission>();
		}

		/// <summary>
		/// Wrapper interface for MsoDebugOptions_UTRunResult which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMsoDebugOptions_UTRunResult WithComCleanupProxy(this Microsoft.Office.Core.MsoDebugOptions_UTRunResult resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.MsoDebugOptions_UTRunResult, Interfaces.IMsoDebugOptions_UTRunResult>();
		}

		/// <summary>
		/// Wrapper interface for MsoDebugOptions_UT which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMsoDebugOptions_UT WithComCleanupProxy(this Microsoft.Office.Core.MsoDebugOptions_UT resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.MsoDebugOptions_UT, Interfaces.IMsoDebugOptions_UT>();
		}

		/// <summary>
		/// Wrapper interface for MsoDebugOptions_UTs which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMsoDebugOptions_UTs WithComCleanupProxy(this Microsoft.Office.Core.MsoDebugOptions_UTs resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.MsoDebugOptions_UTs, Interfaces.IMsoDebugOptions_UTs>();
		}

		/// <summary>
		/// Wrapper interface for MsoDebugOptions_UTManager which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMsoDebugOptions_UTManager WithComCleanupProxy(this Microsoft.Office.Core.MsoDebugOptions_UTManager resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.MsoDebugOptions_UTManager, Interfaces.IMsoDebugOptions_UTManager>();
		}

		/// <summary>
		/// Wrapper interface for MetaProperty which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMetaProperty WithComCleanupProxy(this Microsoft.Office.Core.MetaProperty resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.MetaProperty, Interfaces.IMetaProperty>();
		}

		/// <summary>
		/// Wrapper interface for MetaProperties which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMetaProperties WithComCleanupProxy(this Microsoft.Office.Core.MetaProperties resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.MetaProperties, Interfaces.IMetaProperties>();
		}

		/// <summary>
		/// Wrapper interface for PolicyItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPolicyItem WithComCleanupProxy(this Microsoft.Office.Core.PolicyItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.PolicyItem, Interfaces.IPolicyItem>();
		}

		/// <summary>
		/// Wrapper interface for ServerPolicy which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IServerPolicy WithComCleanupProxy(this Microsoft.Office.Core.ServerPolicy resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ServerPolicy, Interfaces.IServerPolicy>();
		}

		/// <summary>
		/// Wrapper interface for DocumentInspector which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocumentInspector WithComCleanupProxy(this Microsoft.Office.Core.DocumentInspector resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.DocumentInspector, Interfaces.IDocumentInspector>();
		}

		/// <summary>
		/// Wrapper interface for DocumentInspectors which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocumentInspectors WithComCleanupProxy(this Microsoft.Office.Core.DocumentInspectors resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.DocumentInspectors, Interfaces.IDocumentInspectors>();
		}

		/// <summary>
		/// Wrapper interface for WorkflowTask which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWorkflowTask WithComCleanupProxy(this Microsoft.Office.Core.WorkflowTask resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.WorkflowTask, Interfaces.IWorkflowTask>();
		}

		/// <summary>
		/// Wrapper interface for WorkflowTasks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWorkflowTasks WithComCleanupProxy(this Microsoft.Office.Core.WorkflowTasks resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.WorkflowTasks, Interfaces.IWorkflowTasks>();
		}

		/// <summary>
		/// Wrapper interface for WorkflowTemplate which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWorkflowTemplate WithComCleanupProxy(this Microsoft.Office.Core.WorkflowTemplate resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.WorkflowTemplate, Interfaces.IWorkflowTemplate>();
		}

		/// <summary>
		/// Wrapper interface for WorkflowTemplates which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWorkflowTemplates WithComCleanupProxy(this Microsoft.Office.Core.WorkflowTemplates resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.WorkflowTemplates, Interfaces.IWorkflowTemplates>();
		}

		/// <summary>
		/// Wrapper interface for IDocumentInspector which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDocumentInspector WithComCleanupProxy(this Microsoft.Office.Core.IDocumentInspector resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IDocumentInspector, Interfaces.IIDocumentInspector>();
		}

		/// <summary>
		/// Wrapper interface for SignatureSetup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISignatureSetup WithComCleanupProxy(this Microsoft.Office.Core.SignatureSetup resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SignatureSetup, Interfaces.ISignatureSetup>();
		}

		/// <summary>
		/// Wrapper interface for SignatureInfo which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISignatureInfo WithComCleanupProxy(this Microsoft.Office.Core.SignatureInfo resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SignatureInfo, Interfaces.ISignatureInfo>();
		}

		/// <summary>
		/// Wrapper interface for SignatureProvider which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISignatureProvider WithComCleanupProxy(this Microsoft.Office.Core.SignatureProvider resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SignatureProvider, Interfaces.ISignatureProvider>();
		}

		/// <summary>
		/// Wrapper interface for CustomXMLPrefixMapping which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomXMLPrefixMapping WithComCleanupProxy(this Microsoft.Office.Core.CustomXMLPrefixMapping resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.CustomXMLPrefixMapping, Interfaces.ICustomXMLPrefixMapping>();
		}

		/// <summary>
		/// Wrapper interface for CustomXMLPrefixMappings which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomXMLPrefixMappings WithComCleanupProxy(this Microsoft.Office.Core.CustomXMLPrefixMappings resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.CustomXMLPrefixMappings, Interfaces.ICustomXMLPrefixMappings>();
		}

		/// <summary>
		/// Wrapper interface for CustomXMLSchema which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomXMLSchema WithComCleanupProxy(this Microsoft.Office.Core.CustomXMLSchema resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.CustomXMLSchema, Interfaces.ICustomXMLSchema>();
		}

		/// <summary>
		/// Wrapper interface for _CustomXMLSchemaCollection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CustomXMLSchemaCollection WithComCleanupProxy(this Microsoft.Office.Core._CustomXMLSchemaCollection resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._CustomXMLSchemaCollection, Interfaces.I_CustomXMLSchemaCollection>();
		}

		/// <summary>
		/// Wrapper interface for CustomXMLSchemaCollection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomXMLSchemaCollection WithComCleanupProxy(this Microsoft.Office.Core.CustomXMLSchemaCollection resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.CustomXMLSchemaCollection, Interfaces.ICustomXMLSchemaCollection>();
		}

		/// <summary>
		/// Wrapper interface for CustomXMLNodes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomXMLNodes WithComCleanupProxy(this Microsoft.Office.Core.CustomXMLNodes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.CustomXMLNodes, Interfaces.ICustomXMLNodes>();
		}

		/// <summary>
		/// Wrapper interface for CustomXMLNode which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomXMLNode WithComCleanupProxy(this Microsoft.Office.Core.CustomXMLNode resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.CustomXMLNode, Interfaces.ICustomXMLNode>();
		}

		/// <summary>
		/// Wrapper interface for CustomXMLValidationError which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomXMLValidationError WithComCleanupProxy(this Microsoft.Office.Core.CustomXMLValidationError resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.CustomXMLValidationError, Interfaces.ICustomXMLValidationError>();
		}

		/// <summary>
		/// Wrapper interface for CustomXMLValidationErrors which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomXMLValidationErrors WithComCleanupProxy(this Microsoft.Office.Core.CustomXMLValidationErrors resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.CustomXMLValidationErrors, Interfaces.ICustomXMLValidationErrors>();
		}

		/// <summary>
		/// Wrapper interface for _CustomXMLPart which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CustomXMLPart WithComCleanupProxy(this Microsoft.Office.Core._CustomXMLPart resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._CustomXMLPart, Interfaces.I_CustomXMLPart>();
		}

		/// <summary>
		/// Wrapper interface for ICustomXMLPartEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICustomXMLPartEvents WithComCleanupProxy(this Microsoft.Office.Core.ICustomXMLPartEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ICustomXMLPartEvents, Interfaces.IICustomXMLPartEvents>();
		}

		/// <summary>
		/// Wrapper interface for _CustomXMLPartEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CustomXMLPartEvents WithComCleanupProxy(this Microsoft.Office.Core._CustomXMLPartEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._CustomXMLPartEvents, Interfaces.I_CustomXMLPartEvents>();
		}

		/// <summary>
		/// Wrapper interface for _CustomXMLPartEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CustomXMLPartEvents_Event WithComCleanupProxy(this Microsoft.Office.Core._CustomXMLPartEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._CustomXMLPartEvents_Event, Interfaces.I_CustomXMLPartEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for CustomXMLPart which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomXMLPart WithComCleanupProxy(this Microsoft.Office.Core.CustomXMLPart resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.CustomXMLPart, Interfaces.ICustomXMLPart>();
		}

		/// <summary>
		/// Wrapper interface for _CustomXMLParts which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CustomXMLParts WithComCleanupProxy(this Microsoft.Office.Core._CustomXMLParts resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._CustomXMLParts, Interfaces.I_CustomXMLParts>();
		}

		/// <summary>
		/// Wrapper interface for ICustomXMLPartsEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICustomXMLPartsEvents WithComCleanupProxy(this Microsoft.Office.Core.ICustomXMLPartsEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ICustomXMLPartsEvents, Interfaces.IICustomXMLPartsEvents>();
		}

		/// <summary>
		/// Wrapper interface for _CustomXMLPartsEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CustomXMLPartsEvents WithComCleanupProxy(this Microsoft.Office.Core._CustomXMLPartsEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._CustomXMLPartsEvents, Interfaces.I_CustomXMLPartsEvents>();
		}

		/// <summary>
		/// Wrapper interface for _CustomXMLPartsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CustomXMLPartsEvents_Event WithComCleanupProxy(this Microsoft.Office.Core._CustomXMLPartsEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._CustomXMLPartsEvents_Event, Interfaces.I_CustomXMLPartsEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for CustomXMLParts which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomXMLParts WithComCleanupProxy(this Microsoft.Office.Core.CustomXMLParts resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.CustomXMLParts, Interfaces.ICustomXMLParts>();
		}

		/// <summary>
		/// Wrapper interface for GradientStop which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGradientStop WithComCleanupProxy(this Microsoft.Office.Core.GradientStop resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.GradientStop, Interfaces.IGradientStop>();
		}

		/// <summary>
		/// Wrapper interface for GradientStops which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGradientStops WithComCleanupProxy(this Microsoft.Office.Core.GradientStops resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.GradientStops, Interfaces.IGradientStops>();
		}

		/// <summary>
		/// Wrapper interface for SoftEdgeFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISoftEdgeFormat WithComCleanupProxy(this Microsoft.Office.Core.SoftEdgeFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SoftEdgeFormat, Interfaces.ISoftEdgeFormat>();
		}

		/// <summary>
		/// Wrapper interface for GlowFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGlowFormat WithComCleanupProxy(this Microsoft.Office.Core.GlowFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.GlowFormat, Interfaces.IGlowFormat>();
		}

		/// <summary>
		/// Wrapper interface for ReflectionFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IReflectionFormat WithComCleanupProxy(this Microsoft.Office.Core.ReflectionFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ReflectionFormat, Interfaces.IReflectionFormat>();
		}

		/// <summary>
		/// Wrapper interface for ParagraphFormat2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IParagraphFormat2 WithComCleanupProxy(this Microsoft.Office.Core.ParagraphFormat2 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ParagraphFormat2, Interfaces.IParagraphFormat2>();
		}

		/// <summary>
		/// Wrapper interface for Font2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFont2 WithComCleanupProxy(this Microsoft.Office.Core.Font2 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.Font2, Interfaces.IFont2>();
		}

		/// <summary>
		/// Wrapper interface for TextColumn2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextColumn2 WithComCleanupProxy(this Microsoft.Office.Core.TextColumn2 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.TextColumn2, Interfaces.ITextColumn2>();
		}

		/// <summary>
		/// Wrapper interface for TextRange2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextRange2 WithComCleanupProxy(this Microsoft.Office.Core.TextRange2 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.TextRange2, Interfaces.ITextRange2>();
		}

		/// <summary>
		/// Wrapper interface for TextFrame2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextFrame2 WithComCleanupProxy(this Microsoft.Office.Core.TextFrame2 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.TextFrame2, Interfaces.ITextFrame2>();
		}

		/// <summary>
		/// Wrapper interface for ThemeColor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IThemeColor WithComCleanupProxy(this Microsoft.Office.Core.ThemeColor resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ThemeColor, Interfaces.IThemeColor>();
		}

		/// <summary>
		/// Wrapper interface for ThemeColorScheme which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IThemeColorScheme WithComCleanupProxy(this Microsoft.Office.Core.ThemeColorScheme resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ThemeColorScheme, Interfaces.IThemeColorScheme>();
		}

		/// <summary>
		/// Wrapper interface for ThemeFont which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IThemeFont WithComCleanupProxy(this Microsoft.Office.Core.ThemeFont resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ThemeFont, Interfaces.IThemeFont>();
		}

		/// <summary>
		/// Wrapper interface for ThemeFonts which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IThemeFonts WithComCleanupProxy(this Microsoft.Office.Core.ThemeFonts resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ThemeFonts, Interfaces.IThemeFonts>();
		}

		/// <summary>
		/// Wrapper interface for ThemeFontScheme which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IThemeFontScheme WithComCleanupProxy(this Microsoft.Office.Core.ThemeFontScheme resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ThemeFontScheme, Interfaces.IThemeFontScheme>();
		}

		/// <summary>
		/// Wrapper interface for ThemeEffectScheme which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IThemeEffectScheme WithComCleanupProxy(this Microsoft.Office.Core.ThemeEffectScheme resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ThemeEffectScheme, Interfaces.IThemeEffectScheme>();
		}

		/// <summary>
		/// Wrapper interface for OfficeTheme which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOfficeTheme WithComCleanupProxy(this Microsoft.Office.Core.OfficeTheme resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.OfficeTheme, Interfaces.IOfficeTheme>();
		}

		/// <summary>
		/// Wrapper interface for _CustomTaskPane which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CustomTaskPane WithComCleanupProxy(this Microsoft.Office.Core._CustomTaskPane resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._CustomTaskPane, Interfaces.I_CustomTaskPane>();
		}

		/// <summary>
		/// Wrapper interface for CustomTaskPaneEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomTaskPaneEvents WithComCleanupProxy(this Microsoft.Office.Core.CustomTaskPaneEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.CustomTaskPaneEvents, Interfaces.ICustomTaskPaneEvents>();
		}

		/// <summary>
		/// Wrapper interface for _CustomTaskPaneEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CustomTaskPaneEvents WithComCleanupProxy(this Microsoft.Office.Core._CustomTaskPaneEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._CustomTaskPaneEvents, Interfaces.I_CustomTaskPaneEvents>();
		}

		/// <summary>
		/// Wrapper interface for _CustomTaskPaneEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_CustomTaskPaneEvents_Event WithComCleanupProxy(this Microsoft.Office.Core._CustomTaskPaneEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core._CustomTaskPaneEvents_Event, Interfaces.I_CustomTaskPaneEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for CustomTaskPane which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomTaskPane WithComCleanupProxy(this Microsoft.Office.Core.CustomTaskPane resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.CustomTaskPane, Interfaces.ICustomTaskPane>();
		}

		/// <summary>
		/// Wrapper interface for ICTPFactory which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICTPFactory WithComCleanupProxy(this Microsoft.Office.Core.ICTPFactory resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ICTPFactory, Interfaces.IICTPFactory>();
		}

		/// <summary>
		/// Wrapper interface for ICustomTaskPaneConsumer which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICustomTaskPaneConsumer WithComCleanupProxy(this Microsoft.Office.Core.ICustomTaskPaneConsumer resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ICustomTaskPaneConsumer, Interfaces.IICustomTaskPaneConsumer>();
		}

		/// <summary>
		/// Wrapper interface for IRibbonUI which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIRibbonUI WithComCleanupProxy(this Microsoft.Office.Core.IRibbonUI resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IRibbonUI, Interfaces.IIRibbonUI>();
		}

		/// <summary>
		/// Wrapper interface for IRibbonControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIRibbonControl WithComCleanupProxy(this Microsoft.Office.Core.IRibbonControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IRibbonControl, Interfaces.IIRibbonControl>();
		}

		/// <summary>
		/// Wrapper interface for IRibbonExtensibility which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIRibbonExtensibility WithComCleanupProxy(this Microsoft.Office.Core.IRibbonExtensibility resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IRibbonExtensibility, Interfaces.IIRibbonExtensibility>();
		}

		/// <summary>
		/// Wrapper interface for IAssistance which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIAssistance WithComCleanupProxy(this Microsoft.Office.Core.IAssistance resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IAssistance, Interfaces.IIAssistance>();
		}

		/// <summary>
		/// Wrapper interface for IMsoChartData which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoChartData WithComCleanupProxy(this Microsoft.Office.Core.IMsoChartData resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoChartData, Interfaces.IIMsoChartData>();
		}

		/// <summary>
		/// Wrapper interface for IMsoChart which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoChart WithComCleanupProxy(this Microsoft.Office.Core.IMsoChart resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoChart, Interfaces.IIMsoChart>();
		}

		/// <summary>
		/// Wrapper interface for IMsoCorners which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoCorners WithComCleanupProxy(this Microsoft.Office.Core.IMsoCorners resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoCorners, Interfaces.IIMsoCorners>();
		}

		/// <summary>
		/// Wrapper interface for IMsoLegend which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoLegend WithComCleanupProxy(this Microsoft.Office.Core.IMsoLegend resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoLegend, Interfaces.IIMsoLegend>();
		}

		/// <summary>
		/// Wrapper interface for IMsoBorder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoBorder WithComCleanupProxy(this Microsoft.Office.Core.IMsoBorder resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoBorder, Interfaces.IIMsoBorder>();
		}

		/// <summary>
		/// Wrapper interface for IMsoWalls which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoWalls WithComCleanupProxy(this Microsoft.Office.Core.IMsoWalls resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoWalls, Interfaces.IIMsoWalls>();
		}

		/// <summary>
		/// Wrapper interface for IMsoFloor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoFloor WithComCleanupProxy(this Microsoft.Office.Core.IMsoFloor resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoFloor, Interfaces.IIMsoFloor>();
		}

		/// <summary>
		/// Wrapper interface for IMsoPlotArea which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoPlotArea WithComCleanupProxy(this Microsoft.Office.Core.IMsoPlotArea resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoPlotArea, Interfaces.IIMsoPlotArea>();
		}

		/// <summary>
		/// Wrapper interface for IMsoChartArea which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoChartArea WithComCleanupProxy(this Microsoft.Office.Core.IMsoChartArea resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoChartArea, Interfaces.IIMsoChartArea>();
		}

		/// <summary>
		/// Wrapper interface for IMsoSeriesLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoSeriesLines WithComCleanupProxy(this Microsoft.Office.Core.IMsoSeriesLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoSeriesLines, Interfaces.IIMsoSeriesLines>();
		}

		/// <summary>
		/// Wrapper interface for IMsoLeaderLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoLeaderLines WithComCleanupProxy(this Microsoft.Office.Core.IMsoLeaderLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoLeaderLines, Interfaces.IIMsoLeaderLines>();
		}

		/// <summary>
		/// Wrapper interface for GridLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGridLines WithComCleanupProxy(this Microsoft.Office.Core.GridLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.GridLines, Interfaces.IGridLines>();
		}

		/// <summary>
		/// Wrapper interface for IMsoUpBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoUpBars WithComCleanupProxy(this Microsoft.Office.Core.IMsoUpBars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoUpBars, Interfaces.IIMsoUpBars>();
		}

		/// <summary>
		/// Wrapper interface for IMsoDownBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoDownBars WithComCleanupProxy(this Microsoft.Office.Core.IMsoDownBars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoDownBars, Interfaces.IIMsoDownBars>();
		}

		/// <summary>
		/// Wrapper interface for IMsoInterior which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoInterior WithComCleanupProxy(this Microsoft.Office.Core.IMsoInterior resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoInterior, Interfaces.IIMsoInterior>();
		}

		/// <summary>
		/// Wrapper interface for ChartFillFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartFillFormat WithComCleanupProxy(this Microsoft.Office.Core.ChartFillFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ChartFillFormat, Interfaces.IChartFillFormat>();
		}

		/// <summary>
		/// Wrapper interface for LegendEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegendEntries WithComCleanupProxy(this Microsoft.Office.Core.LegendEntries resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.LegendEntries, Interfaces.ILegendEntries>();
		}

		/// <summary>
		/// Wrapper interface for ChartFont which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartFont WithComCleanupProxy(this Microsoft.Office.Core.ChartFont resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ChartFont, Interfaces.IChartFont>();
		}

		/// <summary>
		/// Wrapper interface for ChartColorFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartColorFormat WithComCleanupProxy(this Microsoft.Office.Core.ChartColorFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ChartColorFormat, Interfaces.IChartColorFormat>();
		}

		/// <summary>
		/// Wrapper interface for LegendEntry which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegendEntry WithComCleanupProxy(this Microsoft.Office.Core.LegendEntry resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.LegendEntry, Interfaces.ILegendEntry>();
		}

		/// <summary>
		/// Wrapper interface for IMsoLegendKey which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoLegendKey WithComCleanupProxy(this Microsoft.Office.Core.IMsoLegendKey resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoLegendKey, Interfaces.IIMsoLegendKey>();
		}

		/// <summary>
		/// Wrapper interface for SeriesCollection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISeriesCollection WithComCleanupProxy(this Microsoft.Office.Core.SeriesCollection resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SeriesCollection, Interfaces.ISeriesCollection>();
		}

		/// <summary>
		/// Wrapper interface for IMsoSeries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoSeries WithComCleanupProxy(this Microsoft.Office.Core.IMsoSeries resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoSeries, Interfaces.IIMsoSeries>();
		}

		/// <summary>
		/// Wrapper interface for IMsoErrorBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoErrorBars WithComCleanupProxy(this Microsoft.Office.Core.IMsoErrorBars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoErrorBars, Interfaces.IIMsoErrorBars>();
		}

		/// <summary>
		/// Wrapper interface for IMsoTrendline which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoTrendline WithComCleanupProxy(this Microsoft.Office.Core.IMsoTrendline resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoTrendline, Interfaces.IIMsoTrendline>();
		}

		/// <summary>
		/// Wrapper interface for Trendlines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITrendlines WithComCleanupProxy(this Microsoft.Office.Core.Trendlines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.Trendlines, Interfaces.ITrendlines>();
		}

		/// <summary>
		/// Wrapper interface for IMsoDataLabels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoDataLabels WithComCleanupProxy(this Microsoft.Office.Core.IMsoDataLabels resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoDataLabels, Interfaces.IIMsoDataLabels>();
		}

		/// <summary>
		/// Wrapper interface for IMsoDataLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoDataLabel WithComCleanupProxy(this Microsoft.Office.Core.IMsoDataLabel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoDataLabel, Interfaces.IIMsoDataLabel>();
		}

		/// <summary>
		/// Wrapper interface for Points which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPoints WithComCleanupProxy(this Microsoft.Office.Core.Points resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.Points, Interfaces.IPoints>();
		}

		/// <summary>
		/// Wrapper interface for ChartPoint which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartPoint WithComCleanupProxy(this Microsoft.Office.Core.ChartPoint resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ChartPoint, Interfaces.IChartPoint>();
		}

		/// <summary>
		/// Wrapper interface for Axes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAxes WithComCleanupProxy(this Microsoft.Office.Core.Axes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.Axes, Interfaces.IAxes>();
		}

		/// <summary>
		/// Wrapper interface for IMsoAxis which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoAxis WithComCleanupProxy(this Microsoft.Office.Core.IMsoAxis resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoAxis, Interfaces.IIMsoAxis>();
		}

		/// <summary>
		/// Wrapper interface for IMsoDataTable which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoDataTable WithComCleanupProxy(this Microsoft.Office.Core.IMsoDataTable resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoDataTable, Interfaces.IIMsoDataTable>();
		}

		/// <summary>
		/// Wrapper interface for IMsoChartTitle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoChartTitle WithComCleanupProxy(this Microsoft.Office.Core.IMsoChartTitle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoChartTitle, Interfaces.IIMsoChartTitle>();
		}

		/// <summary>
		/// Wrapper interface for IMsoAxisTitle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoAxisTitle WithComCleanupProxy(this Microsoft.Office.Core.IMsoAxisTitle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoAxisTitle, Interfaces.IIMsoAxisTitle>();
		}

		/// <summary>
		/// Wrapper interface for IMsoDisplayUnitLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoDisplayUnitLabel WithComCleanupProxy(this Microsoft.Office.Core.IMsoDisplayUnitLabel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoDisplayUnitLabel, Interfaces.IIMsoDisplayUnitLabel>();
		}

		/// <summary>
		/// Wrapper interface for IMsoTickLabels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoTickLabels WithComCleanupProxy(this Microsoft.Office.Core.IMsoTickLabels resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoTickLabels, Interfaces.IIMsoTickLabels>();
		}

		/// <summary>
		/// Wrapper interface for IMsoHyperlinks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoHyperlinks WithComCleanupProxy(this Microsoft.Office.Core.IMsoHyperlinks resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoHyperlinks, Interfaces.IIMsoHyperlinks>();
		}

		/// <summary>
		/// Wrapper interface for IMsoDropLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoDropLines WithComCleanupProxy(this Microsoft.Office.Core.IMsoDropLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoDropLines, Interfaces.IIMsoDropLines>();
		}

		/// <summary>
		/// Wrapper interface for IMsoHiLoLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoHiLoLines WithComCleanupProxy(this Microsoft.Office.Core.IMsoHiLoLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoHiLoLines, Interfaces.IIMsoHiLoLines>();
		}

		/// <summary>
		/// Wrapper interface for IMsoChartGroup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoChartGroup WithComCleanupProxy(this Microsoft.Office.Core.IMsoChartGroup resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoChartGroup, Interfaces.IIMsoChartGroup>();
		}

		/// <summary>
		/// Wrapper interface for ChartGroups which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartGroups WithComCleanupProxy(this Microsoft.Office.Core.ChartGroups resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ChartGroups, Interfaces.IChartGroups>();
		}

		/// <summary>
		/// Wrapper interface for IMsoCharacters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoCharacters WithComCleanupProxy(this Microsoft.Office.Core.IMsoCharacters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoCharacters, Interfaces.IIMsoCharacters>();
		}

		/// <summary>
		/// Wrapper interface for IMsoChartFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoChartFormat WithComCleanupProxy(this Microsoft.Office.Core.IMsoChartFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoChartFormat, Interfaces.IIMsoChartFormat>();
		}

		/// <summary>
		/// Wrapper interface for BulletFormat2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBulletFormat2 WithComCleanupProxy(this Microsoft.Office.Core.BulletFormat2 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.BulletFormat2, Interfaces.IBulletFormat2>();
		}

		/// <summary>
		/// Wrapper interface for TabStops2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITabStops2 WithComCleanupProxy(this Microsoft.Office.Core.TabStops2 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.TabStops2, Interfaces.ITabStops2>();
		}

		/// <summary>
		/// Wrapper interface for TabStop2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITabStop2 WithComCleanupProxy(this Microsoft.Office.Core.TabStop2 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.TabStop2, Interfaces.ITabStop2>();
		}

		/// <summary>
		/// Wrapper interface for Ruler2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRuler2 WithComCleanupProxy(this Microsoft.Office.Core.Ruler2 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.Ruler2, Interfaces.IRuler2>();
		}

		/// <summary>
		/// Wrapper interface for RulerLevels2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRulerLevels2 WithComCleanupProxy(this Microsoft.Office.Core.RulerLevels2 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.RulerLevels2, Interfaces.IRulerLevels2>();
		}

		/// <summary>
		/// Wrapper interface for RulerLevel2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRulerLevel2 WithComCleanupProxy(this Microsoft.Office.Core.RulerLevel2 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.RulerLevel2, Interfaces.IRulerLevel2>();
		}

		/// <summary>
		/// Wrapper interface for EncryptionProvider which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEncryptionProvider WithComCleanupProxy(this Microsoft.Office.Core.EncryptionProvider resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.EncryptionProvider, Interfaces.IEncryptionProvider>();
		}

		/// <summary>
		/// Wrapper interface for IBlogExtensibility which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIBlogExtensibility WithComCleanupProxy(this Microsoft.Office.Core.IBlogExtensibility resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IBlogExtensibility, Interfaces.IIBlogExtensibility>();
		}

		/// <summary>
		/// Wrapper interface for IBlogPictureExtensibility which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIBlogPictureExtensibility WithComCleanupProxy(this Microsoft.Office.Core.IBlogPictureExtensibility resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IBlogPictureExtensibility, Interfaces.IIBlogPictureExtensibility>();
		}

		/// <summary>
		/// Wrapper interface for IConverterPreferences which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIConverterPreferences WithComCleanupProxy(this Microsoft.Office.Core.IConverterPreferences resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IConverterPreferences, Interfaces.IIConverterPreferences>();
		}

		/// <summary>
		/// Wrapper interface for IConverterApplicationPreferences which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIConverterApplicationPreferences WithComCleanupProxy(this Microsoft.Office.Core.IConverterApplicationPreferences resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IConverterApplicationPreferences, Interfaces.IIConverterApplicationPreferences>();
		}

		/// <summary>
		/// Wrapper interface for IConverterUICallback which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIConverterUICallback WithComCleanupProxy(this Microsoft.Office.Core.IConverterUICallback resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IConverterUICallback, Interfaces.IIConverterUICallback>();
		}

		/// <summary>
		/// Wrapper interface for IConverter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIConverter WithComCleanupProxy(this Microsoft.Office.Core.IConverter resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IConverter, Interfaces.IIConverter>();
		}

		/// <summary>
		/// Wrapper interface for SmartArt which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartArt WithComCleanupProxy(this Microsoft.Office.Core.SmartArt resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SmartArt, Interfaces.ISmartArt>();
		}

		/// <summary>
		/// Wrapper interface for SmartArtNodes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartArtNodes WithComCleanupProxy(this Microsoft.Office.Core.SmartArtNodes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SmartArtNodes, Interfaces.ISmartArtNodes>();
		}

		/// <summary>
		/// Wrapper interface for SmartArtNode which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartArtNode WithComCleanupProxy(this Microsoft.Office.Core.SmartArtNode resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SmartArtNode, Interfaces.ISmartArtNode>();
		}

		/// <summary>
		/// Wrapper interface for SmartArtLayouts which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartArtLayouts WithComCleanupProxy(this Microsoft.Office.Core.SmartArtLayouts resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SmartArtLayouts, Interfaces.ISmartArtLayouts>();
		}

		/// <summary>
		/// Wrapper interface for SmartArtLayout which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartArtLayout WithComCleanupProxy(this Microsoft.Office.Core.SmartArtLayout resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SmartArtLayout, Interfaces.ISmartArtLayout>();
		}

		/// <summary>
		/// Wrapper interface for SmartArtQuickStyles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartArtQuickStyles WithComCleanupProxy(this Microsoft.Office.Core.SmartArtQuickStyles resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SmartArtQuickStyles, Interfaces.ISmartArtQuickStyles>();
		}

		/// <summary>
		/// Wrapper interface for SmartArtQuickStyle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartArtQuickStyle WithComCleanupProxy(this Microsoft.Office.Core.SmartArtQuickStyle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SmartArtQuickStyle, Interfaces.ISmartArtQuickStyle>();
		}

		/// <summary>
		/// Wrapper interface for SmartArtColors which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartArtColors WithComCleanupProxy(this Microsoft.Office.Core.SmartArtColors resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SmartArtColors, Interfaces.ISmartArtColors>();
		}

		/// <summary>
		/// Wrapper interface for SmartArtColor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartArtColor WithComCleanupProxy(this Microsoft.Office.Core.SmartArtColor resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.SmartArtColor, Interfaces.ISmartArtColor>();
		}

		/// <summary>
		/// Wrapper interface for PickerField which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPickerField WithComCleanupProxy(this Microsoft.Office.Core.PickerField resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.PickerField, Interfaces.IPickerField>();
		}

		/// <summary>
		/// Wrapper interface for PickerFields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPickerFields WithComCleanupProxy(this Microsoft.Office.Core.PickerFields resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.PickerFields, Interfaces.IPickerFields>();
		}

		/// <summary>
		/// Wrapper interface for PickerProperty which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPickerProperty WithComCleanupProxy(this Microsoft.Office.Core.PickerProperty resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.PickerProperty, Interfaces.IPickerProperty>();
		}

		/// <summary>
		/// Wrapper interface for PickerProperties which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPickerProperties WithComCleanupProxy(this Microsoft.Office.Core.PickerProperties resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.PickerProperties, Interfaces.IPickerProperties>();
		}

		/// <summary>
		/// Wrapper interface for PickerResult which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPickerResult WithComCleanupProxy(this Microsoft.Office.Core.PickerResult resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.PickerResult, Interfaces.IPickerResult>();
		}

		/// <summary>
		/// Wrapper interface for PickerResults which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPickerResults WithComCleanupProxy(this Microsoft.Office.Core.PickerResults resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.PickerResults, Interfaces.IPickerResults>();
		}

		/// <summary>
		/// Wrapper interface for PickerDialog which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPickerDialog WithComCleanupProxy(this Microsoft.Office.Core.PickerDialog resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.PickerDialog, Interfaces.IPickerDialog>();
		}

		/// <summary>
		/// Wrapper interface for IMsoContactCard which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMsoContactCard WithComCleanupProxy(this Microsoft.Office.Core.IMsoContactCard resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.IMsoContactCard, Interfaces.IIMsoContactCard>();
		}

		/// <summary>
		/// Wrapper interface for EffectParameter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEffectParameter WithComCleanupProxy(this Microsoft.Office.Core.EffectParameter resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.EffectParameter, Interfaces.IEffectParameter>();
		}

		/// <summary>
		/// Wrapper interface for EffectParameters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEffectParameters WithComCleanupProxy(this Microsoft.Office.Core.EffectParameters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.EffectParameters, Interfaces.IEffectParameters>();
		}

		/// <summary>
		/// Wrapper interface for PictureEffect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPictureEffect WithComCleanupProxy(this Microsoft.Office.Core.PictureEffect resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.PictureEffect, Interfaces.IPictureEffect>();
		}

		/// <summary>
		/// Wrapper interface for PictureEffects which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPictureEffects WithComCleanupProxy(this Microsoft.Office.Core.PictureEffects resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.PictureEffects, Interfaces.IPictureEffects>();
		}

		/// <summary>
		/// Wrapper interface for Crop which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICrop WithComCleanupProxy(this Microsoft.Office.Core.Crop resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.Crop, Interfaces.ICrop>();
		}

		/// <summary>
		/// Wrapper interface for ContactCard which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IContactCard WithComCleanupProxy(this Microsoft.Office.Core.ContactCard resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Core.ContactCard, Interfaces.IContactCard>();
		}

	}
}