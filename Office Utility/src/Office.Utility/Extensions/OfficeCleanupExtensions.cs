using Office.Utility.Extensions;
using Microsoft.Office.Core;

namespace Office.Utility.Extensions
{
	/// <summary>
	/// Provides cleanup extension methods for interfaces exposed by Office.dll
	/// </summary>
	public static class OfficeCleanupExtensions
	{
		/// <summary>
		/// Wrapper interface for IAccessible which adds IDispose to the interface
		/// </summary>
		public static IIAccessible WithComCleanup(this IAccessible resource)
		{
			return resource.WithComCleanup<IAccessible, IIAccessible>();
		}

	/// <summary>
		/// Wrapper interface for _IMsoDispObj which adds IDispose to the interface
		/// </summary>
		public static I_IMsoDispObj WithComCleanup(this _IMsoDispObj resource)
		{
			return resource.WithComCleanup<_IMsoDispObj, I_IMsoDispObj>();
		}

	/// <summary>
		/// Wrapper interface for _IMsoOleAccDispObj which adds IDispose to the interface
		/// </summary>
		public static I_IMsoOleAccDispObj WithComCleanup(this _IMsoOleAccDispObj resource)
		{
			return resource.WithComCleanup<_IMsoOleAccDispObj, I_IMsoOleAccDispObj>();
		}

	/// <summary>
		/// Wrapper interface for _CommandBars which adds IDispose to the interface
		/// </summary>
		public static I_CommandBars WithComCleanup(this _CommandBars resource)
		{
			return resource.WithComCleanup<_CommandBars, I_CommandBars>();
		}

	/// <summary>
		/// Wrapper interface for CommandBar which adds IDispose to the interface
		/// </summary>
		public static ICommandBar WithComCleanup(this CommandBar resource)
		{
			return resource.WithComCleanup<CommandBar, ICommandBar>();
		}

	/// <summary>
		/// Wrapper interface for CommandBarControls which adds IDispose to the interface
		/// </summary>
		public static ICommandBarControls WithComCleanup(this CommandBarControls resource)
		{
			return resource.WithComCleanup<CommandBarControls, ICommandBarControls>();
		}

	/// <summary>
		/// Wrapper interface for CommandBarControl which adds IDispose to the interface
		/// </summary>
		public static ICommandBarControl WithComCleanup(this CommandBarControl resource)
		{
			return resource.WithComCleanup<CommandBarControl, ICommandBarControl>();
		}

	/// <summary>
		/// Wrapper interface for _CommandBarButton which adds IDispose to the interface
		/// </summary>
		public static I_CommandBarButton WithComCleanup(this _CommandBarButton resource)
		{
			return resource.WithComCleanup<_CommandBarButton, I_CommandBarButton>();
		}

	/// <summary>
		/// Wrapper interface for CommandBarPopup which adds IDispose to the interface
		/// </summary>
		public static ICommandBarPopup WithComCleanup(this CommandBarPopup resource)
		{
			return resource.WithComCleanup<CommandBarPopup, ICommandBarPopup>();
		}

	/// <summary>
		/// Wrapper interface for _CommandBarComboBox which adds IDispose to the interface
		/// </summary>
		public static I_CommandBarComboBox WithComCleanup(this _CommandBarComboBox resource)
		{
			return resource.WithComCleanup<_CommandBarComboBox, I_CommandBarComboBox>();
		}

	/// <summary>
		/// Wrapper interface for _CommandBarActiveX which adds IDispose to the interface
		/// </summary>
		public static I_CommandBarActiveX WithComCleanup(this _CommandBarActiveX resource)
		{
			return resource.WithComCleanup<_CommandBarActiveX, I_CommandBarActiveX>();
		}

	/// <summary>
		/// Wrapper interface for Adjustments which adds IDispose to the interface
		/// </summary>
		public static IAdjustments WithComCleanup(this Adjustments resource)
		{
			return resource.WithComCleanup<Adjustments, IAdjustments>();
		}

	/// <summary>
		/// Wrapper interface for CalloutFormat which adds IDispose to the interface
		/// </summary>
		public static ICalloutFormat WithComCleanup(this CalloutFormat resource)
		{
			return resource.WithComCleanup<CalloutFormat, ICalloutFormat>();
		}

	/// <summary>
		/// Wrapper interface for ColorFormat which adds IDispose to the interface
		/// </summary>
		public static IColorFormat WithComCleanup(this ColorFormat resource)
		{
			return resource.WithComCleanup<ColorFormat, IColorFormat>();
		}

	/// <summary>
		/// Wrapper interface for ConnectorFormat which adds IDispose to the interface
		/// </summary>
		public static IConnectorFormat WithComCleanup(this ConnectorFormat resource)
		{
			return resource.WithComCleanup<ConnectorFormat, IConnectorFormat>();
		}

	/// <summary>
		/// Wrapper interface for FillFormat which adds IDispose to the interface
		/// </summary>
		public static IFillFormat WithComCleanup(this FillFormat resource)
		{
			return resource.WithComCleanup<FillFormat, IFillFormat>();
		}

	/// <summary>
		/// Wrapper interface for FreeformBuilder which adds IDispose to the interface
		/// </summary>
		public static IFreeformBuilder WithComCleanup(this FreeformBuilder resource)
		{
			return resource.WithComCleanup<FreeformBuilder, IFreeformBuilder>();
		}

	/// <summary>
		/// Wrapper interface for GroupShapes which adds IDispose to the interface
		/// </summary>
		public static IGroupShapes WithComCleanup(this GroupShapes resource)
		{
			return resource.WithComCleanup<GroupShapes, IGroupShapes>();
		}

	/// <summary>
		/// Wrapper interface for LineFormat which adds IDispose to the interface
		/// </summary>
		public static ILineFormat WithComCleanup(this LineFormat resource)
		{
			return resource.WithComCleanup<LineFormat, ILineFormat>();
		}

	/// <summary>
		/// Wrapper interface for ShapeNode which adds IDispose to the interface
		/// </summary>
		public static IShapeNode WithComCleanup(this ShapeNode resource)
		{
			return resource.WithComCleanup<ShapeNode, IShapeNode>();
		}

	/// <summary>
		/// Wrapper interface for ShapeNodes which adds IDispose to the interface
		/// </summary>
		public static IShapeNodes WithComCleanup(this ShapeNodes resource)
		{
			return resource.WithComCleanup<ShapeNodes, IShapeNodes>();
		}

	/// <summary>
		/// Wrapper interface for PictureFormat which adds IDispose to the interface
		/// </summary>
		public static IPictureFormat WithComCleanup(this PictureFormat resource)
		{
			return resource.WithComCleanup<PictureFormat, IPictureFormat>();
		}

	/// <summary>
		/// Wrapper interface for ShadowFormat which adds IDispose to the interface
		/// </summary>
		public static IShadowFormat WithComCleanup(this ShadowFormat resource)
		{
			return resource.WithComCleanup<ShadowFormat, IShadowFormat>();
		}

	/// <summary>
		/// Wrapper interface for Script which adds IDispose to the interface
		/// </summary>
		public static IScript WithComCleanup(this Script resource)
		{
			return resource.WithComCleanup<Script, IScript>();
		}

	/// <summary>
		/// Wrapper interface for Scripts which adds IDispose to the interface
		/// </summary>
		public static IScripts WithComCleanup(this Scripts resource)
		{
			return resource.WithComCleanup<Scripts, IScripts>();
		}

	/// <summary>
		/// Wrapper interface for Shape which adds IDispose to the interface
		/// </summary>
		public static IShape WithComCleanup(this Shape resource)
		{
			return resource.WithComCleanup<Shape, IShape>();
		}

	/// <summary>
		/// Wrapper interface for ShapeRange which adds IDispose to the interface
		/// </summary>
		public static IShapeRange WithComCleanup(this ShapeRange resource)
		{
			return resource.WithComCleanup<ShapeRange, IShapeRange>();
		}

	/// <summary>
		/// Wrapper interface for Shapes which adds IDispose to the interface
		/// </summary>
		public static IShapes WithComCleanup(this Shapes resource)
		{
			return resource.WithComCleanup<Shapes, IShapes>();
		}

	/// <summary>
		/// Wrapper interface for TextEffectFormat which adds IDispose to the interface
		/// </summary>
		public static ITextEffectFormat WithComCleanup(this TextEffectFormat resource)
		{
			return resource.WithComCleanup<TextEffectFormat, ITextEffectFormat>();
		}

	/// <summary>
		/// Wrapper interface for TextFrame which adds IDispose to the interface
		/// </summary>
		public static ITextFrame WithComCleanup(this TextFrame resource)
		{
			return resource.WithComCleanup<TextFrame, ITextFrame>();
		}

	/// <summary>
		/// Wrapper interface for ThreeDFormat which adds IDispose to the interface
		/// </summary>
		public static IThreeDFormat WithComCleanup(this ThreeDFormat resource)
		{
			return resource.WithComCleanup<ThreeDFormat, IThreeDFormat>();
		}

	/// <summary>
		/// Wrapper interface for IMsoDispCagNotifySink which adds IDispose to the interface
		/// </summary>
		public static IIMsoDispCagNotifySink WithComCleanup(this IMsoDispCagNotifySink resource)
		{
			return resource.WithComCleanup<IMsoDispCagNotifySink, IIMsoDispCagNotifySink>();
		}

	/// <summary>
		/// Wrapper interface for Balloon which adds IDispose to the interface
		/// </summary>
		public static IBalloon WithComCleanup(this Balloon resource)
		{
			return resource.WithComCleanup<Balloon, IBalloon>();
		}

	/// <summary>
		/// Wrapper interface for BalloonCheckboxes which adds IDispose to the interface
		/// </summary>
		public static IBalloonCheckboxes WithComCleanup(this BalloonCheckboxes resource)
		{
			return resource.WithComCleanup<BalloonCheckboxes, IBalloonCheckboxes>();
		}

	/// <summary>
		/// Wrapper interface for BalloonCheckbox which adds IDispose to the interface
		/// </summary>
		public static IBalloonCheckbox WithComCleanup(this BalloonCheckbox resource)
		{
			return resource.WithComCleanup<BalloonCheckbox, IBalloonCheckbox>();
		}

	/// <summary>
		/// Wrapper interface for BalloonLabels which adds IDispose to the interface
		/// </summary>
		public static IBalloonLabels WithComCleanup(this BalloonLabels resource)
		{
			return resource.WithComCleanup<BalloonLabels, IBalloonLabels>();
		}

	/// <summary>
		/// Wrapper interface for BalloonLabel which adds IDispose to the interface
		/// </summary>
		public static IBalloonLabel WithComCleanup(this BalloonLabel resource)
		{
			return resource.WithComCleanup<BalloonLabel, IBalloonLabel>();
		}

	/// <summary>
		/// Wrapper interface for AnswerWizardFiles which adds IDispose to the interface
		/// </summary>
		public static IAnswerWizardFiles WithComCleanup(this AnswerWizardFiles resource)
		{
			return resource.WithComCleanup<AnswerWizardFiles, IAnswerWizardFiles>();
		}

	/// <summary>
		/// Wrapper interface for AnswerWizard which adds IDispose to the interface
		/// </summary>
		public static IAnswerWizard WithComCleanup(this AnswerWizard resource)
		{
			return resource.WithComCleanup<AnswerWizard, IAnswerWizard>();
		}

	/// <summary>
		/// Wrapper interface for Assistant which adds IDispose to the interface
		/// </summary>
		public static IAssistant WithComCleanup(this Assistant resource)
		{
			return resource.WithComCleanup<Assistant, IAssistant>();
		}

	/// <summary>
		/// Wrapper interface for DocumentProperty which adds IDispose to the interface
		/// </summary>
		public static IDocumentProperty WithComCleanup(this DocumentProperty resource)
		{
			return resource.WithComCleanup<DocumentProperty, IDocumentProperty>();
		}

	/// <summary>
		/// Wrapper interface for DocumentProperties which adds IDispose to the interface
		/// </summary>
		public static IDocumentProperties WithComCleanup(this DocumentProperties resource)
		{
			return resource.WithComCleanup<DocumentProperties, IDocumentProperties>();
		}

	/// <summary>
		/// Wrapper interface for IFoundFiles which adds IDispose to the interface
		/// </summary>
		public static IIFoundFiles WithComCleanup(this IFoundFiles resource)
		{
			return resource.WithComCleanup<IFoundFiles, IIFoundFiles>();
		}

	/// <summary>
		/// Wrapper interface for IFind which adds IDispose to the interface
		/// </summary>
		public static IIFind WithComCleanup(this IFind resource)
		{
			return resource.WithComCleanup<IFind, IIFind>();
		}

	/// <summary>
		/// Wrapper interface for FoundFiles which adds IDispose to the interface
		/// </summary>
		public static IFoundFiles WithComCleanup(this FoundFiles resource)
		{
			return resource.WithComCleanup<FoundFiles, IFoundFiles>();
		}

	/// <summary>
		/// Wrapper interface for PropertyTest which adds IDispose to the interface
		/// </summary>
		public static IPropertyTest WithComCleanup(this PropertyTest resource)
		{
			return resource.WithComCleanup<PropertyTest, IPropertyTest>();
		}

	/// <summary>
		/// Wrapper interface for PropertyTests which adds IDispose to the interface
		/// </summary>
		public static IPropertyTests WithComCleanup(this PropertyTests resource)
		{
			return resource.WithComCleanup<PropertyTests, IPropertyTests>();
		}

	/// <summary>
		/// Wrapper interface for FileSearch which adds IDispose to the interface
		/// </summary>
		public static IFileSearch WithComCleanup(this FileSearch resource)
		{
			return resource.WithComCleanup<FileSearch, IFileSearch>();
		}

	/// <summary>
		/// Wrapper interface for COMAddIn which adds IDispose to the interface
		/// </summary>
		public static ICOMAddIn WithComCleanup(this COMAddIn resource)
		{
			return resource.WithComCleanup<COMAddIn, ICOMAddIn>();
		}

	/// <summary>
		/// Wrapper interface for COMAddIns which adds IDispose to the interface
		/// </summary>
		public static ICOMAddIns WithComCleanup(this COMAddIns resource)
		{
			return resource.WithComCleanup<COMAddIns, ICOMAddIns>();
		}

	/// <summary>
		/// Wrapper interface for LanguageSettings which adds IDispose to the interface
		/// </summary>
		public static ILanguageSettings WithComCleanup(this LanguageSettings resource)
		{
			return resource.WithComCleanup<LanguageSettings, ILanguageSettings>();
		}

	/// <summary>
		/// Wrapper interface for ICommandBarsEvents which adds IDispose to the interface
		/// </summary>
		public static IICommandBarsEvents WithComCleanup(this ICommandBarsEvents resource)
		{
			return resource.WithComCleanup<ICommandBarsEvents, IICommandBarsEvents>();
		}

	/// <summary>
		/// Wrapper interface for _CommandBarsEvents which adds IDispose to the interface
		/// </summary>
		public static I_CommandBarsEvents WithComCleanup(this _CommandBarsEvents resource)
		{
			return resource.WithComCleanup<_CommandBarsEvents, I_CommandBarsEvents>();
		}

	/// <summary>
		/// Wrapper interface for _CommandBarsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static I_CommandBarsEvents_Event WithComCleanup(this _CommandBarsEvents_Event resource)
		{
			return resource.WithComCleanup<_CommandBarsEvents_Event, I_CommandBarsEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for CommandBars which adds IDispose to the interface
		/// </summary>
		public static ICommandBars WithComCleanup(this CommandBars resource)
		{
			return resource.WithComCleanup<CommandBars, ICommandBars>();
		}

	/// <summary>
		/// Wrapper interface for ICommandBarComboBoxEvents which adds IDispose to the interface
		/// </summary>
		public static IICommandBarComboBoxEvents WithComCleanup(this ICommandBarComboBoxEvents resource)
		{
			return resource.WithComCleanup<ICommandBarComboBoxEvents, IICommandBarComboBoxEvents>();
		}

	/// <summary>
		/// Wrapper interface for _CommandBarComboBoxEvents which adds IDispose to the interface
		/// </summary>
		public static I_CommandBarComboBoxEvents WithComCleanup(this _CommandBarComboBoxEvents resource)
		{
			return resource.WithComCleanup<_CommandBarComboBoxEvents, I_CommandBarComboBoxEvents>();
		}

	/// <summary>
		/// Wrapper interface for _CommandBarComboBoxEvents_Event which adds IDispose to the interface
		/// </summary>
		public static I_CommandBarComboBoxEvents_Event WithComCleanup(this _CommandBarComboBoxEvents_Event resource)
		{
			return resource.WithComCleanup<_CommandBarComboBoxEvents_Event, I_CommandBarComboBoxEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for CommandBarComboBox which adds IDispose to the interface
		/// </summary>
		public static ICommandBarComboBox WithComCleanup(this CommandBarComboBox resource)
		{
			return resource.WithComCleanup<CommandBarComboBox, ICommandBarComboBox>();
		}

	/// <summary>
		/// Wrapper interface for ICommandBarButtonEvents which adds IDispose to the interface
		/// </summary>
		public static IICommandBarButtonEvents WithComCleanup(this ICommandBarButtonEvents resource)
		{
			return resource.WithComCleanup<ICommandBarButtonEvents, IICommandBarButtonEvents>();
		}

	/// <summary>
		/// Wrapper interface for _CommandBarButtonEvents which adds IDispose to the interface
		/// </summary>
		public static I_CommandBarButtonEvents WithComCleanup(this _CommandBarButtonEvents resource)
		{
			return resource.WithComCleanup<_CommandBarButtonEvents, I_CommandBarButtonEvents>();
		}

	/// <summary>
		/// Wrapper interface for _CommandBarButtonEvents_Event which adds IDispose to the interface
		/// </summary>
		public static I_CommandBarButtonEvents_Event WithComCleanup(this _CommandBarButtonEvents_Event resource)
		{
			return resource.WithComCleanup<_CommandBarButtonEvents_Event, I_CommandBarButtonEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for CommandBarButton which adds IDispose to the interface
		/// </summary>
		public static ICommandBarButton WithComCleanup(this CommandBarButton resource)
		{
			return resource.WithComCleanup<CommandBarButton, ICommandBarButton>();
		}

	/// <summary>
		/// Wrapper interface for WebPageFont which adds IDispose to the interface
		/// </summary>
		public static IWebPageFont WithComCleanup(this WebPageFont resource)
		{
			return resource.WithComCleanup<WebPageFont, IWebPageFont>();
		}

	/// <summary>
		/// Wrapper interface for WebPageFonts which adds IDispose to the interface
		/// </summary>
		public static IWebPageFonts WithComCleanup(this WebPageFonts resource)
		{
			return resource.WithComCleanup<WebPageFonts, IWebPageFonts>();
		}

	/// <summary>
		/// Wrapper interface for HTMLProjectItem which adds IDispose to the interface
		/// </summary>
		public static IHTMLProjectItem WithComCleanup(this HTMLProjectItem resource)
		{
			return resource.WithComCleanup<HTMLProjectItem, IHTMLProjectItem>();
		}

	/// <summary>
		/// Wrapper interface for HTMLProjectItems which adds IDispose to the interface
		/// </summary>
		public static IHTMLProjectItems WithComCleanup(this HTMLProjectItems resource)
		{
			return resource.WithComCleanup<HTMLProjectItems, IHTMLProjectItems>();
		}

	/// <summary>
		/// Wrapper interface for HTMLProject which adds IDispose to the interface
		/// </summary>
		public static IHTMLProject WithComCleanup(this HTMLProject resource)
		{
			return resource.WithComCleanup<HTMLProject, IHTMLProject>();
		}

	/// <summary>
		/// Wrapper interface for MsoDebugOptions which adds IDispose to the interface
		/// </summary>
		public static IMsoDebugOptions WithComCleanup(this MsoDebugOptions resource)
		{
			return resource.WithComCleanup<MsoDebugOptions, IMsoDebugOptions>();
		}

	/// <summary>
		/// Wrapper interface for FileDialogSelectedItems which adds IDispose to the interface
		/// </summary>
		public static IFileDialogSelectedItems WithComCleanup(this FileDialogSelectedItems resource)
		{
			return resource.WithComCleanup<FileDialogSelectedItems, IFileDialogSelectedItems>();
		}

	/// <summary>
		/// Wrapper interface for FileDialogFilter which adds IDispose to the interface
		/// </summary>
		public static IFileDialogFilter WithComCleanup(this FileDialogFilter resource)
		{
			return resource.WithComCleanup<FileDialogFilter, IFileDialogFilter>();
		}

	/// <summary>
		/// Wrapper interface for FileDialogFilters which adds IDispose to the interface
		/// </summary>
		public static IFileDialogFilters WithComCleanup(this FileDialogFilters resource)
		{
			return resource.WithComCleanup<FileDialogFilters, IFileDialogFilters>();
		}

	/// <summary>
		/// Wrapper interface for FileDialog which adds IDispose to the interface
		/// </summary>
		public static IFileDialog WithComCleanup(this FileDialog resource)
		{
			return resource.WithComCleanup<FileDialog, IFileDialog>();
		}

	/// <summary>
		/// Wrapper interface for SignatureSet which adds IDispose to the interface
		/// </summary>
		public static ISignatureSet WithComCleanup(this SignatureSet resource)
		{
			return resource.WithComCleanup<SignatureSet, ISignatureSet>();
		}

	/// <summary>
		/// Wrapper interface for Signature which adds IDispose to the interface
		/// </summary>
		public static ISignature WithComCleanup(this Signature resource)
		{
			return resource.WithComCleanup<Signature, ISignature>();
		}

	/// <summary>
		/// Wrapper interface for IMsoEnvelopeVB which adds IDispose to the interface
		/// </summary>
		public static IIMsoEnvelopeVB WithComCleanup(this IMsoEnvelopeVB resource)
		{
			return resource.WithComCleanup<IMsoEnvelopeVB, IIMsoEnvelopeVB>();
		}

	/// <summary>
		/// Wrapper interface for IMsoEnvelopeVBEvents which adds IDispose to the interface
		/// </summary>
		public static IIMsoEnvelopeVBEvents WithComCleanup(this IMsoEnvelopeVBEvents resource)
		{
			return resource.WithComCleanup<IMsoEnvelopeVBEvents, IIMsoEnvelopeVBEvents>();
		}

	/// <summary>
		/// Wrapper interface for IMsoEnvelopeVBEvents_Event which adds IDispose to the interface
		/// </summary>
		public static IIMsoEnvelopeVBEvents_Event WithComCleanup(this IMsoEnvelopeVBEvents_Event resource)
		{
			return resource.WithComCleanup<IMsoEnvelopeVBEvents_Event, IIMsoEnvelopeVBEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for MsoEnvelope which adds IDispose to the interface
		/// </summary>
		public static IMsoEnvelope WithComCleanup(this MsoEnvelope resource)
		{
			return resource.WithComCleanup<MsoEnvelope, IMsoEnvelope>();
		}

	/// <summary>
		/// Wrapper interface for FileTypes which adds IDispose to the interface
		/// </summary>
		public static IFileTypes WithComCleanup(this FileTypes resource)
		{
			return resource.WithComCleanup<FileTypes, IFileTypes>();
		}

	/// <summary>
		/// Wrapper interface for SearchFolders which adds IDispose to the interface
		/// </summary>
		public static ISearchFolders WithComCleanup(this SearchFolders resource)
		{
			return resource.WithComCleanup<SearchFolders, ISearchFolders>();
		}

	/// <summary>
		/// Wrapper interface for ScopeFolders which adds IDispose to the interface
		/// </summary>
		public static IScopeFolders WithComCleanup(this ScopeFolders resource)
		{
			return resource.WithComCleanup<ScopeFolders, IScopeFolders>();
		}

	/// <summary>
		/// Wrapper interface for ScopeFolder which adds IDispose to the interface
		/// </summary>
		public static IScopeFolder WithComCleanup(this ScopeFolder resource)
		{
			return resource.WithComCleanup<ScopeFolder, IScopeFolder>();
		}

	/// <summary>
		/// Wrapper interface for SearchScope which adds IDispose to the interface
		/// </summary>
		public static ISearchScope WithComCleanup(this SearchScope resource)
		{
			return resource.WithComCleanup<SearchScope, ISearchScope>();
		}

	/// <summary>
		/// Wrapper interface for SearchScopes which adds IDispose to the interface
		/// </summary>
		public static ISearchScopes WithComCleanup(this SearchScopes resource)
		{
			return resource.WithComCleanup<SearchScopes, ISearchScopes>();
		}

	/// <summary>
		/// Wrapper interface for IMsoDiagram which adds IDispose to the interface
		/// </summary>
		public static IIMsoDiagram WithComCleanup(this IMsoDiagram resource)
		{
			return resource.WithComCleanup<IMsoDiagram, IIMsoDiagram>();
		}

	/// <summary>
		/// Wrapper interface for DiagramNodes which adds IDispose to the interface
		/// </summary>
		public static IDiagramNodes WithComCleanup(this DiagramNodes resource)
		{
			return resource.WithComCleanup<DiagramNodes, IDiagramNodes>();
		}

	/// <summary>
		/// Wrapper interface for DiagramNodeChildren which adds IDispose to the interface
		/// </summary>
		public static IDiagramNodeChildren WithComCleanup(this DiagramNodeChildren resource)
		{
			return resource.WithComCleanup<DiagramNodeChildren, IDiagramNodeChildren>();
		}

	/// <summary>
		/// Wrapper interface for DiagramNode which adds IDispose to the interface
		/// </summary>
		public static IDiagramNode WithComCleanup(this DiagramNode resource)
		{
			return resource.WithComCleanup<DiagramNode, IDiagramNode>();
		}

	/// <summary>
		/// Wrapper interface for CanvasShapes which adds IDispose to the interface
		/// </summary>
		public static ICanvasShapes WithComCleanup(this CanvasShapes resource)
		{
			return resource.WithComCleanup<CanvasShapes, ICanvasShapes>();
		}

	/// <summary>
		/// Wrapper interface for OfficeDataSourceObject which adds IDispose to the interface
		/// </summary>
		public static IOfficeDataSourceObject WithComCleanup(this OfficeDataSourceObject resource)
		{
			return resource.WithComCleanup<OfficeDataSourceObject, IOfficeDataSourceObject>();
		}

	/// <summary>
		/// Wrapper interface for ODSOColumn which adds IDispose to the interface
		/// </summary>
		public static IODSOColumn WithComCleanup(this ODSOColumn resource)
		{
			return resource.WithComCleanup<ODSOColumn, IODSOColumn>();
		}

	/// <summary>
		/// Wrapper interface for ODSOColumns which adds IDispose to the interface
		/// </summary>
		public static IODSOColumns WithComCleanup(this ODSOColumns resource)
		{
			return resource.WithComCleanup<ODSOColumns, IODSOColumns>();
		}

	/// <summary>
		/// Wrapper interface for ODSOFilter which adds IDispose to the interface
		/// </summary>
		public static IODSOFilter WithComCleanup(this ODSOFilter resource)
		{
			return resource.WithComCleanup<ODSOFilter, IODSOFilter>();
		}

	/// <summary>
		/// Wrapper interface for ODSOFilters which adds IDispose to the interface
		/// </summary>
		public static IODSOFilters WithComCleanup(this ODSOFilters resource)
		{
			return resource.WithComCleanup<ODSOFilters, IODSOFilters>();
		}

	/// <summary>
		/// Wrapper interface for NewFile which adds IDispose to the interface
		/// </summary>
		public static INewFile WithComCleanup(this NewFile resource)
		{
			return resource.WithComCleanup<NewFile, INewFile>();
		}

	/// <summary>
		/// Wrapper interface for WebComponent which adds IDispose to the interface
		/// </summary>
		public static IWebComponent WithComCleanup(this WebComponent resource)
		{
			return resource.WithComCleanup<WebComponent, IWebComponent>();
		}

	/// <summary>
		/// Wrapper interface for WebComponentWindowExternal which adds IDispose to the interface
		/// </summary>
		public static IWebComponentWindowExternal WithComCleanup(this WebComponentWindowExternal resource)
		{
			return resource.WithComCleanup<WebComponentWindowExternal, IWebComponentWindowExternal>();
		}

	/// <summary>
		/// Wrapper interface for WebComponentFormat which adds IDispose to the interface
		/// </summary>
		public static IWebComponentFormat WithComCleanup(this WebComponentFormat resource)
		{
			return resource.WithComCleanup<WebComponentFormat, IWebComponentFormat>();
		}

	/// <summary>
		/// Wrapper interface for ILicWizExternal which adds IDispose to the interface
		/// </summary>
		public static IILicWizExternal WithComCleanup(this ILicWizExternal resource)
		{
			return resource.WithComCleanup<ILicWizExternal, IILicWizExternal>();
		}

	/// <summary>
		/// Wrapper interface for ILicValidator which adds IDispose to the interface
		/// </summary>
		public static IILicValidator WithComCleanup(this ILicValidator resource)
		{
			return resource.WithComCleanup<ILicValidator, IILicValidator>();
		}

	/// <summary>
		/// Wrapper interface for ILicAgent which adds IDispose to the interface
		/// </summary>
		public static IILicAgent WithComCleanup(this ILicAgent resource)
		{
			return resource.WithComCleanup<ILicAgent, IILicAgent>();
		}

	/// <summary>
		/// Wrapper interface for IMsoEServicesDialog which adds IDispose to the interface
		/// </summary>
		public static IIMsoEServicesDialog WithComCleanup(this IMsoEServicesDialog resource)
		{
			return resource.WithComCleanup<IMsoEServicesDialog, IIMsoEServicesDialog>();
		}

	/// <summary>
		/// Wrapper interface for WebComponentProperties which adds IDispose to the interface
		/// </summary>
		public static IWebComponentProperties WithComCleanup(this WebComponentProperties resource)
		{
			return resource.WithComCleanup<WebComponentProperties, IWebComponentProperties>();
		}

	/// <summary>
		/// Wrapper interface for SmartDocument which adds IDispose to the interface
		/// </summary>
		public static ISmartDocument WithComCleanup(this SmartDocument resource)
		{
			return resource.WithComCleanup<SmartDocument, ISmartDocument>();
		}

	/// <summary>
		/// Wrapper interface for SharedWorkspaceMember which adds IDispose to the interface
		/// </summary>
		public static ISharedWorkspaceMember WithComCleanup(this SharedWorkspaceMember resource)
		{
			return resource.WithComCleanup<SharedWorkspaceMember, ISharedWorkspaceMember>();
		}

	/// <summary>
		/// Wrapper interface for SharedWorkspaceMembers which adds IDispose to the interface
		/// </summary>
		public static ISharedWorkspaceMembers WithComCleanup(this SharedWorkspaceMembers resource)
		{
			return resource.WithComCleanup<SharedWorkspaceMembers, ISharedWorkspaceMembers>();
		}

	/// <summary>
		/// Wrapper interface for SharedWorkspaceTask which adds IDispose to the interface
		/// </summary>
		public static ISharedWorkspaceTask WithComCleanup(this SharedWorkspaceTask resource)
		{
			return resource.WithComCleanup<SharedWorkspaceTask, ISharedWorkspaceTask>();
		}

	/// <summary>
		/// Wrapper interface for SharedWorkspaceTasks which adds IDispose to the interface
		/// </summary>
		public static ISharedWorkspaceTasks WithComCleanup(this SharedWorkspaceTasks resource)
		{
			return resource.WithComCleanup<SharedWorkspaceTasks, ISharedWorkspaceTasks>();
		}

	/// <summary>
		/// Wrapper interface for SharedWorkspaceFile which adds IDispose to the interface
		/// </summary>
		public static ISharedWorkspaceFile WithComCleanup(this SharedWorkspaceFile resource)
		{
			return resource.WithComCleanup<SharedWorkspaceFile, ISharedWorkspaceFile>();
		}

	/// <summary>
		/// Wrapper interface for SharedWorkspaceFiles which adds IDispose to the interface
		/// </summary>
		public static ISharedWorkspaceFiles WithComCleanup(this SharedWorkspaceFiles resource)
		{
			return resource.WithComCleanup<SharedWorkspaceFiles, ISharedWorkspaceFiles>();
		}

	/// <summary>
		/// Wrapper interface for SharedWorkspaceFolder which adds IDispose to the interface
		/// </summary>
		public static ISharedWorkspaceFolder WithComCleanup(this SharedWorkspaceFolder resource)
		{
			return resource.WithComCleanup<SharedWorkspaceFolder, ISharedWorkspaceFolder>();
		}

	/// <summary>
		/// Wrapper interface for SharedWorkspaceFolders which adds IDispose to the interface
		/// </summary>
		public static ISharedWorkspaceFolders WithComCleanup(this SharedWorkspaceFolders resource)
		{
			return resource.WithComCleanup<SharedWorkspaceFolders, ISharedWorkspaceFolders>();
		}

	/// <summary>
		/// Wrapper interface for SharedWorkspaceLink which adds IDispose to the interface
		/// </summary>
		public static ISharedWorkspaceLink WithComCleanup(this SharedWorkspaceLink resource)
		{
			return resource.WithComCleanup<SharedWorkspaceLink, ISharedWorkspaceLink>();
		}

	/// <summary>
		/// Wrapper interface for SharedWorkspaceLinks which adds IDispose to the interface
		/// </summary>
		public static ISharedWorkspaceLinks WithComCleanup(this SharedWorkspaceLinks resource)
		{
			return resource.WithComCleanup<SharedWorkspaceLinks, ISharedWorkspaceLinks>();
		}

	/// <summary>
		/// Wrapper interface for SharedWorkspace which adds IDispose to the interface
		/// </summary>
		public static ISharedWorkspace WithComCleanup(this SharedWorkspace resource)
		{
			return resource.WithComCleanup<SharedWorkspace, ISharedWorkspace>();
		}

	/// <summary>
		/// Wrapper interface for Sync which adds IDispose to the interface
		/// </summary>
		public static ISync WithComCleanup(this Sync resource)
		{
			return resource.WithComCleanup<Sync, ISync>();
		}

	/// <summary>
		/// Wrapper interface for DocumentLibraryVersion which adds IDispose to the interface
		/// </summary>
		public static IDocumentLibraryVersion WithComCleanup(this DocumentLibraryVersion resource)
		{
			return resource.WithComCleanup<DocumentLibraryVersion, IDocumentLibraryVersion>();
		}

	/// <summary>
		/// Wrapper interface for DocumentLibraryVersions which adds IDispose to the interface
		/// </summary>
		public static IDocumentLibraryVersions WithComCleanup(this DocumentLibraryVersions resource)
		{
			return resource.WithComCleanup<DocumentLibraryVersions, IDocumentLibraryVersions>();
		}

	/// <summary>
		/// Wrapper interface for UserPermission which adds IDispose to the interface
		/// </summary>
		public static IUserPermission WithComCleanup(this UserPermission resource)
		{
			return resource.WithComCleanup<UserPermission, IUserPermission>();
		}

	/// <summary>
		/// Wrapper interface for Permission which adds IDispose to the interface
		/// </summary>
		public static IPermission WithComCleanup(this Permission resource)
		{
			return resource.WithComCleanup<Permission, IPermission>();
		}

	/// <summary>
		/// Wrapper interface for MsoDebugOptions_UTRunResult which adds IDispose to the interface
		/// </summary>
		public static IMsoDebugOptions_UTRunResult WithComCleanup(this MsoDebugOptions_UTRunResult resource)
		{
			return resource.WithComCleanup<MsoDebugOptions_UTRunResult, IMsoDebugOptions_UTRunResult>();
		}

	/// <summary>
		/// Wrapper interface for MsoDebugOptions_UT which adds IDispose to the interface
		/// </summary>
		public static IMsoDebugOptions_UT WithComCleanup(this MsoDebugOptions_UT resource)
		{
			return resource.WithComCleanup<MsoDebugOptions_UT, IMsoDebugOptions_UT>();
		}

	/// <summary>
		/// Wrapper interface for MsoDebugOptions_UTs which adds IDispose to the interface
		/// </summary>
		public static IMsoDebugOptions_UTs WithComCleanup(this MsoDebugOptions_UTs resource)
		{
			return resource.WithComCleanup<MsoDebugOptions_UTs, IMsoDebugOptions_UTs>();
		}

	/// <summary>
		/// Wrapper interface for MsoDebugOptions_UTManager which adds IDispose to the interface
		/// </summary>
		public static IMsoDebugOptions_UTManager WithComCleanup(this MsoDebugOptions_UTManager resource)
		{
			return resource.WithComCleanup<MsoDebugOptions_UTManager, IMsoDebugOptions_UTManager>();
		}

	/// <summary>
		/// Wrapper interface for MetaProperty which adds IDispose to the interface
		/// </summary>
		public static IMetaProperty WithComCleanup(this MetaProperty resource)
		{
			return resource.WithComCleanup<MetaProperty, IMetaProperty>();
		}

	/// <summary>
		/// Wrapper interface for MetaProperties which adds IDispose to the interface
		/// </summary>
		public static IMetaProperties WithComCleanup(this MetaProperties resource)
		{
			return resource.WithComCleanup<MetaProperties, IMetaProperties>();
		}

	/// <summary>
		/// Wrapper interface for PolicyItem which adds IDispose to the interface
		/// </summary>
		public static IPolicyItem WithComCleanup(this PolicyItem resource)
		{
			return resource.WithComCleanup<PolicyItem, IPolicyItem>();
		}

	/// <summary>
		/// Wrapper interface for ServerPolicy which adds IDispose to the interface
		/// </summary>
		public static IServerPolicy WithComCleanup(this ServerPolicy resource)
		{
			return resource.WithComCleanup<ServerPolicy, IServerPolicy>();
		}

	/// <summary>
		/// Wrapper interface for DocumentInspector which adds IDispose to the interface
		/// </summary>
		public static IDocumentInspector WithComCleanup(this DocumentInspector resource)
		{
			return resource.WithComCleanup<DocumentInspector, IDocumentInspector>();
		}

	/// <summary>
		/// Wrapper interface for DocumentInspectors which adds IDispose to the interface
		/// </summary>
		public static IDocumentInspectors WithComCleanup(this DocumentInspectors resource)
		{
			return resource.WithComCleanup<DocumentInspectors, IDocumentInspectors>();
		}

	/// <summary>
		/// Wrapper interface for WorkflowTask which adds IDispose to the interface
		/// </summary>
		public static IWorkflowTask WithComCleanup(this WorkflowTask resource)
		{
			return resource.WithComCleanup<WorkflowTask, IWorkflowTask>();
		}

	/// <summary>
		/// Wrapper interface for WorkflowTasks which adds IDispose to the interface
		/// </summary>
		public static IWorkflowTasks WithComCleanup(this WorkflowTasks resource)
		{
			return resource.WithComCleanup<WorkflowTasks, IWorkflowTasks>();
		}

	/// <summary>
		/// Wrapper interface for WorkflowTemplate which adds IDispose to the interface
		/// </summary>
		public static IWorkflowTemplate WithComCleanup(this WorkflowTemplate resource)
		{
			return resource.WithComCleanup<WorkflowTemplate, IWorkflowTemplate>();
		}

	/// <summary>
		/// Wrapper interface for WorkflowTemplates which adds IDispose to the interface
		/// </summary>
		public static IWorkflowTemplates WithComCleanup(this WorkflowTemplates resource)
		{
			return resource.WithComCleanup<WorkflowTemplates, IWorkflowTemplates>();
		}

	/// <summary>
		/// Wrapper interface for IDocumentInspector which adds IDispose to the interface
		/// </summary>
		public static IIDocumentInspector WithComCleanup(this IDocumentInspector resource)
		{
			return resource.WithComCleanup<IDocumentInspector, IIDocumentInspector>();
		}

	/// <summary>
		/// Wrapper interface for SignatureSetup which adds IDispose to the interface
		/// </summary>
		public static ISignatureSetup WithComCleanup(this SignatureSetup resource)
		{
			return resource.WithComCleanup<SignatureSetup, ISignatureSetup>();
		}

	/// <summary>
		/// Wrapper interface for SignatureInfo which adds IDispose to the interface
		/// </summary>
		public static ISignatureInfo WithComCleanup(this SignatureInfo resource)
		{
			return resource.WithComCleanup<SignatureInfo, ISignatureInfo>();
		}

	/// <summary>
		/// Wrapper interface for SignatureProvider which adds IDispose to the interface
		/// </summary>
		public static ISignatureProvider WithComCleanup(this SignatureProvider resource)
		{
			return resource.WithComCleanup<SignatureProvider, ISignatureProvider>();
		}

	/// <summary>
		/// Wrapper interface for CustomXMLPrefixMapping which adds IDispose to the interface
		/// </summary>
		public static ICustomXMLPrefixMapping WithComCleanup(this CustomXMLPrefixMapping resource)
		{
			return resource.WithComCleanup<CustomXMLPrefixMapping, ICustomXMLPrefixMapping>();
		}

	/// <summary>
		/// Wrapper interface for CustomXMLPrefixMappings which adds IDispose to the interface
		/// </summary>
		public static ICustomXMLPrefixMappings WithComCleanup(this CustomXMLPrefixMappings resource)
		{
			return resource.WithComCleanup<CustomXMLPrefixMappings, ICustomXMLPrefixMappings>();
		}

	/// <summary>
		/// Wrapper interface for CustomXMLSchema which adds IDispose to the interface
		/// </summary>
		public static ICustomXMLSchema WithComCleanup(this CustomXMLSchema resource)
		{
			return resource.WithComCleanup<CustomXMLSchema, ICustomXMLSchema>();
		}

	/// <summary>
		/// Wrapper interface for _CustomXMLSchemaCollection which adds IDispose to the interface
		/// </summary>
		public static I_CustomXMLSchemaCollection WithComCleanup(this _CustomXMLSchemaCollection resource)
		{
			return resource.WithComCleanup<_CustomXMLSchemaCollection, I_CustomXMLSchemaCollection>();
		}

	/// <summary>
		/// Wrapper interface for CustomXMLSchemaCollection which adds IDispose to the interface
		/// </summary>
		public static ICustomXMLSchemaCollection WithComCleanup(this CustomXMLSchemaCollection resource)
		{
			return resource.WithComCleanup<CustomXMLSchemaCollection, ICustomXMLSchemaCollection>();
		}

	/// <summary>
		/// Wrapper interface for CustomXMLNodes which adds IDispose to the interface
		/// </summary>
		public static ICustomXMLNodes WithComCleanup(this CustomXMLNodes resource)
		{
			return resource.WithComCleanup<CustomXMLNodes, ICustomXMLNodes>();
		}

	/// <summary>
		/// Wrapper interface for CustomXMLNode which adds IDispose to the interface
		/// </summary>
		public static ICustomXMLNode WithComCleanup(this CustomXMLNode resource)
		{
			return resource.WithComCleanup<CustomXMLNode, ICustomXMLNode>();
		}

	/// <summary>
		/// Wrapper interface for CustomXMLValidationError which adds IDispose to the interface
		/// </summary>
		public static ICustomXMLValidationError WithComCleanup(this CustomXMLValidationError resource)
		{
			return resource.WithComCleanup<CustomXMLValidationError, ICustomXMLValidationError>();
		}

	/// <summary>
		/// Wrapper interface for CustomXMLValidationErrors which adds IDispose to the interface
		/// </summary>
		public static ICustomXMLValidationErrors WithComCleanup(this CustomXMLValidationErrors resource)
		{
			return resource.WithComCleanup<CustomXMLValidationErrors, ICustomXMLValidationErrors>();
		}

	/// <summary>
		/// Wrapper interface for _CustomXMLPart which adds IDispose to the interface
		/// </summary>
		public static I_CustomXMLPart WithComCleanup(this _CustomXMLPart resource)
		{
			return resource.WithComCleanup<_CustomXMLPart, I_CustomXMLPart>();
		}

	/// <summary>
		/// Wrapper interface for ICustomXMLPartEvents which adds IDispose to the interface
		/// </summary>
		public static IICustomXMLPartEvents WithComCleanup(this ICustomXMLPartEvents resource)
		{
			return resource.WithComCleanup<ICustomXMLPartEvents, IICustomXMLPartEvents>();
		}

	/// <summary>
		/// Wrapper interface for _CustomXMLPartEvents which adds IDispose to the interface
		/// </summary>
		public static I_CustomXMLPartEvents WithComCleanup(this _CustomXMLPartEvents resource)
		{
			return resource.WithComCleanup<_CustomXMLPartEvents, I_CustomXMLPartEvents>();
		}

	/// <summary>
		/// Wrapper interface for _CustomXMLPartEvents_Event which adds IDispose to the interface
		/// </summary>
		public static I_CustomXMLPartEvents_Event WithComCleanup(this _CustomXMLPartEvents_Event resource)
		{
			return resource.WithComCleanup<_CustomXMLPartEvents_Event, I_CustomXMLPartEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for CustomXMLPart which adds IDispose to the interface
		/// </summary>
		public static ICustomXMLPart WithComCleanup(this CustomXMLPart resource)
		{
			return resource.WithComCleanup<CustomXMLPart, ICustomXMLPart>();
		}

	/// <summary>
		/// Wrapper interface for _CustomXMLParts which adds IDispose to the interface
		/// </summary>
		public static I_CustomXMLParts WithComCleanup(this _CustomXMLParts resource)
		{
			return resource.WithComCleanup<_CustomXMLParts, I_CustomXMLParts>();
		}

	/// <summary>
		/// Wrapper interface for ICustomXMLPartsEvents which adds IDispose to the interface
		/// </summary>
		public static IICustomXMLPartsEvents WithComCleanup(this ICustomXMLPartsEvents resource)
		{
			return resource.WithComCleanup<ICustomXMLPartsEvents, IICustomXMLPartsEvents>();
		}

	/// <summary>
		/// Wrapper interface for _CustomXMLPartsEvents which adds IDispose to the interface
		/// </summary>
		public static I_CustomXMLPartsEvents WithComCleanup(this _CustomXMLPartsEvents resource)
		{
			return resource.WithComCleanup<_CustomXMLPartsEvents, I_CustomXMLPartsEvents>();
		}

	/// <summary>
		/// Wrapper interface for _CustomXMLPartsEvents_Event which adds IDispose to the interface
		/// </summary>
		public static I_CustomXMLPartsEvents_Event WithComCleanup(this _CustomXMLPartsEvents_Event resource)
		{
			return resource.WithComCleanup<_CustomXMLPartsEvents_Event, I_CustomXMLPartsEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for CustomXMLParts which adds IDispose to the interface
		/// </summary>
		public static ICustomXMLParts WithComCleanup(this CustomXMLParts resource)
		{
			return resource.WithComCleanup<CustomXMLParts, ICustomXMLParts>();
		}

	/// <summary>
		/// Wrapper interface for GradientStop which adds IDispose to the interface
		/// </summary>
		public static IGradientStop WithComCleanup(this GradientStop resource)
		{
			return resource.WithComCleanup<GradientStop, IGradientStop>();
		}

	/// <summary>
		/// Wrapper interface for GradientStops which adds IDispose to the interface
		/// </summary>
		public static IGradientStops WithComCleanup(this GradientStops resource)
		{
			return resource.WithComCleanup<GradientStops, IGradientStops>();
		}

	/// <summary>
		/// Wrapper interface for SoftEdgeFormat which adds IDispose to the interface
		/// </summary>
		public static ISoftEdgeFormat WithComCleanup(this SoftEdgeFormat resource)
		{
			return resource.WithComCleanup<SoftEdgeFormat, ISoftEdgeFormat>();
		}

	/// <summary>
		/// Wrapper interface for GlowFormat which adds IDispose to the interface
		/// </summary>
		public static IGlowFormat WithComCleanup(this GlowFormat resource)
		{
			return resource.WithComCleanup<GlowFormat, IGlowFormat>();
		}

	/// <summary>
		/// Wrapper interface for ReflectionFormat which adds IDispose to the interface
		/// </summary>
		public static IReflectionFormat WithComCleanup(this ReflectionFormat resource)
		{
			return resource.WithComCleanup<ReflectionFormat, IReflectionFormat>();
		}

	/// <summary>
		/// Wrapper interface for ParagraphFormat2 which adds IDispose to the interface
		/// </summary>
		public static IParagraphFormat2 WithComCleanup(this ParagraphFormat2 resource)
		{
			return resource.WithComCleanup<ParagraphFormat2, IParagraphFormat2>();
		}

	/// <summary>
		/// Wrapper interface for Font2 which adds IDispose to the interface
		/// </summary>
		public static IFont2 WithComCleanup(this Font2 resource)
		{
			return resource.WithComCleanup<Font2, IFont2>();
		}

	/// <summary>
		/// Wrapper interface for TextColumn2 which adds IDispose to the interface
		/// </summary>
		public static ITextColumn2 WithComCleanup(this TextColumn2 resource)
		{
			return resource.WithComCleanup<TextColumn2, ITextColumn2>();
		}

	/// <summary>
		/// Wrapper interface for TextRange2 which adds IDispose to the interface
		/// </summary>
		public static ITextRange2 WithComCleanup(this TextRange2 resource)
		{
			return resource.WithComCleanup<TextRange2, ITextRange2>();
		}

	/// <summary>
		/// Wrapper interface for TextFrame2 which adds IDispose to the interface
		/// </summary>
		public static ITextFrame2 WithComCleanup(this TextFrame2 resource)
		{
			return resource.WithComCleanup<TextFrame2, ITextFrame2>();
		}

	/// <summary>
		/// Wrapper interface for ThemeColor which adds IDispose to the interface
		/// </summary>
		public static IThemeColor WithComCleanup(this ThemeColor resource)
		{
			return resource.WithComCleanup<ThemeColor, IThemeColor>();
		}

	/// <summary>
		/// Wrapper interface for ThemeColorScheme which adds IDispose to the interface
		/// </summary>
		public static IThemeColorScheme WithComCleanup(this ThemeColorScheme resource)
		{
			return resource.WithComCleanup<ThemeColorScheme, IThemeColorScheme>();
		}

	/// <summary>
		/// Wrapper interface for ThemeFont which adds IDispose to the interface
		/// </summary>
		public static IThemeFont WithComCleanup(this ThemeFont resource)
		{
			return resource.WithComCleanup<ThemeFont, IThemeFont>();
		}

	/// <summary>
		/// Wrapper interface for ThemeFonts which adds IDispose to the interface
		/// </summary>
		public static IThemeFonts WithComCleanup(this ThemeFonts resource)
		{
			return resource.WithComCleanup<ThemeFonts, IThemeFonts>();
		}

	/// <summary>
		/// Wrapper interface for ThemeFontScheme which adds IDispose to the interface
		/// </summary>
		public static IThemeFontScheme WithComCleanup(this ThemeFontScheme resource)
		{
			return resource.WithComCleanup<ThemeFontScheme, IThemeFontScheme>();
		}

	/// <summary>
		/// Wrapper interface for ThemeEffectScheme which adds IDispose to the interface
		/// </summary>
		public static IThemeEffectScheme WithComCleanup(this ThemeEffectScheme resource)
		{
			return resource.WithComCleanup<ThemeEffectScheme, IThemeEffectScheme>();
		}

	/// <summary>
		/// Wrapper interface for OfficeTheme which adds IDispose to the interface
		/// </summary>
		public static IOfficeTheme WithComCleanup(this OfficeTheme resource)
		{
			return resource.WithComCleanup<OfficeTheme, IOfficeTheme>();
		}

	/// <summary>
		/// Wrapper interface for _CustomTaskPane which adds IDispose to the interface
		/// </summary>
		public static I_CustomTaskPane WithComCleanup(this _CustomTaskPane resource)
		{
			return resource.WithComCleanup<_CustomTaskPane, I_CustomTaskPane>();
		}

	/// <summary>
		/// Wrapper interface for CustomTaskPaneEvents which adds IDispose to the interface
		/// </summary>
		public static ICustomTaskPaneEvents WithComCleanup(this CustomTaskPaneEvents resource)
		{
			return resource.WithComCleanup<CustomTaskPaneEvents, ICustomTaskPaneEvents>();
		}

	/// <summary>
		/// Wrapper interface for _CustomTaskPaneEvents which adds IDispose to the interface
		/// </summary>
		public static I_CustomTaskPaneEvents WithComCleanup(this _CustomTaskPaneEvents resource)
		{
			return resource.WithComCleanup<_CustomTaskPaneEvents, I_CustomTaskPaneEvents>();
		}

	/// <summary>
		/// Wrapper interface for _CustomTaskPaneEvents_Event which adds IDispose to the interface
		/// </summary>
		public static I_CustomTaskPaneEvents_Event WithComCleanup(this _CustomTaskPaneEvents_Event resource)
		{
			return resource.WithComCleanup<_CustomTaskPaneEvents_Event, I_CustomTaskPaneEvents_Event>();
		}

	/// <summary>
		/// Wrapper interface for CustomTaskPane which adds IDispose to the interface
		/// </summary>
		public static ICustomTaskPane WithComCleanup(this CustomTaskPane resource)
		{
			return resource.WithComCleanup<CustomTaskPane, ICustomTaskPane>();
		}

	/// <summary>
		/// Wrapper interface for ICTPFactory which adds IDispose to the interface
		/// </summary>
		public static IICTPFactory WithComCleanup(this ICTPFactory resource)
		{
			return resource.WithComCleanup<ICTPFactory, IICTPFactory>();
		}

	/// <summary>
		/// Wrapper interface for ICustomTaskPaneConsumer which adds IDispose to the interface
		/// </summary>
		public static IICustomTaskPaneConsumer WithComCleanup(this ICustomTaskPaneConsumer resource)
		{
			return resource.WithComCleanup<ICustomTaskPaneConsumer, IICustomTaskPaneConsumer>();
		}

	/// <summary>
		/// Wrapper interface for IRibbonUI which adds IDispose to the interface
		/// </summary>
		public static IIRibbonUI WithComCleanup(this IRibbonUI resource)
		{
			return resource.WithComCleanup<IRibbonUI, IIRibbonUI>();
		}

	/// <summary>
		/// Wrapper interface for IRibbonControl which adds IDispose to the interface
		/// </summary>
		public static IIRibbonControl WithComCleanup(this IRibbonControl resource)
		{
			return resource.WithComCleanup<IRibbonControl, IIRibbonControl>();
		}

	/// <summary>
		/// Wrapper interface for IRibbonExtensibility which adds IDispose to the interface
		/// </summary>
		public static IIRibbonExtensibility WithComCleanup(this IRibbonExtensibility resource)
		{
			return resource.WithComCleanup<IRibbonExtensibility, IIRibbonExtensibility>();
		}

	/// <summary>
		/// Wrapper interface for IAssistance which adds IDispose to the interface
		/// </summary>
		public static IIAssistance WithComCleanup(this IAssistance resource)
		{
			return resource.WithComCleanup<IAssistance, IIAssistance>();
		}

	/// <summary>
		/// Wrapper interface for IMsoChartData which adds IDispose to the interface
		/// </summary>
		public static IIMsoChartData WithComCleanup(this IMsoChartData resource)
		{
			return resource.WithComCleanup<IMsoChartData, IIMsoChartData>();
		}

	/// <summary>
		/// Wrapper interface for IMsoChart which adds IDispose to the interface
		/// </summary>
		public static IIMsoChart WithComCleanup(this IMsoChart resource)
		{
			return resource.WithComCleanup<IMsoChart, IIMsoChart>();
		}

	/// <summary>
		/// Wrapper interface for IMsoCorners which adds IDispose to the interface
		/// </summary>
		public static IIMsoCorners WithComCleanup(this IMsoCorners resource)
		{
			return resource.WithComCleanup<IMsoCorners, IIMsoCorners>();
		}

	/// <summary>
		/// Wrapper interface for IMsoLegend which adds IDispose to the interface
		/// </summary>
		public static IIMsoLegend WithComCleanup(this IMsoLegend resource)
		{
			return resource.WithComCleanup<IMsoLegend, IIMsoLegend>();
		}

	/// <summary>
		/// Wrapper interface for IMsoBorder which adds IDispose to the interface
		/// </summary>
		public static IIMsoBorder WithComCleanup(this IMsoBorder resource)
		{
			return resource.WithComCleanup<IMsoBorder, IIMsoBorder>();
		}

	/// <summary>
		/// Wrapper interface for IMsoWalls which adds IDispose to the interface
		/// </summary>
		public static IIMsoWalls WithComCleanup(this IMsoWalls resource)
		{
			return resource.WithComCleanup<IMsoWalls, IIMsoWalls>();
		}

	/// <summary>
		/// Wrapper interface for IMsoFloor which adds IDispose to the interface
		/// </summary>
		public static IIMsoFloor WithComCleanup(this IMsoFloor resource)
		{
			return resource.WithComCleanup<IMsoFloor, IIMsoFloor>();
		}

	/// <summary>
		/// Wrapper interface for IMsoPlotArea which adds IDispose to the interface
		/// </summary>
		public static IIMsoPlotArea WithComCleanup(this IMsoPlotArea resource)
		{
			return resource.WithComCleanup<IMsoPlotArea, IIMsoPlotArea>();
		}

	/// <summary>
		/// Wrapper interface for IMsoChartArea which adds IDispose to the interface
		/// </summary>
		public static IIMsoChartArea WithComCleanup(this IMsoChartArea resource)
		{
			return resource.WithComCleanup<IMsoChartArea, IIMsoChartArea>();
		}

	/// <summary>
		/// Wrapper interface for IMsoSeriesLines which adds IDispose to the interface
		/// </summary>
		public static IIMsoSeriesLines WithComCleanup(this IMsoSeriesLines resource)
		{
			return resource.WithComCleanup<IMsoSeriesLines, IIMsoSeriesLines>();
		}

	/// <summary>
		/// Wrapper interface for IMsoLeaderLines which adds IDispose to the interface
		/// </summary>
		public static IIMsoLeaderLines WithComCleanup(this IMsoLeaderLines resource)
		{
			return resource.WithComCleanup<IMsoLeaderLines, IIMsoLeaderLines>();
		}

	/// <summary>
		/// Wrapper interface for GridLines which adds IDispose to the interface
		/// </summary>
		public static IGridLines WithComCleanup(this GridLines resource)
		{
			return resource.WithComCleanup<GridLines, IGridLines>();
		}

	/// <summary>
		/// Wrapper interface for IMsoUpBars which adds IDispose to the interface
		/// </summary>
		public static IIMsoUpBars WithComCleanup(this IMsoUpBars resource)
		{
			return resource.WithComCleanup<IMsoUpBars, IIMsoUpBars>();
		}

	/// <summary>
		/// Wrapper interface for IMsoDownBars which adds IDispose to the interface
		/// </summary>
		public static IIMsoDownBars WithComCleanup(this IMsoDownBars resource)
		{
			return resource.WithComCleanup<IMsoDownBars, IIMsoDownBars>();
		}

	/// <summary>
		/// Wrapper interface for IMsoInterior which adds IDispose to the interface
		/// </summary>
		public static IIMsoInterior WithComCleanup(this IMsoInterior resource)
		{
			return resource.WithComCleanup<IMsoInterior, IIMsoInterior>();
		}

	/// <summary>
		/// Wrapper interface for ChartFillFormat which adds IDispose to the interface
		/// </summary>
		public static IChartFillFormat WithComCleanup(this ChartFillFormat resource)
		{
			return resource.WithComCleanup<ChartFillFormat, IChartFillFormat>();
		}

	/// <summary>
		/// Wrapper interface for LegendEntries which adds IDispose to the interface
		/// </summary>
		public static ILegendEntries WithComCleanup(this LegendEntries resource)
		{
			return resource.WithComCleanup<LegendEntries, ILegendEntries>();
		}

	/// <summary>
		/// Wrapper interface for ChartFont which adds IDispose to the interface
		/// </summary>
		public static IChartFont WithComCleanup(this ChartFont resource)
		{
			return resource.WithComCleanup<ChartFont, IChartFont>();
		}

	/// <summary>
		/// Wrapper interface for ChartColorFormat which adds IDispose to the interface
		/// </summary>
		public static IChartColorFormat WithComCleanup(this ChartColorFormat resource)
		{
			return resource.WithComCleanup<ChartColorFormat, IChartColorFormat>();
		}

	/// <summary>
		/// Wrapper interface for LegendEntry which adds IDispose to the interface
		/// </summary>
		public static ILegendEntry WithComCleanup(this LegendEntry resource)
		{
			return resource.WithComCleanup<LegendEntry, ILegendEntry>();
		}

	/// <summary>
		/// Wrapper interface for IMsoLegendKey which adds IDispose to the interface
		/// </summary>
		public static IIMsoLegendKey WithComCleanup(this IMsoLegendKey resource)
		{
			return resource.WithComCleanup<IMsoLegendKey, IIMsoLegendKey>();
		}

	/// <summary>
		/// Wrapper interface for SeriesCollection which adds IDispose to the interface
		/// </summary>
		public static ISeriesCollection WithComCleanup(this SeriesCollection resource)
		{
			return resource.WithComCleanup<SeriesCollection, ISeriesCollection>();
		}

	/// <summary>
		/// Wrapper interface for IMsoSeries which adds IDispose to the interface
		/// </summary>
		public static IIMsoSeries WithComCleanup(this IMsoSeries resource)
		{
			return resource.WithComCleanup<IMsoSeries, IIMsoSeries>();
		}

	/// <summary>
		/// Wrapper interface for IMsoErrorBars which adds IDispose to the interface
		/// </summary>
		public static IIMsoErrorBars WithComCleanup(this IMsoErrorBars resource)
		{
			return resource.WithComCleanup<IMsoErrorBars, IIMsoErrorBars>();
		}

	/// <summary>
		/// Wrapper interface for IMsoTrendline which adds IDispose to the interface
		/// </summary>
		public static IIMsoTrendline WithComCleanup(this IMsoTrendline resource)
		{
			return resource.WithComCleanup<IMsoTrendline, IIMsoTrendline>();
		}

	/// <summary>
		/// Wrapper interface for Trendlines which adds IDispose to the interface
		/// </summary>
		public static ITrendlines WithComCleanup(this Trendlines resource)
		{
			return resource.WithComCleanup<Trendlines, ITrendlines>();
		}

	/// <summary>
		/// Wrapper interface for IMsoDataLabels which adds IDispose to the interface
		/// </summary>
		public static IIMsoDataLabels WithComCleanup(this IMsoDataLabels resource)
		{
			return resource.WithComCleanup<IMsoDataLabels, IIMsoDataLabels>();
		}

	/// <summary>
		/// Wrapper interface for IMsoDataLabel which adds IDispose to the interface
		/// </summary>
		public static IIMsoDataLabel WithComCleanup(this IMsoDataLabel resource)
		{
			return resource.WithComCleanup<IMsoDataLabel, IIMsoDataLabel>();
		}

	/// <summary>
		/// Wrapper interface for Points which adds IDispose to the interface
		/// </summary>
		public static IPoints WithComCleanup(this Points resource)
		{
			return resource.WithComCleanup<Points, IPoints>();
		}

	/// <summary>
		/// Wrapper interface for ChartPoint which adds IDispose to the interface
		/// </summary>
		public static IChartPoint WithComCleanup(this ChartPoint resource)
		{
			return resource.WithComCleanup<ChartPoint, IChartPoint>();
		}

	/// <summary>
		/// Wrapper interface for Axes which adds IDispose to the interface
		/// </summary>
		public static IAxes WithComCleanup(this Axes resource)
		{
			return resource.WithComCleanup<Axes, IAxes>();
		}

	/// <summary>
		/// Wrapper interface for IMsoAxis which adds IDispose to the interface
		/// </summary>
		public static IIMsoAxis WithComCleanup(this IMsoAxis resource)
		{
			return resource.WithComCleanup<IMsoAxis, IIMsoAxis>();
		}

	/// <summary>
		/// Wrapper interface for IMsoDataTable which adds IDispose to the interface
		/// </summary>
		public static IIMsoDataTable WithComCleanup(this IMsoDataTable resource)
		{
			return resource.WithComCleanup<IMsoDataTable, IIMsoDataTable>();
		}

	/// <summary>
		/// Wrapper interface for IMsoChartTitle which adds IDispose to the interface
		/// </summary>
		public static IIMsoChartTitle WithComCleanup(this IMsoChartTitle resource)
		{
			return resource.WithComCleanup<IMsoChartTitle, IIMsoChartTitle>();
		}

	/// <summary>
		/// Wrapper interface for IMsoAxisTitle which adds IDispose to the interface
		/// </summary>
		public static IIMsoAxisTitle WithComCleanup(this IMsoAxisTitle resource)
		{
			return resource.WithComCleanup<IMsoAxisTitle, IIMsoAxisTitle>();
		}

	/// <summary>
		/// Wrapper interface for IMsoDisplayUnitLabel which adds IDispose to the interface
		/// </summary>
		public static IIMsoDisplayUnitLabel WithComCleanup(this IMsoDisplayUnitLabel resource)
		{
			return resource.WithComCleanup<IMsoDisplayUnitLabel, IIMsoDisplayUnitLabel>();
		}

	/// <summary>
		/// Wrapper interface for IMsoTickLabels which adds IDispose to the interface
		/// </summary>
		public static IIMsoTickLabels WithComCleanup(this IMsoTickLabels resource)
		{
			return resource.WithComCleanup<IMsoTickLabels, IIMsoTickLabels>();
		}

	/// <summary>
		/// Wrapper interface for IMsoHyperlinks which adds IDispose to the interface
		/// </summary>
		public static IIMsoHyperlinks WithComCleanup(this IMsoHyperlinks resource)
		{
			return resource.WithComCleanup<IMsoHyperlinks, IIMsoHyperlinks>();
		}

	/// <summary>
		/// Wrapper interface for IMsoDropLines which adds IDispose to the interface
		/// </summary>
		public static IIMsoDropLines WithComCleanup(this IMsoDropLines resource)
		{
			return resource.WithComCleanup<IMsoDropLines, IIMsoDropLines>();
		}

	/// <summary>
		/// Wrapper interface for IMsoHiLoLines which adds IDispose to the interface
		/// </summary>
		public static IIMsoHiLoLines WithComCleanup(this IMsoHiLoLines resource)
		{
			return resource.WithComCleanup<IMsoHiLoLines, IIMsoHiLoLines>();
		}

	/// <summary>
		/// Wrapper interface for IMsoChartGroup which adds IDispose to the interface
		/// </summary>
		public static IIMsoChartGroup WithComCleanup(this IMsoChartGroup resource)
		{
			return resource.WithComCleanup<IMsoChartGroup, IIMsoChartGroup>();
		}

	/// <summary>
		/// Wrapper interface for ChartGroups which adds IDispose to the interface
		/// </summary>
		public static IChartGroups WithComCleanup(this ChartGroups resource)
		{
			return resource.WithComCleanup<ChartGroups, IChartGroups>();
		}

	/// <summary>
		/// Wrapper interface for IMsoCharacters which adds IDispose to the interface
		/// </summary>
		public static IIMsoCharacters WithComCleanup(this IMsoCharacters resource)
		{
			return resource.WithComCleanup<IMsoCharacters, IIMsoCharacters>();
		}

	/// <summary>
		/// Wrapper interface for IMsoChartFormat which adds IDispose to the interface
		/// </summary>
		public static IIMsoChartFormat WithComCleanup(this IMsoChartFormat resource)
		{
			return resource.WithComCleanup<IMsoChartFormat, IIMsoChartFormat>();
		}

	/// <summary>
		/// Wrapper interface for BulletFormat2 which adds IDispose to the interface
		/// </summary>
		public static IBulletFormat2 WithComCleanup(this BulletFormat2 resource)
		{
			return resource.WithComCleanup<BulletFormat2, IBulletFormat2>();
		}

	/// <summary>
		/// Wrapper interface for TabStops2 which adds IDispose to the interface
		/// </summary>
		public static ITabStops2 WithComCleanup(this TabStops2 resource)
		{
			return resource.WithComCleanup<TabStops2, ITabStops2>();
		}

	/// <summary>
		/// Wrapper interface for TabStop2 which adds IDispose to the interface
		/// </summary>
		public static ITabStop2 WithComCleanup(this TabStop2 resource)
		{
			return resource.WithComCleanup<TabStop2, ITabStop2>();
		}

	/// <summary>
		/// Wrapper interface for Ruler2 which adds IDispose to the interface
		/// </summary>
		public static IRuler2 WithComCleanup(this Ruler2 resource)
		{
			return resource.WithComCleanup<Ruler2, IRuler2>();
		}

	/// <summary>
		/// Wrapper interface for RulerLevels2 which adds IDispose to the interface
		/// </summary>
		public static IRulerLevels2 WithComCleanup(this RulerLevels2 resource)
		{
			return resource.WithComCleanup<RulerLevels2, IRulerLevels2>();
		}

	/// <summary>
		/// Wrapper interface for RulerLevel2 which adds IDispose to the interface
		/// </summary>
		public static IRulerLevel2 WithComCleanup(this RulerLevel2 resource)
		{
			return resource.WithComCleanup<RulerLevel2, IRulerLevel2>();
		}

	/// <summary>
		/// Wrapper interface for EncryptionProvider which adds IDispose to the interface
		/// </summary>
		public static IEncryptionProvider WithComCleanup(this EncryptionProvider resource)
		{
			return resource.WithComCleanup<EncryptionProvider, IEncryptionProvider>();
		}

	/// <summary>
		/// Wrapper interface for IBlogExtensibility which adds IDispose to the interface
		/// </summary>
		public static IIBlogExtensibility WithComCleanup(this IBlogExtensibility resource)
		{
			return resource.WithComCleanup<IBlogExtensibility, IIBlogExtensibility>();
		}

	/// <summary>
		/// Wrapper interface for IBlogPictureExtensibility which adds IDispose to the interface
		/// </summary>
		public static IIBlogPictureExtensibility WithComCleanup(this IBlogPictureExtensibility resource)
		{
			return resource.WithComCleanup<IBlogPictureExtensibility, IIBlogPictureExtensibility>();
		}

	/// <summary>
		/// Wrapper interface for IConverterPreferences which adds IDispose to the interface
		/// </summary>
		public static IIConverterPreferences WithComCleanup(this IConverterPreferences resource)
		{
			return resource.WithComCleanup<IConverterPreferences, IIConverterPreferences>();
		}

	/// <summary>
		/// Wrapper interface for IConverterApplicationPreferences which adds IDispose to the interface
		/// </summary>
		public static IIConverterApplicationPreferences WithComCleanup(this IConverterApplicationPreferences resource)
		{
			return resource.WithComCleanup<IConverterApplicationPreferences, IIConverterApplicationPreferences>();
		}

	/// <summary>
		/// Wrapper interface for IConverterUICallback which adds IDispose to the interface
		/// </summary>
		public static IIConverterUICallback WithComCleanup(this IConverterUICallback resource)
		{
			return resource.WithComCleanup<IConverterUICallback, IIConverterUICallback>();
		}

	/// <summary>
		/// Wrapper interface for IConverter which adds IDispose to the interface
		/// </summary>
		public static IIConverter WithComCleanup(this IConverter resource)
		{
			return resource.WithComCleanup<IConverter, IIConverter>();
		}

	/// <summary>
		/// Wrapper interface for SmartArt which adds IDispose to the interface
		/// </summary>
		public static ISmartArt WithComCleanup(this SmartArt resource)
		{
			return resource.WithComCleanup<SmartArt, ISmartArt>();
		}

	/// <summary>
		/// Wrapper interface for SmartArtNodes which adds IDispose to the interface
		/// </summary>
		public static ISmartArtNodes WithComCleanup(this SmartArtNodes resource)
		{
			return resource.WithComCleanup<SmartArtNodes, ISmartArtNodes>();
		}

	/// <summary>
		/// Wrapper interface for SmartArtNode which adds IDispose to the interface
		/// </summary>
		public static ISmartArtNode WithComCleanup(this SmartArtNode resource)
		{
			return resource.WithComCleanup<SmartArtNode, ISmartArtNode>();
		}

	/// <summary>
		/// Wrapper interface for SmartArtLayouts which adds IDispose to the interface
		/// </summary>
		public static ISmartArtLayouts WithComCleanup(this SmartArtLayouts resource)
		{
			return resource.WithComCleanup<SmartArtLayouts, ISmartArtLayouts>();
		}

	/// <summary>
		/// Wrapper interface for SmartArtLayout which adds IDispose to the interface
		/// </summary>
		public static ISmartArtLayout WithComCleanup(this SmartArtLayout resource)
		{
			return resource.WithComCleanup<SmartArtLayout, ISmartArtLayout>();
		}

	/// <summary>
		/// Wrapper interface for SmartArtQuickStyles which adds IDispose to the interface
		/// </summary>
		public static ISmartArtQuickStyles WithComCleanup(this SmartArtQuickStyles resource)
		{
			return resource.WithComCleanup<SmartArtQuickStyles, ISmartArtQuickStyles>();
		}

	/// <summary>
		/// Wrapper interface for SmartArtQuickStyle which adds IDispose to the interface
		/// </summary>
		public static ISmartArtQuickStyle WithComCleanup(this SmartArtQuickStyle resource)
		{
			return resource.WithComCleanup<SmartArtQuickStyle, ISmartArtQuickStyle>();
		}

	/// <summary>
		/// Wrapper interface for SmartArtColors which adds IDispose to the interface
		/// </summary>
		public static ISmartArtColors WithComCleanup(this SmartArtColors resource)
		{
			return resource.WithComCleanup<SmartArtColors, ISmartArtColors>();
		}

	/// <summary>
		/// Wrapper interface for SmartArtColor which adds IDispose to the interface
		/// </summary>
		public static ISmartArtColor WithComCleanup(this SmartArtColor resource)
		{
			return resource.WithComCleanup<SmartArtColor, ISmartArtColor>();
		}

	/// <summary>
		/// Wrapper interface for PickerField which adds IDispose to the interface
		/// </summary>
		public static IPickerField WithComCleanup(this PickerField resource)
		{
			return resource.WithComCleanup<PickerField, IPickerField>();
		}

	/// <summary>
		/// Wrapper interface for PickerFields which adds IDispose to the interface
		/// </summary>
		public static IPickerFields WithComCleanup(this PickerFields resource)
		{
			return resource.WithComCleanup<PickerFields, IPickerFields>();
		}

	/// <summary>
		/// Wrapper interface for PickerProperty which adds IDispose to the interface
		/// </summary>
		public static IPickerProperty WithComCleanup(this PickerProperty resource)
		{
			return resource.WithComCleanup<PickerProperty, IPickerProperty>();
		}

	/// <summary>
		/// Wrapper interface for PickerProperties which adds IDispose to the interface
		/// </summary>
		public static IPickerProperties WithComCleanup(this PickerProperties resource)
		{
			return resource.WithComCleanup<PickerProperties, IPickerProperties>();
		}

	/// <summary>
		/// Wrapper interface for PickerResult which adds IDispose to the interface
		/// </summary>
		public static IPickerResult WithComCleanup(this PickerResult resource)
		{
			return resource.WithComCleanup<PickerResult, IPickerResult>();
		}

	/// <summary>
		/// Wrapper interface for PickerResults which adds IDispose to the interface
		/// </summary>
		public static IPickerResults WithComCleanup(this PickerResults resource)
		{
			return resource.WithComCleanup<PickerResults, IPickerResults>();
		}

	/// <summary>
		/// Wrapper interface for PickerDialog which adds IDispose to the interface
		/// </summary>
		public static IPickerDialog WithComCleanup(this PickerDialog resource)
		{
			return resource.WithComCleanup<PickerDialog, IPickerDialog>();
		}

	/// <summary>
		/// Wrapper interface for IMsoContactCard which adds IDispose to the interface
		/// </summary>
		public static IIMsoContactCard WithComCleanup(this IMsoContactCard resource)
		{
			return resource.WithComCleanup<IMsoContactCard, IIMsoContactCard>();
		}

	/// <summary>
		/// Wrapper interface for EffectParameter which adds IDispose to the interface
		/// </summary>
		public static IEffectParameter WithComCleanup(this EffectParameter resource)
		{
			return resource.WithComCleanup<EffectParameter, IEffectParameter>();
		}

	/// <summary>
		/// Wrapper interface for EffectParameters which adds IDispose to the interface
		/// </summary>
		public static IEffectParameters WithComCleanup(this EffectParameters resource)
		{
			return resource.WithComCleanup<EffectParameters, IEffectParameters>();
		}

	/// <summary>
		/// Wrapper interface for PictureEffect which adds IDispose to the interface
		/// </summary>
		public static IPictureEffect WithComCleanup(this PictureEffect resource)
		{
			return resource.WithComCleanup<PictureEffect, IPictureEffect>();
		}

	/// <summary>
		/// Wrapper interface for PictureEffects which adds IDispose to the interface
		/// </summary>
		public static IPictureEffects WithComCleanup(this PictureEffects resource)
		{
			return resource.WithComCleanup<PictureEffects, IPictureEffects>();
		}

	/// <summary>
		/// Wrapper interface for Crop which adds IDispose to the interface
		/// </summary>
		public static ICrop WithComCleanup(this Crop resource)
		{
			return resource.WithComCleanup<Crop, ICrop>();
		}

	/// <summary>
		/// Wrapper interface for ContactCard which adds IDispose to the interface
		/// </summary>
		public static IContactCard WithComCleanup(this ContactCard resource)
		{
			return resource.WithComCleanup<ContactCard, IContactCard>();
		}

	}
}